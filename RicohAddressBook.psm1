#=====================================================================#
#                                                                     #
#  Ricoh Multi Function Printer (MFP) Address Book PowerShell Module  #
#  Original Author: Alexander Krause                                  #
#  Author: Matthew Desjardins                                         #
#  Creation Date: 10.04.2013                                          #
#  Modified Date: 24.05.2021                                          #
#                                                                     #
#=====================================================================#

function ConvertTo-Base64 {
    param(
        [string]
        $String,

        [System.Text.Encoding]
        $Encoding = [System.Text.Encoding]::UTF8
    )

    [System.Convert]::ToBase64String($Encoding.GetBytes($String))
}

$namespaces = @{
    s = 'http://schemas.xmlsoap.org/soap/envelope/'
    u = 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory'
}

function Get-Request {
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [string]
        [Parameter(Mandatory)]
        $Action,

        [string]
        [Parameter(Mandatory)]
        $Message
    )

    if (-not $Hostname.IsAbsoluteUri) {
        $Hostname = 'http://' + $Hostname.OriginalString
    }

    $body = [xml]@"
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
            s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
    <s:Body>
        <u:$Action xmlns:xs="http://www.w3.org/2001/XMLSchema"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                   xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                   xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                   xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
            $Message
        </u:$Action>
    </s:Body>
</s:Envelope>
"@

    @{
        Uri         = [uri]::new($Hostname, 'DH/udirectory')
        Method      = 'Post'
        ContentType = 'text/xml'
        Body        = $body
        Headers     = @{
            SOAPAction =
                "http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#$Action"
        }
    }
}

function Connect-Session {
    [CmdletBinding()]
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [pscredential]
        [Parameter(Mandatory)]
        $Credential,

        [switch]
        $ReadOnly
    )

    $encodedUsername = ConvertTo-Base64 $Credential.UserName
    $encodedPassword = ConvertTo-Base64 $Credential.GetNetworkCredential().Password
    $stringIn = @("SCHEME=QkFTSUM=",               # ConvertTo-Base64 'BASIC'
                  "UID:UserName=$encodedUsername",
                  "PWD:Password=$encodedPassword",
                  "PES:Encoding=") -join ';'
    $lockMode = if ($ReadOnly) {
        'S'
    } else {
        'X'
    }

    $login = Get-Request -Hostname $Hostname -Action startSession -Message @"
        <stringIn>$stringIn</stringIn>
        <timeLimit>30</timeLimit>
        <lockMode>$lockMode</lockMode>
"@

    try {
        Invoke-WebRequest @login |
        Select-Xml -Namespace $namespaces -XPath '/s:Envelope/s:Body/u:startSessionResponse/stringOut/text()' |
        ForEach-Object {$_.Node.Value}
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Search-AddressBookEntry {
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [string]
        [Parameter(Mandatory)]
        $Session
    )

    $offset = 0
    do {
        $search = Get-Request -Hostname $Hostname -Action searchObjects -Message @"
            <sessionId>$Session</sessionId>
            <selectProps xsi:type="soap-enc:Array"
                         soap-enc:arrayType="xs:string[]">
                <item>id</item>
            </selectProps>
            <fromClass>entry</fromClass>
            <orderBy xsi:type="soap-enc:Array"
                     soap-enc:arrayType="ricoh:queryOrderBy[]">
                <item>
                    <propName>index</propName>
                    <isDescending>false</isDescending>
                </item>
            </orderBy>
            <rowOffset>$offset</rowOffset>
            <rowCount>50</rowCount>
"@

        [xml]$response = Invoke-WebRequest @search
        $numberOfResults = $response.Envelope.Body.searchObjectsResponse.numOfResults - 10

        $response |
        Select-Xml -Namespace $namespaces -XPath '/s:Envelope/s:Body/u:searchObjectsResponse/rowList/item/item[propName/text()="id"]/propVal/text()' |
        Select-Object -First ([System.Math]::Min($numberOfResults - $offset, 50)) |
        ForEach-Object {[uint32]$_.Node.Value}

        $offset += 50
        $numberRemaining = $numberOfResults - $offset
    } while ($numberRemaining -gt 0)
}

$letters = @{
    A = 2
    B = 2
    C = 3
    D = 3
    E = 4
    F = 4
    G = 5
    H = 5
    I = 6
    J = 6
    K = 6
    L = 7
    M = 7
    N = 7
    O = 8
    P = 8
    Q = 8
    R = 9
    S = 9
    T = 9
    U = 10
    V = 10
    W = 10
    X = 11
    Y = 11
    Z = 11
}
enum TagId {
    AB = 2
    CD = 3
    EF = 4
    GH = 5
    IJK = 6
    LMN = 7
    OPQ = 8
    RST = 9
    UVW = 10
    XYZ = 11
}

function Add-PropertyList {
    param(
        [System.Xml.XmlNode]
        $Parent,

        [hashtable]
        $Properties
    )

    $document = $Parent.OwnerDocument

    foreach ($pair in $Properties.GetEnumerator()) {
        $item = $document.CreateElement('item')
        $propertyName  = $document.CreateElement('propName')
        $propertyValue = $document.CreateElement('propVal')

        $propertyName.InnerText  = $pair.Key
        $propertyValue.InnerText = $pair.Value

        $item.AppendChild($propertyName)  > $null
        $item.AppendChild($propertyValue) > $null

        $Parent.AppendChild($item) > $null
    }
}

function Test-Property {
    param(
        [hashtable]
        $Properties,

        [string]
        $Name
    )

    $Properties.ContainsKey($Name) -and $Properties[$Name] -eq 'true'
}

<#
.Synopsis
    Retrieves address book entries from a Ricoh multi-function printer

.Description
    Get-AddressBookEntry retrieves address book entries from Ricoh
    multi-function printers.

.Parameter Hostname
    The hostname of the printer from which address book entries are to be
    retrieved.

    By default, it will use HTTP; if HTTPS is required, specify that in the URI,
    like "-Hostname https://printername"

.Parameter Credential
    The username and password to use to connect to the Ricoh printer.

.Parameter Id
    Only retrieve address book entries matching this ID.

.Parameter Name
    Only retrieve address book entries matching this name. Wildcards are
    permitted.

.Inputs
    None. You cannot pipe objects to Get-AddressBookEntry.

.Outputs
    A Ricoh.AddressBook.Entry object containing properties describing an address
    book entry.

.Example
    PS> Get-AddressBookEntry -Hostname https://10.10.10.10 -Credential admin

    Retrieves a list of address book entries on the printer at IP address
    10.10.10.10 using HTTPS. It will prompt for a password.
#>
function Get-AddressBookEntry {
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [pscredential]
        [Parameter(Mandatory)]
        $Credential,

        [uint32[]]
        $Id,

        [string]
        $Name
    )

    try {
        $session = Connect-Session -Hostname $Hostname -Credential $Credential -ReadOnly
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }

    $get = Get-Request -Hostname $Hostname -Action getObjectsProps -Message @"
        <sessionId>$session</sessionId>
        <objectIdList xsi:type="soap-enc:Array"
                      soap-enc:arrayType="xs:string[]" />
        <selectProps xsi:type="soap-enc:Array"
                     soap-enc:arrayType="xs:string[]">
            <item>id</item>
            <item>index</item>
            <item>displayedOrder</item>
            <item>name</item>
            <item>longName</item>
            <item>tagId</item>
            <item>lastAccessDateTime</item>
            <item>isDestination</item>
            <item>isSender</item>
            <item>auth:</item>
            <item>auth:name</item>
            <item>mail:</item>
            <item>mail:address</item>
            <item>mail:parameter</item>
            <item>mail:isDirectSMTP</item>
            <item>remoteFolder:</item>
            <item>remoteFolder:type</item>
            <item>remoteFolder:path</item>
            <item>remoteFolder:port</item>
            <item>remoteFolder:accountName</item>
        </selectProps>
"@

        $objectIdList = $get.Body.Envelope.Body.getObjectsProps.objectIdList
        if ($null -eq $Id) {
            $Id = Search-AddressBookEntry $Hostname $session
        }
        $entries = do {
            $objectIdList.IsEmpty = $true # Empties the node
            $Id |
            Select-Object -First 50 |
            ForEach-Object {
                $item = $get.Body.CreateElement('item')
                $item.InnerText = "entry:$_"
                $objectIdList.AppendChild($item) > $null
            }

            Invoke-WebRequest @get |
            Select-Xml -Namespace $namespaces -XPath '/s:Envelope/s:Body/u:getObjectsPropsResponse/returnValue/item'

            $Id = $Id | Select-Object -Skip 50
        } while ($Id.Count -gt 0)

        foreach ($entry in $entries) {
            $properties = @{}
            foreach ($property in $entry.Node.ChildNodes) {
                if (-not [string]::IsNullOrEmpty($property.propVal)) {
                    $properties[$property.propName] = $property.propVal
                }
            }

            if (-not [string]::IsNullOrEmpty($Name) -and $properties['name'] -notlike $Name) {
                continue
            }

            $output = [ordered]@{
                PSTypeName = 'Ricoh.AddressBook.Entry'

                ID       = [uint32]$properties['id']
                Index    = [uint32]$properties['index']
                Priority = [uint32]$properties['displayedOrder']
                Name     = $properties['name']
                LongName = $properties['longName']
            }

            $tags = $properties['tagId'] -split ','
            $output.Frequent = $tags -contains '1'
            foreach ($tag in $tags) {
                $tag = [int]$tag
                if ($tag -gt 1 -and $tag -le 11) {
                    $output.Title1 = [TagId]$tag
                } elseif ($tag -gt 11 -and $tag -le 21) {
                    $output.Title2 = $tag - 11
                } elseif ($tag -gt 21) {
                    $output.Title3 = $tag - 21
                }
            }

            if ($properties['lastAccessDateTime'] -ne '1970-01-01T00:00:00Z') {
                $output.LastUsed = [datetime]$properties['lastAccessDateTime']
            }

            if (Test-Property $properties 'auth:') {
                $output.UserCode = $properties['auth:name']
            }

            if (Test-Property $properties 'remoteFolder:') {
                $output.RemoteFolderType = $properties['remoteFolder:type']
                $output.RemoteFolderPath = $properties['remoteFolder:path']
                $output.RemoteFolderPort = [uint32]$properties['remoteFolder:port']
                $output.RemoteFolderAccount = $properties['remoteFolder:accountName']
            }

            if (Test-Property $properties 'mail:') {
                $output.EmailAddress  = $properties['mail:address']
                $output.IsSender      = [bool]$properties['isSender']
                $output.IsDestination = [bool]$properties['isDestination']
            }

            [PSCustomObject]$output
        }

    Disconnect-Session $Hostname $session
}

<#
.Synopsis
    Modifies address book entries on a Ricoh multi-function printer

.Description
    Update-AddressBookEntry modifies address book entries on Ricoh
    multi-function printers.

.Parameter Hostname
    The hostname of the printer from which address book entries are to be
    modified.

    By default, it will use HTTP; if HTTPS is required, specify that in the URI,
    like "-Hostname https://printername"

.Parameter Credential
    The username and password to use to connect to the Ricoh printer.

.Parameter Id
    The ID of the address book entry to modify. Find the ID from
    Get-AddressBookEntry.

.Parameter Name
    The new name for the address book entry.

.Parameter LongName
    The new "long name" for the address book entry.

.Parameter ScanAccount
    The account to use to save the scanned files to a network location.

.Parameters Path
    The network path used to save scanned files.

.Example
    PS> Update-AddressBookEntry -Hostname https://10.10.10.10 -Credential admin -Id 1 -Name 'Matthew D'

    Sets the name of the address book entry with ID of 1 to "Matthew D".

.Example
    PS> $entries = @(
        [PSCustomObject]@{
            Id = 1
            Name = 'Matthew D'
        }
        [PSCustomObject]@{
            Id = 2
            Name = 'John D'
        }
    )

    PS> $entries | Update-AddressBookEntry -Hostname https://10.10.10.10 -Credential admin

    Updates multiple entries based on objects received in the pipeline.
#>
function Update-AddressBookEntry {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [pscredential]
        [Parameter(Mandatory)]
        $Credential,

        [int]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $Id,

        [string]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Name,

        [string]
        [Parameter(ValueFromPipelineByPropertyName)]
        $LongName,

        [pscredential]
        [Parameter(ValueFromPipelineByPropertyName)]
        $ScanAccount,

        [string]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Path
    )

    begin {
        try {
            $session = Connect-Session -Hostname $Hostname -Credential $Credential
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }

    process {
        $set = Get-Request -Hostname $Hostname -Action putObjectProps -Message @"
            <sessionId>$session</sessionId>
            <objectId>entry:$Id</objectId>
            <propList xsi:type="soap-enc:Array"
                      soap-enc:arrayType="ricoh:property[]" />
"@

        $properties = @{}

        if ($PSBoundParameters.ContainsKey('ScanAccount')) {
            $properties['remoteFolder:accountName'] = $ScanAccount.UserName
            $properties['remoteFolder:password'] = ConvertTo-Base64 $ScanAccount.GetNetworkCredential().Password
        }
        if ($PSBoundParameters.ContainsKey('Path')) {
            $properties['remoteFolder:path'] = $Path
        }
        if ($PSBoundParameters.ContainsKey('Name')) {
            $properties['name'] = $Name
        }
        if ($PSBoundParameters.ContainsKey('LongName')) {
            $properties['longName'] = $LongName
        }

        Add-PropertyList $set.Body.Envelope.Body.putObjectProps.propList $properties

        if ($PSCmdlet.ShouldProcess(
                "Updating address book entry with ID of $Id.",
                "Update address book entry with ID of ${Id}?",
                "Confirm address book update.")) {
            Invoke-WebRequest @set > $null
        }
    }

    end {
        Disconnect-Session $Hostname $session
    }
}

<#
.Synopsis
    Adds address book entries to a Ricoh multi-function printer

.Description
    Add-AddressBookEntry adds address book entries to Ricoh
    multi-function printers.

.Parameter Hostname
    The hostname of the printer from which address book entries are to be
    added.

    By default, it will use HTTP; if HTTPS is required, specify that in the URI,
    like "-Hostname https://printername"

.Parameter Credential
    The username and password to use to connect to the Ricoh printer.

.Parameter Name
    The name for the address book entry.

.Parameter LongName
    The "long name" for the address book entry.

.Parameter ScanAccount
    The account to use to save the scanned files to a network location.

.Parameters Path
    The network path used to save scanned files.

.Example
    PS> $entry = @{
        Hostname = 'https://10.10.10.10'
        Credential = Get-Credential admin
        Name = 'Matthew D'
        LongName = 'Matthew Desjardins'
        ScanAccount = Get-Credential ScanAccount
        FolderPath = '\\my\path\here'
    }
    PS> Add-AddressBookEntry @entry

.Example
    PS> $scanAccount = Get-Credential ScanAccount
    PS> $entries = @(
        [PSCustomObject]@{
            Name = 'Matthew D'
            LongName = 'Matthew Desjardins'
            ScanAccount = $scanAccount
            FolderPath = '\\my\path\here'
        }
        [PSCustomObject]@{
            Name = 'John D'
            LongName = 'John Doe'
            ScanAccount = $scanAccount
            FolderPath = '\\my\path\here'
        }
    )
    PS> $entries | Add-AddressBookEntry -Hostname https://10.10.10.10 -Credential admin

    Adds multiple entries based on objects received in the pipeline.
#>
function Add-AddressBookEntry {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [pscredential]
        [Parameter(Mandatory)]
        $Credential,

        [string]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $Name,

        [string]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $LongName,

        [pscredential]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $ScanAccount,

        [string]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $FolderPath
    )

    begin {
        try {
            $session = Connect-Session -Hostname $Hostname -Credential $Credential
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }

        $add = Get-Request -Hostname $Hostname -Action putObjects -Message @"
            <sessionId>$session</sessionId>
            <objectClass>entry</objectClass>
            <propListList xsi:type="soap-enc:Array"
                          soap-enc:arrayType="ricoh:propertyList[]" />
"@
    }

    process {
        $tagId = $letters[($Name[0].ToString())]

        $entry = $add.Body.CreateElement('item')
        $add.Body.Envelope.Body.putObjects.propListList.AppendChild($entry) > $null

        Add-PropertyList $entry @{
            'entryType'                = 'user'
            'name'                     = $Name
            'longName'                 = $LongName
            'remoteFolder:path'        = $FolderPath
            'remoteFolder:accountName' = $ScanAccount.UserName
            'remoteFolder:password'    = ConvertTo-Base64 $ScanAccount.GetNetworkCredential().Password
            'remoteFolder:port'        = 21
            'remoteFolder:select'      = 'private'
            'tagId'                    = "1,$tagId"
            'isDestination'            = 'true'
        }
    }

    end {
        $names = Select-Xml -Xml $add.Body -Namespace $namespaces -XPath '/s:Envelope/s:Body/u:putObjects/propListList/item/item[propName/text()="name"]/propVal'

        $allNames = $names -join ', '
        if ($PSCmdlet.ShouldProcess(
                "Adding address book entries for $allNames",
                "Add address book entry for ${allNames}?",
                'Confirm address book addition.')) {
            Invoke-WebRequest @add > $null
        }
        Disconnect-Session $Hostname $session
    }
}

<#
.Synopsis
    Removes address book entries from a Ricoh multi-function printer

.Description
    Remove-AddressBookEntry removes address book entries from Ricoh
    multi-function printers.

.Parameter Hostname
    The hostname of the printer from which address book entries are to be
    removed.

    By default, it will use HTTP; if HTTPS is required, specify that in the URI,
    like "-Hostname https://printername"

.Parameter Credential
    The username and password to use to connect to the Ricoh printer.

.Parameter Id
    The IDs to be removed. Find the IDs from Get-AddressBookEntry.

.Example
    PS> Remove-AddressBookEntry -Hostname https://10.10.10.10 -Credential admin -Id 1, 2

.Example
    PS> 1, 2 | Remove-AddressBookEntry -Hostname https://10.10.10.10 -Credential admin

.Example
    PS> $users = Get-AddressBookEntry -Hostname https://10.10.10.10 -Credential admin |
                 Where-Object {[string]::IsNullOrEmpty($_.LongName)}
    PS> $users | Remove-AddressBookEntry -Hostname https://10.10.10.10 -Credential admin

    These need to be run as separate commands, as the second command cannot run
    while the first command has the Ricoh address book open.
#>
function Remove-AddressBookEntry {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [pscredential]
        [Parameter(Mandatory)]
        $Credential,

        [int[]]
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        $Id
    )

    begin {
        try {
            $session = Connect-Session -Hostname $Hostname -Credential $Credential
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }

        $remove = Get-Request -Hostname $Hostname -Action deleteObjects -Message @"
            <sessionId>$session</sessionId>
            <objectIdList xsi:type="soap-enc:Array"
                        soap-enc:arrayType="xs:string[]" />
"@
        $objectIdList = $remove.Body.Envelope.Body.deleteObjects.objectIdList
    }

    process {
        foreach ($item in $Id) {
            $element = $remove.Body.CreateElement('item')
            $element.InnerText = "entry:$item"
            $objectIdList.AppendChild($element) > $null
        }
    }

    end {
        $entries =
            Select-Xml -Xml $remove.Body -Namespace $namespaces -XPath '/s:Envelope/s:Body/u:deleteObjects/objectIdList/item' |
            ForEach-Object {
                if ($_.Node.InnerText -match '^entry:(\d+)$') {
                    $Matches[1]
                }
            }
        $allID = $entries -join ', '
        if ($PSCmdlet.ShouldProcess(
                "Removing IDs $allID.",
                "Remove IDs ${allID}?",
                'Confirm removing entries')) {
            Invoke-WebRequest @remove > $null
        }

        Disconnect-Session $Hostname $session
    }
}

function Disconnect-Session {
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [string]
        [Parameter(Mandatory)]
        $Session
    )

    $logout = Get-Request -Hostname $Hostname -Action terminateSession -Message "<sessionId>$Session</sessionId>"

    try {
        Invoke-WebRequest @logout > $null
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
