#=====================================================================#
#                                                                     #
#  Ricoh Multi Function Printer (MFP) Address Book PowerShell Module  #
#  Created By: Alexander Krause                                       #
#  Mantainer: Matthew Desjardins                                      #
#                                                                     #
#=====================================================================#

enum RicohMethodType {
    startSession
    terminateSession
    searchObjects
    getObjectsProps
    putObjectProps
    putObjects
    deleteObjects
}

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

function Invoke-SOAPRequest {
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [xml]
        [Parameter(Mandatory)]
        $Body,

        [RicohMethodType]
        [Parameter(Mandatory)]
        $Method,

        [switch]
        $SkipCertificateCheck
    )

    if (-not $Hostname.IsAbsoluteUri) {
        $Hostname = 'http://' + $Hostname.OriginalString
    }

    $webRequest = @{
        Uri         = [uri]::new($Hostname, 'DH/udirectory')
        Method      = [Microsoft.PowerShell.Commands.WebRequestMethod]::Post
        ContentType = 'text/xml'
        Body        = $Body
        Headers     = @{
            SOAPAction =
                "http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#$Method"
        }
    }

    try {
        $response =
            if (-not $SkipCertificateCheck) {
                Invoke-WebRequest @webRequest
            } elseif ((Get-Command Invoke-WebRequest).Parameters.ContainsKey('SkipCertificateCheck')) {
                Invoke-WebRequest -SkipCertificateCheck @webRequest
            } else {
                $originalCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {
                    $true
                }
                Invoke-WebRequest @webRequest
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $originalCallback
            }
        [xml]$response
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-Template {
    param(
        [RicohMethodType]
        [Parameter(Mandatory)]
        $Method
    )

    $content = Get-Content -Path "$PSScriptRoot\Templates\$Method.xml"
    [xml]$content
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
        $ReadOnly,

        [switch]
        $SkipCertificateCheck
    )

    $method = [RicohMethodType]::startSession
    $template = Get-Template $method

    $scheme = ConvertTo-Base64 BASIC
    $username = ConvertTo-Base64 $Credential.UserName
    $password = ConvertTo-Base64 $Credential.GetNetworkCredential().Password
    $template.Envelope.Body.$method.stringIn =
        "SCHEME=$scheme;UID:UserName=$username;PWD:Password=$password;PES:Encoding="
    $template.Envelope.Body.$method.lockMode =
        if ($ReadOnly) {
            'S'
        } else {
            'X'
        }

    $request = @{
        Hostname = $Hostname
        Body     = $template
        Method   = $method
        SkipCertificateCheck = $SkipCertificateCheck
    }

    try {
        $response = Invoke-SOAPRequest @request
        $response.Envelope.Body.startSessionResponse.stringOut
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
        $Session,

        [switch]
        $SkipCertificateCheck
    )

    $method = [RicohMethodType]::searchObjects
    $template = Get-Template $method
    $template.Envelope.Body.$method.sessionId = $Session

    $offset = 0
    do {
        $template.Envelope.Body.$method.rowOffset = [string]$offset

        $request = @{
            Hostname = $Hostname
            Body     = $template
            Method   = $method
            SkipCertificateCheck = $SkipCertificateCheck
        }

        $response = Invoke-SOAPRequest @request
        $totalResults = $response.Envelope.Body.searchObjectsResponse.numOfResults

        $selection = @{
            Xml = $response
            Namespace = $namespaces
            XPath =
                '/s:Envelope/s:Body/u:searchObjectsResponse/rowList/item/item[propName/text()="id"]/propVal/text()'
        }

        Select-Xml @selection |
        Select-Object -First ([System.Math]::Min($totalResults - $offset, 50)) |
        ForEach-Object {
            $id = $_.Node.Value
            if ($id.Length -lt 10) {
                [uint32]$id
            }
        }

        $offset += 50
    } while ($response.Envelope.Body.searchObjectsResponse.returnValue -ne 'EOD')
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

function Get-Title1Tag {
    param(
        [char]
        [ValidatePattern('^[A-Za-z]$')]
        $Letter
    )

    switch ([char]::ToUpper($Letter)) {
        A { [TagId]::AB }
        B { [TagId]::AB }
        C { [TagId]::CD }
        D { [TagId]::CD }
        E { [TagId]::EF }
        F { [TagId]::EF }
        G { [TagId]::GH }
        H { [TagId]::GH }
        I { [TagId]::IJK }
        J { [TagId]::IJK }
        K { [TagId]::IJK }
        L { [TagId]::LMN }
        M { [TagId]::LMN }
        N { [TagId]::LMN }
        O { [TagId]::OPQ }
        P { [TagId]::OPQ }
        Q { [TagId]::OPQ }
        R { [TagId]::RST }
        S { [TagId]::RST }
        T { [TagId]::RST }
        U { [TagId]::UVW }
        V { [TagId]::UVW }
        W { [TagId]::UVW }
        X { [TagId]::XYZ }
        Y { [TagId]::XYZ }
        Z { [TagId]::XYZ }
    }
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

.Parameter SkipCertificateCheck
    Skips certificate validation checks. This includes all validation such as expiration,
    revocation, trusted root authority, etc.

    > [!WARNING] Using this parameter is not secure and is not recommended. This switch is only intended to be
    used against known hosts using a self-signed certificate for testing purposes. Use at your own risk.

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
        $Name,

        [switch]
        $SkipCertificateCheck
    )

    try {
        $connection = @{
            Hostname   = $Hostname
            Credential = $Credential
            SkipCertificateCheck = $SkipCertificateCheck
        }

        $session = Connect-Session @connection -ReadOnly
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }

    $method = [RicohMethodType]::getObjectsProps
    $template = Get-Template $method
    $template.Envelope.Body.$method.sessionId = $session

    $objectIdList = $template.Envelope.Body.getObjectsProps.objectIdList
    if ($null -eq $Id) {
        $selection = @{
            Hostname = $Hostname
            Session  = $session
            SkipCertificateCheck = $SkipCertificateCheck
        }
        $Id = Search-AddressBookEntry @selection
    }
    $entries = do {
        $objectIdList.IsEmpty = $true # Empties the node
        $Id |
        Select-Object -First 50 |
        ForEach-Object {
            $item = $template.CreateElement('item')
            $item.InnerText = "entry:$_"
            $objectIdList.AppendChild($item) > $null
        }

        $request = @{
            Hostname = $Hostname
            Body     = $template
            Method   = $method
            SkipCertificateCheck = $SkipCertificateCheck
        }

        Invoke-SOAPRequest @request |
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

        $output.Frequent = $false
        switch ($properties['tagId'] -split ',') {
            1 {
                $output.Frequent = $true
            }
            {2..11 -contains $_} {
                $output.Title1 = [TagId]$_
            }
            {12..21 -contains $_} {
                $output.Title2 = $_ - 11
            }
            default {
                $output.Title3 = $_ - 21
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

    Disconnect-Session -Hostname $Hostname -Session $session -SkipCertificateCheck:$SkipCertificateCheck
}

function Get-TagIdValue {
    param(
        [hashtable]
        $Parameters
    )

    $tags = [System.Collections.Generic.List[byte]]::new(4)
    if ($Parameters.ContainsKey('Frequent')) {
        $tags.Add([byte]$Parameters.Frequent.IsPresent)
    }
    if ($Parameters.ContainsKey('Title1')) {
        $tags.Add([byte]$Title1)
    }
    if ($Parameters.ContainsKey('Title2')) {
        $tags.Add($Title2 + 11)
    }
    if ($Parameters.ContainsKey('Title3')) {
        $tags.Add($Title3 + 21)
    }
    $tags -join ','
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

.Parameters FolderPath
    The network path used to save scanned files.

.Parameters DisplayPriority
    The display order of the user in address book list. Sorting is done first by
    DisplayPriority, then by ID.

.Parameters Frequent
    Whether the user is to be added to the frequently used list.

.Parameters Title1
    The heading to list this user under for Title1.

    Title1 is usually the default grouping, and is the one that lists users by
    letters (AB, CD, etc.) on the scanner.

.Parameters Title2
    The heading to list this user under for Title2.

    Title2 is a range from 1 to 10, and is another option for grouping users on
    the scanner.

.Parameters Title3
    The heading to list this user under for Title3.

    Title3 is a range from 1 to 5, and is another option for grouping users on
    the scanner.

.Parameter SkipCertificateCheck
    Skips certificate validation checks. This includes all validation such as expiration,
    revocation, trusted root authority, etc.

    > [!WARNING] Using this parameter is not secure and is not recommended. This switch is only intended to be
    used against known hosts using a self-signed certificate for testing purposes. Use at your own risk.

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
        $FolderPath,

        [byte]
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(1, 10)]
        $DisplayPriority,

        [switch]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Frequent,

        [TagId]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Title1,

        [byte]
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(1, 10)]
        $Title2,

        [byte]
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(1, 5)]
        $Title3,

        [switch]
        $SkipCertificateCheck
    )

    begin {
        try {
            $connection = @{
                Hostname   = $Hostname
                Credential = $Credential
                SkipCertificateCheck = $SkipCertificateCheck
            }

            $session = Connect-Session @connection
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }

        $method = [RicohMethodType]::putObjectProps
        $template = Get-Template $method
        $template.Envelope.Body.$method.sessionId = $session
    }

    process {
        $template.Envelope.Body.$method.objectId = "entry:$Id"

        $properties = @{}

        if ($PSBoundParameters.ContainsKey('ScanAccount')) {
            $properties['remoteFolder:accountName'] = $ScanAccount.UserName
            $properties['remoteFolder:password'] = ConvertTo-Base64 $ScanAccount.GetNetworkCredential().Password
        }
        if ($PSBoundParameters.ContainsKey('FolderPath')) {
            $properties['remoteFolder:path'] = $FolderPath
        }
        if ($PSBoundParameters.ContainsKey('Name')) {
            $properties['name'] = $Name
        }
        if ($PSBoundParameters.ContainsKey('LongName')) {
            $properties['longName'] = $LongName
        }
        if ($PSBoundParameters.ContainsKey('DisplayPriority')) {
            $properties['displayedOrder'] = $DisplayPriority
        }

        # Tags (Frequent, Title1, Title2, Title3)
        $tags = Get-TagIdValue $PSBoundParameters
        if (-not [string]::IsNullOrEmpty($tags)) {
            $properties['tagId'] = $tags
        }

        Add-PropertyList $template.Envelope.Body.putObjectProps.propList $properties

        if ($PSCmdlet.ShouldProcess(
                "Updating address book entry with ID of $Id.",
                "Update address book entry with ID of ${Id}?",
                "Confirm address book update.")) {
            $request = @{
                Hostname = $Hostname
                Body     = $template
                Method   = $method
                SkipCertificateCheck = $SkipCertificateCheck
            }

            Invoke-SOAPRequest @request > $null
        }
    }

    end {
        Disconnect-Session -Hostname $Hostname -Session $session -SkipCertificateCheck:$SkipCertificateCheck
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

.Parameters FolderPath
    The network path used to save scanned files.

.Parameters Frequent
    Whether the user is to be added to the frequently used list.

.Parameters Title1
    The heading to list this user under for Title1.

    Title1 is usually the default grouping, and is the one that lists users by
    letters (AB, CD, etc.) on the scanner.

.Parameters Title2
    The heading to list this user under for Title2.

    Title2 is a range from 1 to 10, and is another option for grouping users on
    the scanner.

.Parameters Title3
    The heading to list this user under for Title3.

    Title3 is a range from 1 to 5, and is another option for grouping users on
    the scanner.

.Parameter SkipCertificateCheck
    Skips certificate validation checks. This includes all validation such as expiration,
    revocation, trusted root authority, etc.

    > [!WARNING] Using this parameter is not secure and is not recommended. This switch is only intended to be
    used against known hosts using a self-signed certificate for testing purposes. Use at your own risk.

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
        $FolderPath,

        [switch]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Frequent,

        [TagId]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Title1,

        [byte]
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(1, 10)]
        $Title2,

        [byte]
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(1, 5)]
        $Title3,

        [switch]
        $SkipCertificateCheck
    )

    begin {
        try {
            $connection = @{
                Hostname   = $Hostname
                Credential = $Credential
                SkipCertificateCheck = $SkipCertificateCheck
            }

            $session = Connect-Session @connection
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        $method = [RicohMethodType]::putObjects
        $template = Get-Template $method
        $template.Envelope.Body.$method.sessionId = $session
    }

    process {
        # Tags (Frequent, Title1, Title2, Title3)
        $tagId = Get-TagIdValue $PSBoundParameters

        $entry = $template.CreateElement('item')
        $template.Envelope.Body.$method.propListList.AppendChild($entry) > $null

        Add-PropertyList $entry @{
            'entryType'                = 'user'
            'name'                     = $Name
            'longName'                 = $LongName
            'remoteFolder:path'        = $FolderPath
            'remoteFolder:accountName' = $ScanAccount.UserName
            'remoteFolder:password'    = ConvertTo-Base64 $ScanAccount.GetNetworkCredential().Password
            'remoteFolder:port'        = 21
            'remoteFolder:select'      = 'private'
            'tagId'                    = $tagId
            'isDestination'            = 'true'
        }
    }

    end {
        $selection = @{
            Xml       = $template
            Namespace = $namespaces
            XPath     = '/s:Envelope/s:Body/u:putObjects/propListList/item/item[propName/text()="name"]/propVal'
        }
        $names = Select-Xml @selection

        $allNames = $names -join ', '
        if ($PSCmdlet.ShouldProcess(
                "Adding address book entries for $allNames",
                "Add address book entry for ${allNames}?",
                'Confirm address book addition.')) {
            $request = @{
                Hostname = $Hostname
                Body     = $template
                Method   = $method
                SkipCertificateCheck = $SkipCertificateCheck
            }

            Invoke-SOAPRequest @request > $null
        }
        Disconnect-Session -Hostname $Hostname -Session $session -SkipCertificateCheck:$SkipCertificateCheck
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

.Parameter SkipCertificateCheck
    Skips certificate validation checks. This includes all validation such as expiration,
    revocation, trusted root authority, etc.

    > [!WARNING] Using this parameter is not secure and is not recommended. This switch is only intended to be
    used against known hosts using a self-signed certificate for testing purposes. Use at your own risk.

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
        $Id,

        [switch]
        $SkipCertificateCheck
    )

    begin {
        try {
            $connection = @{
                Hostname   = $Hostname
                Credential = $Credential
                SkipCertificateCheck = $SkipCertificateCheck
            }

            $session = Connect-Session @connection
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }

        $method = [RicohMethodType]::deleteObjects
        $template = Get-Template $method
        $template.Envelope.Body.$method.sessionId = $session

        $objectIdList = $template.Envelope.Body.deleteObjects.objectIdList
    }

    process {
        foreach ($item in $Id) {
            $element = $template.CreateElement('item')
            $element.InnerText = "entry:$item"
            $objectIdList.AppendChild($element) > $null
        }
    }

    end {
        $entries =
            Select-Xml -Xml $template -Namespace $namespaces -XPath '/s:Envelope/s:Body/u:deleteObjects/objectIdList/item' |
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
            $request = @{
                Hostname = $Hostname
                Body     = $template
                Method   = $method
                SkipCertificateCheck = $SkipCertificateCheck
            }

            Invoke-SOAPRequest @request > $null
        }

        Disconnect-Session -Hostname $Hostname -Session $session -SkipCertificateCheck:$SkipCertificateCheck
    }
}

function Disconnect-Session {
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [string]
        [Parameter(Mandatory)]
        $Session,

        [switch]
        $SkipCertificateCheck
    )

    $method = [RicohMethodType]::terminateSession
    $template = Get-Template $method
    $template.Envelope.Body.$method.sessionId = $session

    $request = @{
        Hostname = $Hostname
        Body     = $template
        Method   = $method
        SkipCertificateCheck = $SkipCertificateCheck
    }

    try {
        Invoke-SOAPRequest @request > $null
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
