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
    [OutputType([xml])]
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
            SOAPAction = "http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#$Method"
        }
    }

    try {
        $response = if (-not $SkipCertificateCheck) {
            Invoke-WebRequest @webRequest
        } elseif ((Get-Command Invoke-WebRequest).Parameters.ContainsKey('SkipCertificateCheck')) {
            Invoke-WebRequest -SkipCertificateCheck @webRequest
        } else {
            try {
                $originalCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {
                    $true
                }
                Invoke-WebRequest @webRequest
            } finally {
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $originalCallback
            }
        }

        [xml]$response
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-Template {
    [OutputType([xml])]
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
    [xml] $template = Get-Template $method
    $content = $template.Envelope.Body.$method

    $scheme = ConvertTo-Base64 BASIC
    $username = ConvertTo-Base64 $Credential.UserName
    $password = ConvertTo-Base64 $Credential.GetNetworkCredential().Password
    $content.stringIn = "SCHEME=$scheme;UID:UserName=$username;PWD:Password=$password;PES:Encoding="
    $content.lockMode = if ($ReadOnly) {
        'S'
    } else {
        'X'
    }

    $request = @{
        Hostname             = $Hostname
        Body                 = $template
        Method               = $method
        SkipCertificateCheck = $SkipCertificateCheck
    }

    try {
        [xml] $response = Invoke-SOAPRequest @request
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
    [xml] $template = Get-Template $method
    $template.Envelope.Body.$method.sessionId = $Session

    $offset = 0
    # This function operates on batches of 50, as scanners that have more than
    # that in their address book will only return the first 50.
    do {
        [xml] $message = $template.Clone()
        $message.Envelope.Body.$method.rowOffset = [string]$offset

        $request = @{
            Hostname             = $Hostname
            Body                 = $message
            Method               = $method
            SkipCertificateCheck = $SkipCertificateCheck
        }

        [xml] $response = Invoke-SOAPRequest @request
        $responseBody = $response.Envelope.Body.searchObjectsResponse
        $totalResults = $responseBody.numOfResults

        $selection = @{
            Xml       = $response
            Namespace = $namespaces
            XPath     = '/s:Envelope/s:Body/u:searchObjectsResponse/rowList/item/item[propName/text()="id"]/propVal/text()'
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
    } while ($responseBody.returnValue -ne 'EOD')
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
    [CmdletBinding()]
    [OutputType([TagId])]
    param(
        [char]
        [ValidatePattern('^[A-Za-z]$')]
        [Parameter(Mandatory)]
        $Letter
    )

    switch ([char]::ToUpper($Letter)) {
        { 'A', 'B' -contains $_ } { [TagId]::AB; break }
        { 'C', 'D' -contains $_ } { [TagId]::CD; break }
        { 'E', 'F' -contains $_ } { [TagId]::EF; break }
        { 'G', 'H' -contains $_ } { [TagId]::GH; break }
        { 'I', 'J', 'K' -contains $_ } { [TagId]::IJK; break }
        { 'L', 'M', 'N' -contains $_ } { [TagId]::LMN; break }
        { 'O', 'P', 'Q' -contains $_ } { [TagId]::OPQ; break }
        { 'R', 'S', 'T' -contains $_ } { [TagId]::RST; break }
        { 'U', 'V', 'W' -contains $_ } { [TagId]::UVW; break }
        { 'X', 'Y', 'Z' -contains $_ } { [TagId]::XYZ; break }
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

    To reuse the credentials between commands, you can use the Get-Credential
    cmdlet and store the results to a variable. Otherwise, just passing a string
    will open a dialog box for the user to enter a password. See the help for
    Get-Credential for information on using this in a script non-interactively.

.Parameter Id
    Only retrieve address book entries matching this ID.

.Parameter Name
    Only retrieve address book entries matching this name. Wildcards are
    permitted.

.Parameter SkipCertificateCheck
    Skips certificate validation checks. This includes all validation such as
    expiration, revocation, trusted root authority, etc.

    > [!WARNING] Using this parameter is not secure and is not recommended. This
    switch is only intended to be used against known hosts using a self-signed
    certificate for testing purposes. Use at your own risk.

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
    [OutputType('Ricoh.AddressBook.Entry')]
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
        [SupportsWildcards()]
        $Name,

        [switch]
        $SkipCertificateCheck
    )

    try {
        $connection = @{
            Hostname             = $Hostname
            Credential           = $Credential
            SkipCertificateCheck = $SkipCertificateCheck
        }

        $session = Connect-Session @connection -ReadOnly
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }

    $method = [RicohMethodType]::getObjectsProps
    [xml] $template = Get-Template $method
    $template.Envelope.Body.$method.sessionId = $session

    if ($null -eq $Id) {
        $selection = @{
            Hostname             = $Hostname
            Session              = $session
            SkipCertificateCheck = $SkipCertificateCheck
        }
        $Id = Search-AddressBookEntry @selection
    }
    # This function operates on batches of 50, as scanners that have more than
    # that in their address book will return an error if more than 50 entries
    # are requested at once.
    $entries = do {
        [xml] $message = $template.Clone()
        $Id |
            Select-Object -First 50 |
            ForEach-Object {
                $item = $message.CreateElement('item')
                $item.InnerText = "entry:$_"
                $message.Envelope.Body.$method.objectIdList.AppendChild($item) > $null
            }

        $request = @{
            Hostname             = $Hostname
            Body                 = $message
            Method               = $method
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

            ID         = [uint32]$properties['id']
            Index      = [uint32]$properties['index']
            Priority   = [uint32]$properties['displayedOrder']
            Name       = $properties['name']
            LongName   = $properties['longName']
        }

        $output.Frequent = $false
        switch ($properties['tagId'] -split ',') {
            1 {
                $output.Frequent = $true
            }
            { 2..11 -contains $_ } {
                $output.Title1 = [TagId]$_
            }
            { 12..21 -contains $_ } {
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
            $output.EmailAddress = $properties['mail:address']
            $output.IsSender = $properties['isSender'] -eq 'true'
        }

        $output.IsDestination = $properties['isDestination'] -eq 'true'

        [PSCustomObject]$output
    }

    Disconnect-Session -Hostname $Hostname -Session $session -SkipCertificateCheck:$SkipCertificateCheck
}

function Add-Property {
    param(
        [System.Xml.XmlNode]
        $Parent,

        [string]
        $Key,

        [object]
        $Value
    )

    $document = $Parent.OwnerDocument
    $item = $document.CreateElement('item')
    $propertyName = $document.CreateElement('propName')
    $propertyValue = $document.CreateElement('propVal')

    $propertyName.InnerText = $Key
    $propertyValue.InnerText = $Value

    $item.AppendChild($propertyName)  > $null
    $item.AppendChild($propertyValue) > $null
    $Parent.AppendChild($item)        > $null
}

function Get-TagIdValue {
    $tags = [System.Collections.Generic.List[byte]]::new(4)
    if ($Frequent) {
        $tags.Add(1)
    }
    if ($null -ne $Title1) {
        $tags.Add([byte]$Title1)
    }
    if (0 -ne $Title2) {
        $tags.Add($Title2 + 11)
    }
    if (0 -ne $Title3) {
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

    To reuse the credentials between commands, you can use the Get-Credential
    cmdlet and store the results to a variable. Otherwise, just passing a string
    will open a dialog box for the user to enter a password. See the help for
    Get-Credential for information on using this in a script non-interactively.

.Parameter Id
    The ID of the address book entry to modify. Find the ID from
    Get-AddressBookEntry. ID is **not** the Index.

.Parameter Name
    The new name for the address book entry.

.Parameter LongName
    The new "long name" for the address book entry.

.Parameter FolderPath
    The network path used to save scanned files.

.Parameter ScanAccount
    The account to use to save the scanned files to a network location.

    To reuse the credentials between commands, you can use the Get-Credential
    cmdlet and store the results to a variable. Otherwise, just passing a string
    will open a dialog box for the user to enter a password. See the help for
    Get-Credential for information on using this in a script non-interactively.

.Parameter EmailAddress
    The email address used to send scanned files.

.Parameter IsSender
    Whether the given email address is registered as a sender.

.Parameter IsDestination
    Whether the given email address is registered as a destination.

.Parameter DisplayPriority
    The display order of the user in address book list. Sorting is done first by
    DisplayPriority, then by ID.

.Parameter Frequent
    Whether the user is to be added to the frequently used list.

    Frequent, Title1, Title2, and Title3 are all stored in the scanner as
    attributes of a single property, and updating any single value will reset
    the value of the other properties unless they are also provided. In
    addition, at least one value must be set in order for the entry to be
    visible on the scanner.

.Parameter Title1
    The heading to list this user under for Title1.

    Title1 is usually the default grouping, and is the one that lists users by
    letters (AB, CD, etc.) on the scanner.

    Frequent, Title1, Title2, and Title3 are all stored in the scanner as
    attributes of a single property, and updating any single value will reset
    the value of the other properties unless they are also provided. In
    addition, at least one value must be set in order for the entry to be
    visible on the scanner.

.Parameter Title2
    The heading to list this user under for Title2.

    Title2 is a range from 1 to 10, and is another option for grouping users on
    the scanner.

    Frequent, Title1, Title2, and Title3 are all stored in the scanner as
    attributes of a single property, and updating any single value will reset
    the value of the other properties unless they are also provided. In
    addition, at least one value must be set in order for the entry to be
    visible on the scanner.

.Parameter Title3
    The heading to list this user under for Title3.

    Title3 is a range from 1 to 5, and is another option for grouping users on
    the scanner.

    Frequent, Title1, Title2, and Title3 are all stored in the scanner as
    attributes of a single property, and updating any single value will reset
    the value of the other properties unless they are also provided. In
    addition, at least one value must be set in order for the entry to be
    visible on the scanner.

.Parameter SkipCertificateCheck
    Skips certificate validation checks. This includes all validation such as
    expiration, revocation, trusted root authority, etc.

    > [!WARNING] Using this parameter is not secure and is not recommended. This
    switch is only intended to be used against known hosts using a self-signed
    certificate for testing purposes. Use at your own risk.

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
    [CmdletBinding(SupportsShouldProcess, PositionalBinding = $false)]
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

        [string]
        [Parameter(ValueFromPipelineByPropertyName)]
        $FolderPath,

        [pscredential]
        [Parameter(ValueFromPipelineByPropertyName)]
        $ScanAccount,

        [string]
        [Parameter(ValueFromPipelineByPropertyName)]
        $EmailAddress,

        [nullable[bool]]
        [Parameter(ValueFromPipelineByPropertyName)]
        $IsSender,

        [nullable[bool]]
        [Parameter(ValueFromPipelineByPropertyName)]
        $IsDestination,

        [byte]
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(1, 10)]
        $DisplayPriority,

        [nullable[bool]]
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
                Hostname             = $Hostname
                Credential           = $Credential
                SkipCertificateCheck = $SkipCertificateCheck
            }

            $session = Connect-Session @connection
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }

        $method = [RicohMethodType]::putObjectProps
        [xml] $template = Get-Template $method
        $template.Envelope.Body.$method.sessionId = $session
    }

    process {
        [xml] $message = $template.Clone()
        $content = $message.Envelope.Body.$method
        $content.objectId = "entry:$Id"

        function add($key, $value) {
            Add-Property $content.propList $key $value
        }

        if (-not [string]::IsNullOrEmpty($Name)) {
            add 'name' $Name
        }
        if (-not [string]::IsNullOrEmpty($LongName)) {
            add 'longName' $LongName
        }

        if (-not [string]::IsNullOrEmpty($FolderPath)) {
            add 'remoteFolder:path' $FolderPath
        }
        if ($null -ne $ScanAccount) {
            add 'remoteFolder:accountName' $ScanAccount.UserName
            add 'remoteFolder:password' (ConvertTo-Base64 $ScanAccount.GetNetworkCredential().Password)
        }

        if (-not [string]::IsNullOrEmpty($EmailAddress)) {
            add 'mail:address' $EmailAddress
        }
        if ($null -ne $IsSender) {
            add 'isSender' $IsSender.ToString().ToLower()
        }
        if ($null -ne $IsDestination) {
            add 'isDestination' $IsDestination.ToString().ToLower()
        }

        if (0 -ne $DisplayPriority) {
            add 'displayedOrder' $DisplayPriority
        }

        # Tags (Frequent, Title1, Title2, Title3)
        $tags = Get-TagIdValue
        if (-not [string]::IsNullOrEmpty($tags)) {
            add 'tagId' $tags
        }

        if ($PSCmdlet.ShouldProcess(
                "Updating address book entry with ID of $Id.",
                "Update address book entry with ID of ${Id}?",
                'Confirm address book update.')) {
            $request = @{
                Hostname             = $Hostname
                Body                 = $message
                Method               = $method
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
    Add-AddressBookEntry adds address book entries to Ricoh multi-function
    printers.

.Parameter Hostname
    The hostname of the printer from which address book entries are to be added.

    By default, it will use HTTP; if HTTPS is required, specify that in the URI,
    like "-Hostname https://printername"

.Parameter Credential
    The username and password to use to connect to the Ricoh printer.

    To reuse the credentials between commands, you can use the Get-Credential
    cmdlet and store the results to a variable. Otherwise, just passing a string
    will open a dialog box for the user to enter a password. See the help for
    Get-Credential for information on using this in a script non-interactively.

.Parameter Name
    The name for the address book entry.

.Parameter LongName
    The "long name" for the address book entry.

.Parameter FolderPath
    The network path used to save scanned files.

.Parameter ScanAccount
    The account to use to save the scanned files to a network location.

    To reuse the credentials between commands, you can use the Get-Credential
    cmdlet and store the results to a variable. Otherwise, just passing a string
    will open a dialog box for the user to enter a password. See the help for
    Get-Credential for information on using this in a script non-interactively.

.Parameter EmailAddress
    The email address used to send scanned files.

.Parameter IsSender
    Whether the given email address is registered as a sender.

.Parameter IsDestination
    Whether the given email address is registered as a destination.

.Parameter Frequent
    Whether the user is to be added to the frequently used list.

.Parameter Title1
    The heading to list this user under for Title1.

    Title1 is usually the default grouping, and is the one that lists users by
    letters (AB, CD, etc.) on the scanner.

.Parameter Title2
    The heading to list this user under for Title2.

    Title2 is a range from 1 to 10, and is another option for grouping users on
    the scanner.

.Parameter Title3
    The heading to list this user under for Title3.

    Title3 is a range from 1 to 5, and is another option for grouping users on
    the scanner.

.Parameter SkipCertificateCheck
    Skips certificate validation checks. This includes all validation such as
    expiration, revocation, trusted root authority, etc.

    > [!WARNING] Using this parameter is not secure and is not recommended. This
    switch is only intended to be used against known hosts using a self-signed
    certificate for testing purposes. Use at your own risk.

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

        [string]
        [Parameter(ParameterSetName = 'Folder', Mandatory, ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'FolderAndEmail', Mandatory, ValueFromPipelineByPropertyName)]
        $FolderPath,

        [pscredential]
        [Parameter(ParameterSetName = 'Folder', ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'FolderAndEmail', ValueFromPipelineByPropertyName)]
        $ScanAccount,

        [string]
        [Parameter(ParameterSetName = 'Email', Mandatory, ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'FolderAndEmail', Mandatory, ValueFromPipelineByPropertyName)]
        $EmailAddress,

        [nullable[bool]]
        [Parameter(ParameterSetName = 'Email', ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'FolderAndEmail', ValueFromPipelineByPropertyName)]
        $IsSender,

        [nullable[bool]]
        [Parameter(ValueFromPipelineByPropertyName)]
        $IsDestination,

        [nullable[bool]]
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
                Hostname             = $Hostname
                Credential           = $Credential
                SkipCertificateCheck = $SkipCertificateCheck
            }

            $session = Connect-Session @connection
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        $method = [RicohMethodType]::putObjects
        [xml] $template = Get-Template $method
        $content = $template.Envelope.Body.$method
        $content.sessionId = $session
    }

    process {
        # Tags (Frequent, Title1, Title2, Title3)
        $tagId = Get-TagIdValue

        $entry = $template.CreateElement('item')
        $content.propListList.AppendChild($entry) > $null

        function add($key, $value) {
            Add-Property $entry $key $value
        }

        add 'entryType' 'user'
        add 'name' $Name
        add 'longName' $LongName
        add 'tagId' $tagId

        if (-not [string]::IsNullOrEmpty($FolderPath)) {
            add 'remoteFolder:path' $FolderPath

            if ($null -ne $ScanAccount) {
                add 'remoteFolder:accountName' $ScanAccount.UserName
                add 'remoteFolder:password' (ConvertTo-Base64 $ScanAccount.GetNetworkCredential().Password)
            }

            add 'remoteFolder:port' 21
            add 'remoteFolder:select' 'private'
        }

        if (-not [string]::IsNullOrEmpty($EmailAddress)) {
            if ($null -eq $IsSender) {
                $IsSender = $false
            }
            add 'mail:' 'true'
            add 'mail:address' $EmailAddress
            add 'isSender' $IsSender.ToString().ToLower()
        }

        if ($null -eq $IsDestination) {
            $IsDestination = $true
        }
        add 'isDestination' $IsDestination.ToString().ToLower()
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
                Hostname             = $Hostname
                Body                 = $template
                Method               = $method
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

    To reuse the credentials between commands, you can use the Get-Credential
    cmdlet and store the results to a variable. Otherwise, just passing a string
    will open a dialog box for the user to enter a password. See the help for
    Get-Credential for information on using this in a script non-interactively.

.Parameter Id
    The IDs to be removed. Find the IDs from Get-AddressBookEntry. ID is **not**
    the Index.

.Parameter SkipCertificateCheck
    Skips certificate validation checks. This includes all validation such as
    expiration, revocation, trusted root authority, etc.

    > [!WARNING] Using this parameter is not secure and is not recommended. This
    switch is only intended to be used against known hosts using a self-signed
    certificate for testing purposes. Use at your own risk.

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
    [CmdletBinding(SupportsShouldProcess, PositionalBinding = $false)]
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
                Hostname             = $Hostname
                Credential           = $Credential
                SkipCertificateCheck = $SkipCertificateCheck
            }

            $session = Connect-Session @connection
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }

        $method = [RicohMethodType]::deleteObjects
        [xml] $template = Get-Template $method
        $content = $template.Envelope.Body.$method
        $content.sessionId = $session
    }

    process {
        foreach ($item in $Id) {
            $element = $template.CreateElement('item')
            $element.InnerText = "entry:$item"
            $content.objectIdList.AppendChild($element) > $null
        }
    }

    end {
        $entries = Select-Xml -Xml $template -Namespace $namespaces -XPath '/s:Envelope/s:Body/u:deleteObjects/objectIdList/item' |
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
                Hostname             = $Hostname
                Body                 = $template
                Method               = $method
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
    [xml] $template = Get-Template $method
    $template.Envelope.Body.$method.sessionId = $session

    $request = @{
        Hostname             = $Hostname
        Body                 = $template
        Method               = $method
        SkipCertificateCheck = $SkipCertificateCheck
    }

    try {
        Invoke-SOAPRequest @request > $null
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
