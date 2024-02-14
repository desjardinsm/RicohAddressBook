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
    } catch [System.Net.WebException] {
        $message = [xml]$_.ErrorDetails.Message
        $fault = $message.Envelope.Body.Fault

        $errorMessage = $fault.detail.rdhError.errorDescription
        if ([string]::IsNullOrEmpty($errorMessage)) {
            $errorMessage = $fault.detail.rdhError.errorCode

            if ([string]::IsNullOrEmpty($errorMessage)) {
                $errorMessage = $fault.faultstring
            }
        }

        if (-not [string]::IsNullOrEmpty($errorMessage)) {
            $_.ErrorDetails = [System.Management.Automation.ErrorDetails]::new(
                "The server returned an error ($errorMessage)."
            )
        }

        $PSCmdlet.ThrowTerminatingError($_)
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

<#
.Synopsis
    Returns a Title1 TagId for a given letter

.Description
    Get-Title1Tag returns a TagId for use in the Title1 parameter of the
    Update-AddressBookEntry or Add-AddressBookEntry functions.

    TagId is an enum that groups letters together, matching the Title1 grouping
    seen on the scanner itself:
        - AB
        - CD
        - EF
        - GH
        - IJK
        - LMN
        - OPQ
        - RST
        - UVW
        - XYZ

    NOTE: This function is not strictly necessary, as just passing a string
    (like "-Title1 AB") will cast this value to the correct enum value. This is
    more useful when adding or updating entries in batches (for example, to
    automatically generate a Title1 TagId from the first letter of each entry's
    name); see the first example for a possible use case.

.Parameter Letter
    The letter to use to look up the correct TagId enum value.

.Inputs
    None. You cannot pipe input to Get-Title1Tag

.Outputs
    A TagId enum value representing the letter provided.

.Example
    PS> $entries = @(
        @{
            Name = 'Matthew D'
            KeyDisplay = 'Matt D'
            FolderScanPath = '\\folder\path'
        }
        @{
            Name = 'John D'
            KeyDisplay = 'John D'
            FolderScanPath = '\\folder\path'
        }
    ) | ForEach-Object {
        $_.Title1 = Get-Title1Tag $_.Name[0]

        [PSCustomObject]$_
    }

    PS> $entries |
        Add-AddressBookEntry -Hostname https://10.10.10.10 -Credential admin

.Example
    PS> # Using delay-bound parameters

    PS> @(
            [PSCustomObject]@{
                Name = 'Matthew D'
                KeyDisplay = 'Matt D'
                FolderScanPath = '\\folder\path'
            }
            [PSCustomObject]@{
                Name = 'John D'
                KeyDisplay = 'John D'
                FolderScanPath = '\\folder\path'
            }
        ) |
            Add-AddressBookEntry `
                -Hostname https://10.10.10.10 `
                -Credential admin `
                -Title1 { Get-Title1Tag $_.KeyDisplay[0] }

    These commands use a delay-bound parameter, which is an implicit feature of
    PowerShell for parameters that can take pipeline input (whether by value or,
    like in this case, by name) and have a type that is not [scriptblock] or
    [object]. The scriptblock provided to the -Title1 parameter will be called
    for each object that is piped in, with the pipeline variable "$_"
    representing the current object. The scriptblock's return value will then be
    bound to "Title1" for the current pipeline object.
#>
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

function Format-PropertyList {
    [OutputType('Ricoh.AddressBook.Entry')]
    param(
        [System.Xml.XmlNode]
        $PropertyList
    )

    $properties = @{}
    foreach ($property in $PropertyList.ChildNodes) {
        if (-not [string]::IsNullOrEmpty($property.propVal)) {
            $properties[$property.propName] = $property.propVal
        }
    }

    $output = [ordered]@{
        PSTypeName = 'Ricoh.AddressBook.Entry'
    }

    if ($properties.ContainsKey('id')) {
        $output.ID = [uint32]$properties['id']
    }

    if ($properties.ContainsKey('index')) {
        $output.RegistrationNumber = [uint32]$properties['index']
    }

    if ($properties.ContainsKey('name')) {
        $output.Name = $properties['name']
    }

    if ($properties.ContainsKey('longName')) {
        $output.KeyDisplay = $properties['longName']
    }

    if ($properties.ContainsKey('displayedOrder')) {
        $output.DisplayPriority = [uint32]$properties['displayedOrder']
    }

    if ($properties.ContainsKey('tagId')) {
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
    }

    if (Test-Property $properties 'auth:') {
        $output.UserCode = $properties['auth:name']
    }

    if ($properties.ContainsKey('lastAccessDateTime') -and $properties['lastAccessDateTime'] -ne '1970-01-01T00:00:00Z') {
        $output.LastUsed = [datetime]$properties['lastAccessDateTime']
    }

    if (Test-Property $properties 'remoteFolder:') {
        $output.FolderScanType = $properties['remoteFolder:type']
        $output.FolderScanPath = $properties['remoteFolder:path']
        $output.FolderScanPort = [uint32]$properties['remoteFolder:port']

        if ($properties.ContainsKey('remoteFolder:select') -and 'private' -eq $properties['remoteFolder:select']) {
            $output.FolderScanAccount = $properties['remoteFolder:accountName']
        }
    }

    if (Test-Property $properties 'mail:') {
        $output.EmailAddress = $properties['mail:address']
        $output.IsSender = $properties['isSender'] -eq 'true'
    }

    if ($properties.ContainsKey('isDestination')) {
        $output.IsDestination = $properties['isDestination'] -eq 'true'
    }

    [PSCustomObject]$output
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
        $disconnect = @{
            Hostname             = $Hostname
            SkipCertificateCheck = $SkipCertificateCheck
        }
        $connection = @{
            Hostname             = $Hostname
            Credential           = $Credential
            SkipCertificateCheck = $SkipCertificateCheck
        }

        $session = Connect-Session @connection -ReadOnly
        $disconnect.Session = $session

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
            $result = Format-PropertyList $entry.Node

            if ([string]::IsNullOrEmpty($Name) -or $result.Name -like $Name) {
                $result
            }
        }
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    } finally {
        if ($null -ne $disconnect.Session) {
            Disconnect-Session @disconnect
        }
    }
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

    For the non-required parameters, they can be omitted in order to not modify
    their values. The one exception to this is if Frequent is true, or if
    Title1, Title2, or Title3 are provided, the value of the other three will be
    reset (unless a value is also provided for them). This is because these four
    values are stored on the scanner as a single property, and setting one
    without the other three will set that property to that single value.

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

.Parameter KeyDisplay
    The new key display for the address book entry.

.Parameter DisplayPriority
    The display order of the user in address book list. Sorting is done first by
    DisplayPriority, then by ID.

.Parameter Frequent
    A switch indicating whether the user is to be added to the frequently used
    list.

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

.Parameter UserCode
    The User Code property used for authentication management.

.Parameter FolderScanPath
    The network path used to save scanned files.

.Parameter FolderScanAccount
    The account to use to save the scanned files to a network location.

    To reuse the credentials between commands, you can use the Get-Credential
    cmdlet and store the results to a variable. Otherwise, just passing a string
    will open a dialog box for the user to enter a password. See the help for
    Get-Credential for information on using this in a script non-interactively.

.Parameter EmailAddress
    The email address used to send scanned files.

.Parameter IsSender
    A boolean indicating whether the given email address is registered as a
    sender.

.Parameter IsDestination
    A boolean indicating whether the given entry is registered as a destination.

    If this value is set to false, this entry will not be visible on the scanner
    as a valid destination.

.Parameter ForceUserCode
    Force setting the user code property, even if $UserCode is empty.

    If $UserCode is empty and this switch is set, the user code will be removed
    from this address book entry.

.Parameter ForceFolderScanPath
    Force setting the folder scan path, even if $FolderScanPath is empty.

    If $FolderScanPath is empty and this switch is set, folder scanning will be
    removed from this address book entry.

.Parameter ForceFolderScanAccount
    Force setting the folder scan account, even if $FolderScanAccount is empty.

    If $FolderScanAccount is empty and this switch is set, the folder scan
    account will be removed from this address book entry, and folder scanning
    will instead use the file transfer account from the device configuration.

.Parameter ForceEmailAddress
    Force setting the email address, even if $EmailAddress is empty.

    If $EmailAddress is empty and this switch is set, email scanning will be
    removed from this address book entry.

.Parameter PassThru
    If -PassThru is true, return a Ricoh.AddressBook.Entry object for each
    address book entry with the values that were updated.

.Parameter SkipCertificateCheck
    Skips certificate validation checks. This includes all validation such as
    expiration, revocation, trusted root authority, etc.

    > [!WARNING] Using this parameter is not secure and is not recommended. This
    switch is only intended to be used against known hosts using a self-signed
    certificate for testing purposes. Use at your own risk.

.Inputs
    Custom objects used to update multiple entries at once.

.Outputs
    None, unless the -PassThru switch is provided. If -PassThru is provided,
    then outputs a Ricoh.AddressBook.Entry object for each address book entry
    with the values that were updated.

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

        [uint32]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $Id,

        [string]
        [ValidateLength(1, 20)]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Name,

        [string]
        [ValidateLength(1, 16)]
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('LongName')]
        $KeyDisplay,

        [byte]
        [ValidateRange(1, 10)]
        [Parameter(ValueFromPipelineByPropertyName)]
        $DisplayPriority,

        [switch]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Frequent,

        [TagId]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Title1,

        [byte]
        [ValidateRange(1, 10)]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Title2,

        [byte]
        [ValidateRange(1, 5)]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Title3,

        [string]
        [ValidateLength(1, 8)]
        [ValidatePattern('^\d+$')]
        [Parameter(ValueFromPipelineByPropertyName)]
        $UserCode,

        [string]
        [ValidateLength(1, 256)]
        [Parameter(ValueFromPipelineByPropertyName)]
        $FolderScanPath,

        [pscredential]
        [Parameter(ValueFromPipelineByPropertyName)]
        $FolderScanAccount,

        [string]
        [ValidateLength(1, 128)]
        [Parameter(ValueFromPipelineByPropertyName)]
        $EmailAddress,

        [nullable[bool]]
        [ValidateNotNull()]
        [Parameter(ValueFromPipelineByPropertyName)]
        $IsSender,

        [nullable[bool]]
        [ValidateNotNull()]
        [Parameter(ValueFromPipelineByPropertyName)]
        $IsDestination,

        [switch]
        [Parameter(ValueFromPipelineByPropertyName)]
        $ForceUserCode,

        [switch]
        [Parameter(ValueFromPipelineByPropertyName)]
        $ForceFolderScanPath,

        [switch]
        [Parameter(ValueFromPipelineByPropertyName)]
        $ForceFolderScanAccount,

        [switch]
        [Parameter(ValueFromPipelineByPropertyName)]
        $ForceEmailAddress,

        [switch]
        $PassThru,

        [switch]
        $SkipCertificateCheck
    )

    begin {
        try {
            $disconnect = @{
                Hostname             = $Hostname
                SkipCertificateCheck = $SkipCertificateCheck
            }
            $connection = @{
                Hostname             = $Hostname
                Credential           = $Credential
                SkipCertificateCheck = $SkipCertificateCheck
            }

            $session = Connect-Session @connection
            $disconnect.Session = $session

            $method = [RicohMethodType]::putObjectProps
            [xml] $template = Get-Template $method
            $template.Envelope.Body.$method.sessionId = $session
        } catch {
            if ($null -ne $disconnect.Session) {
                Disconnect-Session @disconnect
            }
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }

    process {
        try {
            [xml] $message = $template.Clone()
            $content = $message.Envelope.Body.$method
            $content.objectId = "entry:$Id"

            function add($key, $value) {
                Add-Property $content.propList $key $value
            }

            if (-not [string]::IsNullOrEmpty($Name)) {
                add 'name' $Name
            }
            if (-not [string]::IsNullOrEmpty($KeyDisplay)) {
                add 'longName' $KeyDisplay
            }

            if (0 -ne $DisplayPriority) {
                add 'displayedOrder' $DisplayPriority
            }

            # Tags (Frequent, Title1, Title2, Title3)
            $tagId = Get-TagIdValue
            if (-not [string]::IsNullOrEmpty($tagId)) {
                add 'tagId' $tagId
            }

            if (-not [string]::IsNullOrEmpty($UserCode)) {
                add 'auth:' 'true'
                add 'auth:name' $UserCode
            } elseif ($ForceUserCode) {
                add 'auth:' 'false'
            }

            if (-not [string]::IsNullOrEmpty($FolderScanPath)) {
                add 'remoteFolder:' 'true'
                add 'remoteFolder:type' 'smb'
                add 'remoteFolder:path' $FolderScanPath
                add 'remoteFolder:port' 21
            } elseif ($ForceFolderScanPath) {
                add 'remoteFolder:' 'false'
            }

            if ($null -ne $FolderScanAccount) {
                add 'remoteFolder:select' 'private'
                add 'remoteFolder:accountName' $FolderScanAccount.UserName
                add 'remoteFolder:password' (ConvertTo-Base64 $FolderScanAccount.GetNetworkCredential().Password)
            } elseif ($ForceFolderScanAccount) {
                add 'remoteFolder:select' ''
            }

            if (-not [string]::IsNullOrEmpty($EmailAddress)) {
                add 'mail:' 'true'
                add 'mail:address' $EmailAddress
            } elseif ($ForceEmailAddress) {
                add 'mail:' 'false'
            }

            if ($null -ne $IsSender) {
                add 'isSender' $IsSender.ToString().ToLower()
            }
            if ($null -ne $IsDestination) {
                add 'isDestination' $IsDestination.ToString().ToLower()
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

            if ($PassThru) {
                $result = Format-PropertyList $content.propList
                Add-Member -InputObject $result -MemberType NoteProperty -Name ID -Value $Id
                $result
            }
        } catch {
            Write-Error $_
        }
    }

    end {
        Disconnect-Session @disconnect
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

.Parameter KeyDisplay
    The key display for the address book entry.

.Parameter DisplayPriority
    The display order of the user in address book list. Sorting is done first by
    DisplayPriority, then by ID.

.Parameter Frequent
    A boolean indicating whether the user is to be added to the frequently used
    list. At least one of either Frequent, Title1, Title2, or Title3 should be
    set.

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

.Parameter UserCode
    The User Code property used for authentication management.

.Parameter FolderScanPath
    The network path used to save scanned files.

.Parameter FolderScanAccount
    The account to use to save the scanned files to a network location.

    To reuse the credentials between commands, you can use the Get-Credential
    cmdlet and store the results to a variable. Otherwise, just passing a string
    will open a dialog box for the user to enter a password. See the help for
    Get-Credential for information on using this in a script non-interactively.

.Parameter EmailAddress
    The email address used to send scanned files.

.Parameter IsSender
    A boolean indicating whether the given email address is registered as a
    sender. Omit to use the default value of $false.

.Parameter IsDestination
    A boolean indicating whether the given entry is registered as a destination.
    Omit to use the default value of $true.

    If this value is set to false, this entry will not be visible on the scanner
    as a valid destination.

.Parameter PassThru
    If -PassThru is true, return a Ricoh.AddressBook.Entry object for each
    address book entry that was submitted.

.Parameter SkipCertificateCheck
    Skips certificate validation checks. This includes all validation such as
    expiration, revocation, trusted root authority, etc.

    > [!WARNING] Using this parameter is not secure and is not recommended. This
    switch is only intended to be used against known hosts using a self-signed
    certificate for testing purposes. Use at your own risk.

.Inputs
    Custom objects used to add multiple entries at once.

.Outputs
    None, unless the -PassThru switch is provided. If -PassThru is provided,
    then outputs a Ricoh.AddressBook.Entry object for each address book entry
    that was submitted.

.Example
    PS> $entry = @{
        Hostname = 'https://10.10.10.10'
        Credential = Get-Credential admin
        Name = 'Matthew D'
        KeyDisplay = 'Matthew Desjardins'
        FolderScanPath = '\\my\path\here'
        FolderScanAccount = Get-Credential ScanAccount
    }
    PS> Add-AddressBookEntry @entry

.Example
    PS> $scanAccount = Get-Credential ScanAccount
    PS> $entries = @(
        [PSCustomObject]@{
            Name = 'Matthew D'
            KeyDisplay = 'Matthew Desjardins'
            Frequent = $true
            Title1 = 'LMN'
            FolderScanPath = '\\my\path\here'
            FolderScanAccount = $scanAccount
        }
        [PSCustomObject]@{
            Name = 'John D'
            KeyDisplay = 'John Doe'
            Title1 = 'IJK'
            FolderScanPath = '\\my\path\here'
            FolderScanAccount = $scanAccount
        }
    )
    PS> $entries | Add-AddressBookEntry -Hostname https://10.10.10.10 -Credential admin

    Adds multiple entries based on objects received in the pipeline.
#>
function Add-AddressBookEntry {
    [CmdletBinding(SupportsShouldProcess, PositionalBinding = $false)]
    param(
        [uri]
        [Parameter(Mandatory)]
        $Hostname,

        [pscredential]
        [Parameter(Mandatory)]
        $Credential,

        [string]
        [ValidateLength(1, 20)]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $Name,

        [string]
        [ValidateLength(1, 16)]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('LongName')]
        $KeyDisplay,

        [byte]
        [ValidateRange(1, 10)]
        [Parameter(ValueFromPipelineByPropertyName)]
        $DisplayPriority = 5,

        [bool]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Frequent = $true,

        [TagId]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Title1,

        [byte]
        [ValidateRange(1, 10)]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Title2,

        [byte]
        [ValidateRange(1, 5)]
        [Parameter(ValueFromPipelineByPropertyName)]
        $Title3,

        [string]
        [ValidateLength(1, 8)]
        [ValidatePattern('^\d+$')]
        [Parameter(ValueFromPipelineByPropertyName)]
        $UserCode,

        [string]
        [ValidateLength(1, 256)]
        [Parameter(ParameterSetName = 'Folder', Mandatory, ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'FolderAndEmail', Mandatory, ValueFromPipelineByPropertyName)]
        $FolderScanPath,

        [pscredential]
        [Parameter(ParameterSetName = 'Folder', ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'FolderAndEmail', ValueFromPipelineByPropertyName)]
        $FolderScanAccount,

        [string]
        [ValidateLength(1, 128)]
        [Parameter(ParameterSetName = 'Email', Mandatory, ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'FolderAndEmail', Mandatory, ValueFromPipelineByPropertyName)]
        $EmailAddress,

        [nullable[bool]]
        [ValidateNotNull()]
        [Parameter(ParameterSetName = 'Email', ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'FolderAndEmail', ValueFromPipelineByPropertyName)]
        [PSDefaultValue(Help = $false, Value = $false)]
        $IsSender,

        [nullable[bool]]
        [ValidateNotNull()]
        [Parameter(ValueFromPipelineByPropertyName)]
        [PSDefaultValue(Help = $true, Value = $true)]
        $IsDestination,

        [switch]
        $PassThru,

        [switch]
        $SkipCertificateCheck
    )

    begin {
        try {
            $disconnect = @{
                Hostname             = $Hostname
                SkipCertificateCheck = $SkipCertificateCheck
            }
            $connection = @{
                Hostname             = $Hostname
                Credential           = $Credential
                SkipCertificateCheck = $SkipCertificateCheck
            }

            $session = Connect-Session @connection
            $disconnect.Session = $session

            $method = [RicohMethodType]::putObjects
            [xml] $template = Get-Template $method
            $content = $template.Envelope.Body.$method
            $content.sessionId = $session
        } catch {
            if ($null -ne $disconnect.Session) {
                Disconnect-Session @disconnect
            }
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }

    process {
        try {
            # Tags (Frequent, Title1, Title2, Title3)
            $tagId = Get-TagIdValue

            if ([string]::IsNullOrEmpty($tagId)) {
                throw 'At least one of the following is required: Frequent, Title1, Title2, or Title3'
            }

            $entry = $template.CreateElement('item')

            function add($key, $value) {
                Add-Property $entry $key $value
            }

            add 'entryType' 'user'
            add 'name' $Name
            add 'longName' $KeyDisplay
            add 'displayedOrder' $DisplayPriority
            add 'tagId' $tagId

            if (-not [string]::IsNullOrEmpty($UserCode)) {
                add 'auth:' 'true'
                add 'auth:name' $UserCode
            }

            if (-not [string]::IsNullOrEmpty($FolderScanPath)) {
                add 'remoteFolder:' 'true'
                add 'remoteFolder:type' 'smb'
                add 'remoteFolder:path' $FolderScanPath
                add 'remoteFolder:port' 21

                if ($null -ne $FolderScanAccount) {
                    add 'remoteFolder:select' 'private'
                    add 'remoteFolder:accountName' $FolderScanAccount.UserName
                    add 'remoteFolder:password' (ConvertTo-Base64 $FolderScanAccount.GetNetworkCredential().Password)
                }
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

            $content.propListList.AppendChild($entry) > $null
        } catch {
            Write-Error $_
        }
    }

    end {
        try {
            $selection = @{
                Xml       = $template
                Namespace = $namespaces
                XPath     = '/s:Envelope/s:Body/u:putObjects/propListList/item/item[propName/text()="name"]/propVal'
            }
            $names = Select-Xml @selection

            $allNames = $names -join ', '
            if ($allNames.Length -gt 0 -and $PSCmdlet.ShouldProcess(
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

            if ($PassThru) {
                $selection = @{
                    Xml       = $template
                    Namespace = $namespaces
                    XPath     = '/s:Envelope/s:Body/u:putObjects/propListList/item'
                }
                $entries = Select-Xml @selection

                foreach ($entry in $entries) {
                    Format-PropertyList $entry.Node
                }
            }
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        } finally {
            Disconnect-Session @disconnect
        }
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
                 Where-Object {[string]::IsNullOrEmpty($_.KeyDisplay)}
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

        [uint32[]]
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        $Id,

        [switch]
        $SkipCertificateCheck
    )

    begin {
        try {
            $disconnect = @{
                Hostname             = $Hostname
                SkipCertificateCheck = $SkipCertificateCheck
            }
            $connection = @{
                Hostname             = $Hostname
                Credential           = $Credential
                SkipCertificateCheck = $SkipCertificateCheck
            }

            $session = Connect-Session @connection
            $disconnect.Session = $session

            $method = [RicohMethodType]::deleteObjects
            [xml] $template = Get-Template $method
            $content = $template.Envelope.Body.$method
            $content.sessionId = $session
        } catch {
            if ($null -ne $disconnect.Session) {
                Disconnect-Session @disconnect
            }
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }

    process {
        foreach ($item in $Id) {
            try {
                $element = $template.CreateElement('item')
                $element.InnerText = "entry:$item"
                $content.objectIdList.AppendChild($element) > $null
            } catch {
                Write-Error $_
            }
        }
    }

    end {
        try {
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
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        } finally {
            Disconnect-Session @disconnect
        }
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
