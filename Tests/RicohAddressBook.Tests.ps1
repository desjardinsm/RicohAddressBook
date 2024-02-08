BeforeAll {
    Import-Module -Force -Name (Join-Path $PSScriptRoot 'TestSetup')
    Import-Module -Force -Name (Join-Path $PSScriptRoot '../Module/RicohAddressBook')

    $password = ConvertTo-SecureString -String 'MockPassword' -AsPlainText -Force
    $credential = [pscredential]::new('admin', $password)

    $commonParameters = @{}
    $commonParameters.Hostname = '10.10.10.10'
    $commonParameters.Credential = $credential

    & {
        $mockParameters = @{
            ModuleName  = 'RicohAddressBook'
            CommandName = 'Invoke-WebRequest'
        }

        Mock @mockParameters {}
        Mock @mockParameters (Get-SoapResponseMock startSession) -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#startSession'
        }
        Mock @mockParameters (Get-SoapResponseMock searchObjects) -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#searchObjects'
        }
        Mock @mockParameters (Get-SoapResponseMock getObjectsProps) -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#getObjectsProps'
        }
    }
}

AfterAll {
    Remove-Module TestSetup
    Remove-Module Pester
}

Describe 'Invoke-WebRequest' {
    Context 'With an HTTP URL' {
        It 'Is called with the required parameters' {
            Get-AddressBookEntry @commonParameters

            Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Times 1 -ParameterFilter {
                $Uri -eq [uri]'http://10.10.10.10/DH/udirectory' -and
                $Method -eq [Microsoft.PowerShell.Commands.WebRequestMethod]::Post -and
                $ContentType -eq 'text/xml'
            }
        }
    }

    Context 'With an HTTPS URL' {
        Context 'When not called with -SkipCertificateCheck' {
            It 'Is called with the required parameters' {
                Get-AddressBookEntry -Hostname "https://$($commonParameters.Hostname)" -Credential $commonParameters.Credential

                Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Times 1 -ParameterFilter {
                    $Uri -eq [uri]'https://10.10.10.10/DH/udirectory' -and
                    $Method -eq [Microsoft.PowerShell.Commands.WebRequestMethod]::Post -and
                    $ContentType -eq 'text/xml' -and
                    -not $SkipCertificateCheck
                }
            }
        }
    }
}

Describe 'Connect-Session' {
    Context 'With a read-only method' {
        It 'Creates a connection with an S lock mode for Get-AddressBookEntry' {
            Get-AddressBookEntry @commonParameters

            Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Times 1 -ParameterFilter {
                $expected = [xml]'<?xml version="1.0" encoding="utf-8"?>
                <s:Envelope
                        xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                        s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                    <s:Body>
                        <u:startSession
                                xmlns:xs="http://www.w3.org/2001/XMLSchema"
                                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                            <stringIn>SCHEME=QkFTSUM=;UID:UserName=YWRtaW4=;PWD:Password=TW9ja1Bhc3N3b3Jk;PES:Encoding=</stringIn>
                            <timeLimit>30</timeLimit>
                            <lockMode>S</lockMode>
                        </u:startSession>
                    </s:Body>
                </s:Envelope>'

                $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#startSession' -and
                $Body.OuterXml -eq $expected.OuterXml
            }
        }
    }

    Context 'With an editing method' {
        It -ForEach @(
            @{ Function = 'Update'; Arguments = @{ Id = 1 } }
            @{ Function = 'Add'; Arguments = @{ Name = 'A'; KeyDisplay = 'B'; EmailAddress = 'example@example.com'; Frequent = $true } }
            @{ Function = 'Remove'; Arguments = @{Id = @(1) } }
        ) 'Creates a connection with an X lock mode for <Function>-AddressBookEntry' {
            & "$Function-AddressBookEntry" @commonParameters @Arguments

            Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                $expected = [xml]'<?xml version="1.0" encoding="utf-8"?>
                <s:Envelope
                        xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                        s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                    <s:Body>
                        <u:startSession
                                xmlns:xs="http://www.w3.org/2001/XMLSchema"
                                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                            <stringIn>SCHEME=QkFTSUM=;UID:UserName=YWRtaW4=;PWD:Password=TW9ja1Bhc3N3b3Jk;PES:Encoding=</stringIn>
                            <timeLimit>30</timeLimit>
                            <lockMode>X</lockMode>
                        </u:startSession>
                    </s:Body>
                </s:Envelope>'

                $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#startSession' -and
                $Body.OuterXml -eq $expected.OuterXml
            }
        }
    }
}

Describe 'Disconnect-Session' {
    It -ForEach @(
        @{ Function = 'Get'; Arguments = @{} }
        @{ Function = 'Update'; Arguments = @{ Id = 1 } }
        @{ Function = 'Add'; Arguments = @{ Name = 'A'; KeyDisplay = 'B'; EmailAddress = 'example@example.com'; Frequent = $true } }
        @{ Function = 'Remove'; Arguments = @{Id = @(1) } }
    ) 'Disconnects the session for <Function>-AddressBookEntry' {
        & "$Function-AddressBookEntry" @commonParameters @Arguments

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $expected = [xml]'<?xml version="1.0" encoding="utf-8"?>
            <s:Envelope
                    xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                <s:Body>
                    <u:terminateSession
                            xmlns:xs="http://www.w3.org/2001/XMLSchema"
                            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                            xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                            xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                            xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                        <sessionId>12345</sessionId>
                    </u:terminateSession>
                </s:Body>
            </s:Envelope>'

            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#terminateSession' -and
            $Body.OuterXml -eq $expected.OuterXml
        }
    }

    Context 'Errored operations still call Disconnect-Session' {
        Context 'When Connect-Session' {
            Context -ForEach @('script', 'statement') 'throws a <_>-terminating error' {
                BeforeAll {
                    $throw = $_ -eq 'script'
                    Mock -ModuleName RicohAddressBook Invoke-WebRequest {
                        if ($throw) {
                            throw 'Script-terminating error'
                        }

                        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                            [System.Exception]::new('Statement-terminating-error'),
                            'StatementTerminatingErrorId',
                            'NotSpecified',
                            $null
                        )
                        $PSCmdlet.ThrowTerminatingError($errorRecord)
                    } -ParameterFilter {
                        $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#startSession'
                    }
                }

                It -ForEach @(
                    @{ Function = 'Get'; Arguments = @{} }
                    @{ Function = 'Update'; Arguments = @{ Id = 1 } }
                    @{ Function = 'Add'; Arguments = @{ Name = 'A'; KeyDisplay = 'B'; EmailAddress = 'example@example.com' } }
                    @{ Function = 'Remove'; Arguments = @{Id = @(1) } }
                ) '<Function>-AddressBookEntry does not need to call Disconnect-Session' {
                    try {
                        & "$Function-AddressBookEntry" @commonParameters @Arguments
                    } catch {}

                    Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Times 0 -ParameterFilter {
                        $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#terminateSession'
                    }
                }
            }
        }

        Context 'When Get-Template' {
            Context -ForEach @('script', 'statement') 'throws a <_>-terminating error' {
                BeforeAll {
                    $throw = $_ -eq 'script'
                    Mock -ModuleName RicohAddressBook Get-Template {
                        if ($throw) {
                            throw 'Script-terminating exception'
                        }

                        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                            [System.Exception]::new('Statement-terminating-error'),
                            'StatementTerminatingErrorId',
                            'NotSpecified',
                            $null
                        )
                        $PSCmdlet.ThrowTerminatingError($errorRecord)
                    } -ParameterFilter {
                        $Method -ne 'startSession' -and $Method -ne 'terminateSession'
                    }
                }

                It -ForEach @(
                    @{ Function = 'Get'; Arguments = @{} }
                    @{ Function = 'Update'; Arguments = @{ Id = 1 } }
                    @{ Function = 'Add'; Arguments = @{ Name = 'A'; KeyDisplay = 'B'; EmailAddress = 'example@example.com' } }
                    @{ Function = 'Remove'; Arguments = @{Id = @(1) } }
                ) '<Function>-AddressBookEntry calls Disconnect-Session' {
                    try {
                        & "$Function-AddressBookEntry" @commonParameters @Arguments
                    } catch {}

                    Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                        $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#terminateSession'
                    }
                }
            }
        }

        Context 'When an editing operation' {
            Context -ForEach @('script', 'statement') 'throws a <_>-terminating error' {
                BeforeAll {
                    $throw = $_ -eq 'script'
                    Mock -ModuleName RicohAddressBook Invoke-WebRequest {
                        if ($throw) {
                            throw 'Script-terminating exception'
                        }

                        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                            [System.Exception]::new('Statement-terminating-error'),
                            'StatementTerminatingErrorId',
                            'NotSpecified',
                            $null
                        )
                        $PSCmdlet.ThrowTerminatingError($errorRecord)
                    } -ParameterFilter {
                        $Headers.SOAPAction -ne 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#startSession' -and
                        $Headers.SOAPAction -ne 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#terminateSession'
                    }
                }

                It -ForEach @(
                    @{ Function = 'Update'; Arguments = @{ Id = 1 } }
                    @{ Function = 'Add'; Arguments = @{ Name = 'A'; KeyDisplay = 'B'; EmailAddress = 'example@example.com' } }
                    @{ Function = 'Remove'; Arguments = @{Id = @(1) } }
                ) '<Function>-AddressBookEntry calls Disconnect-Session' {
                    try {
                        & "$Function-AddressBookEntry" @commonParameters @Arguments -ErrorAction SilentlyContinue
                    } catch {}

                    Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                        $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#terminateSession'
                    }
                }
            }
        }

        Context 'When an element of the pipeline' {
            Context -ForEach @('script', 'statement') 'throws a <_>-terminating error' {
                BeforeAll {
                    Mock -ModuleName RicohAddressBook Connect-Session { '12345' }

                    $throw = $_ -eq 'script'
                    Mock -ModuleName RicohAddressBook ConvertTo-Base64 {
                        if ($throw) {
                            throw 'Script-terminating exception'
                        }

                        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                            [System.Exception]::new('Statement-terminating-error'),
                            'StatementTerminatingErrorId',
                            'NotSpecified',
                            $null
                        )
                        $PSCmdlet.ThrowTerminatingError($errorRecord)
                    } -ParameterFilter { $String -eq 'throws' }
                }

                Context 'Update-AddressBookEntry' {
                    It 'Calls Disconnect-Session' {
                        @(
                            [PSCustomObject]@{
                                Id   = 1
                                Name = 'New Name 1'
                            }
                            [PSCustomObject]@{
                                Id                = 2
                                FolderScanAccount = [pscredential]::new(
                                    'NewScanAccount',
                                    (ConvertTo-SecureString -String 'throws' -AsPlainText -Force)
                                )
                            }
                            [PSCustomObject]@{
                                Id   = 3
                                Name = 'New Name 2'
                            }
                        ) | Update-AddressBookEntry @commonParameters -ErrorAction SilentlyContinue

                        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#terminateSession'
                        }
                    }

                    It 'Updates the other elements of the pipeline' {
                        @(
                            [PSCustomObject]@{
                                Id   = 1
                                Name = 'New Name 1'
                            }
                            [PSCustomObject]@{
                                Id                = 2
                                FolderScanAccount = [pscredential]::new(
                                    'NewScanAccount',
                                    (ConvertTo-SecureString -String 'throws' -AsPlainText -Force)
                                )
                            }
                            [PSCustomObject]@{
                                Id   = 3
                                Name = 'New Name 2'
                            }
                        ) | Update-AddressBookEntry @commonParameters -ErrorAction SilentlyContinue

                        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 2 -ParameterFilter {
                            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps'
                        }

                        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
                            $Body.OuterXml -eq ([xml]'<?xml version="1.0" encoding="utf-8"?>
                            <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                                <s:Body>
                                    <u:putObjectProps xmlns:xs="http://www.w3.org/2001/XMLSchema"
                                            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                            xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                            xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                            xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                                        <sessionId>12345</sessionId>
                                        <objectId>entry:1</objectId>
                                        <propList xsi:type="soap-enc:Array" soap-enc:arrayType="ricoh:property[]">
                                            <item>
                                                <propName>name</propName>
                                                <propVal>New Name 1</propVal>
                                            </item>
                                        </propList>
                                    </u:putObjectProps>
                                </s:Body>
                            </s:Envelope>').OuterXml
                        }

                        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
                            $Body.OuterXml -eq ([xml]'<?xml version="1.0" encoding="utf-8"?>
                            <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                                <s:Body>
                                    <u:putObjectProps xmlns:xs="http://www.w3.org/2001/XMLSchema"
                                            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                            xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                            xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                            xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                                        <sessionId>12345</sessionId>
                                        <objectId>entry:3</objectId>
                                        <propList xsi:type="soap-enc:Array" soap-enc:arrayType="ricoh:property[]">
                                            <item>
                                                <propName>name</propName>
                                                <propVal>New Name 2</propVal>
                                            </item>
                                        </propList>
                                    </u:putObjectProps>
                                </s:Body>
                            </s:Envelope>').OuterXml
                        }
                    }
                }

                Context 'Add-AddressBookEntry' {
                    It 'Calls Disconnect-Session' {
                        @(
                            [PSCustomObject]@{
                                Name           = 'Name 1'
                                KeyDisplay     = 'Key Display 1'
                                FolderScanPath = '\\folder\path1'
                            }
                            [PSCustomObject]@{
                                Name              = 'Name 2'
                                KeyDisplay        = 'Key Display 2'
                                FolderScanPath    = '\\folder\path2'
                                FolderScanAccount = [pscredential]::new(
                                    'NewScanAccount',
                                    (ConvertTo-SecureString -String 'throws' -AsPlainText -Force)
                                )
                            }
                            [PSCustomObject]@{
                                Name           = 'Name 3'
                                KeyDisplay     = 'Key Display 3'
                                FolderScanPath = '\\folder\path3'
                            }
                        ) | Add-AddressBookEntry @commonParameters -ErrorAction SilentlyContinue

                        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#terminateSession'
                        }
                    }

                    It 'Adds the other elements of the pipeline' {
                        @(
                            [PSCustomObject]@{
                                Name           = 'Name 1'
                                KeyDisplay     = 'Key Display 1'
                                FolderScanPath = '\\folder\path1'
                            }
                            [PSCustomObject]@{
                                Name              = 'Name 2'
                                KeyDisplay        = 'Key Display 2'
                                FolderScanPath    = '\\folder\path2'
                                FolderScanAccount = [pscredential]::new(
                                    'NewScanAccount',
                                    (ConvertTo-SecureString -String 'throws' -AsPlainText -Force)
                                )
                            }
                            [PSCustomObject]@{
                                Name           = 'Name 3'
                                KeyDisplay     = 'Key Display 3'
                                FolderScanPath = '\\folder\path3'
                            }
                        ) | Add-AddressBookEntry @commonParameters -ErrorAction SilentlyContinue

                        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjects' -and
                            $Body.OuterXml -eq ([xml]'<?xml version="1.0" encoding="utf-8"?>
                            <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                                <s:Body>
                                    <u:putObjects xmlns:xs="http://www.w3.org/2001/XMLSchema"
                                            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                            xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                            xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                            xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                                        <sessionId>12345</sessionId>
                                        <objectClass>entry</objectClass>
                                        <propListList xsi:type="soap-enc:Array" soap-enc:arrayType="ricoh:propertyList[]">
                                            <item>
                                                <item>
                                                    <propName>entryType</propName>
                                                    <propVal>user</propVal>
                                                </item>
                                                <item>
                                                    <propName>name</propName>
                                                    <propVal>Name 1</propVal>
                                                </item>
                                                <item>
                                                    <propName>longName</propName>
                                                    <propVal>Key Display 1</propVal>
                                                </item>
                                                <item>
                                                    <propName>displayedOrder</propName>
                                                    <propVal>5</propVal>
                                                </item>
                                                <item>
                                                    <propName>tagId</propName>
                                                    <propVal>1</propVal>
                                                </item>
                                                <item>
                                                    <propName>remoteFolder:</propName>
                                                    <propVal>true</propVal>
                                                </item>
                                                <item>
                                                    <propName>remoteFolder:type</propName>
                                                    <propVal>smb</propVal>
                                                </item>
                                                <item>
                                                    <propName>remoteFolder:path</propName>
                                                    <propVal>\\folder\path1</propVal>
                                                </item>
                                                <item>
                                                    <propName>remoteFolder:port</propName>
                                                    <propVal>21</propVal>
                                                </item>
                                                <item>
                                                    <propName>isDestination</propName>
                                                    <propVal>true</propVal>
                                                </item>
                                            </item>
                                            <item>
                                                <item>
                                                    <propName>entryType</propName>
                                                    <propVal>user</propVal>
                                                </item>
                                                <item>
                                                    <propName>name</propName>
                                                    <propVal>Name 3</propVal>
                                                </item>
                                                <item>
                                                    <propName>longName</propName>
                                                    <propVal>Key Display 3</propVal>
                                                </item>
                                                <item>
                                                    <propName>displayedOrder</propName>
                                                    <propVal>5</propVal>
                                                </item>
                                                <item>
                                                    <propName>tagId</propName>
                                                    <propVal>1</propVal>
                                                </item>
                                                <item>
                                                    <propName>remoteFolder:</propName>
                                                    <propVal>true</propVal>
                                                </item>
                                                <item>
                                                    <propName>remoteFolder:type</propName>
                                                    <propVal>smb</propVal>
                                                </item>
                                                <item>
                                                    <propName>remoteFolder:path</propName>
                                                    <propVal>\\folder\path3</propVal>
                                                </item>
                                                <item>
                                                    <propName>remoteFolder:port</propName>
                                                    <propVal>21</propVal>
                                                </item>
                                                <item>
                                                    <propName>isDestination</propName>
                                                    <propVal>true</propVal>
                                                </item>
                                            </item>
                                        </propListList>
                                    </u:putObjects>
                                </s:Body>
                            </s:Envelope>').OuterXml
                        }
                    }
                }
            }
        }
    }
}

Describe 'Get-AddressBookEntry' {
    It 'Has the correct parameters' {
        $command = Get-Command Get-AddressBookEntry

        $command | Should -HaveParameter Hostname -Mandatory -Type [uri]
        $command | Should -HaveParameter Credential -Mandatory -Type [pscredential]
        $command | Should -HaveParameter Id -Not -Mandatory
        $command | Should -HaveParameter Id -Type [uint32[]]
        $command | Should -HaveParameter Name -Not -Mandatory
        $command | Should -HaveParameter Name -Type [string]
        $command | Should -HaveParameter SkipCertificateCheck -Not -Mandatory
        $command | Should -HaveParameter SkipCertificateCheck -Type [switch]
    }

    Context 'Search-AddressBookEntry' {
        It 'Does not call searchObjects when -Id is given' {
            Get-AddressBookEntry @commonParameters -Id @()

            Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Times 0 -ParameterFilter {
                $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#searchObjects'
            }
        }

        It 'Calls searchObjects when -Id is not given' {
            Get-AddressBookEntry @commonParameters

            Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                $expected = [xml]'<?xml version="1.0" encoding="utf-8"?>
                <s:Envelope
                        xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                        s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                    <s:Body>
                        <u:searchObjects
                                xmlns:xs="http://www.w3.org/2001/XMLSchema"
                                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                            <sessionId>12345</sessionId>
                            <selectProps xsi:type="soap-enc:Array" soap-enc:arrayType="xs:string[]">
                                <item>id</item>
                            </selectProps>
                            <fromClass>entry</fromClass>
                            <orderBy xsi:type="soap-enc:Array" soap-enc:arrayType="ricoh:queryOrderBy[]">
                                <item>
                                    <propName>index</propName>
                                    <isDescending>false</isDescending>
                                </item>
                            </orderBy>
                            <rowOffset>0</rowOffset>
                            <rowCount>50</rowCount>
                        </u:searchObjects>
                    </s:Body>
                </s:Envelope>'

                $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#searchObjects' -and
                $Body.OuterXml -eq $expected.OuterXml
            }
        }

        Context 'When searchObjects returns more than 50 results' {
            BeforeEach {
                $max = 150
                $batchSize = 50
                Mock -ModuleName RicohAddressBook Invoke-WebRequest {
                    $response = [xml]'<?xml version="1.0" encoding="UTF-8"?>
                    <s:Envelope
                            xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                            s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                        <s:Body>
                            <tns:searchObjectsResponse
                                    xmlns:tns="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                                <returnValue>EOD</returnValue>
                                <resultSetId></resultSetId>
                                <numOfResults />
                                <rowList
                                        xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                        xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                        soap-enc:arrayType="itt:property[][]" />
                            </tns:searchObjectsResponse>
                        </s:Body>
                    </s:Envelope>'

                    $rowOffset = [uint32]$Body.Envelope.Body.searchObjects.rowOffset
                    $rowList = $response.Envelope.Body.searchObjectsResponse.rowList
                    $range = ($rowOffset + 1)..$max | Select-Object -First $batchSize

                    foreach ($index in $range) {
                        $item = $response.CreateElement('item', 'http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes')
                        $innerItem = $response.CreateElement('item')

                        $propName = $response.CreateElement('propName')
                        $propValue = $response.CreateElement('propValue')
                        $propName.InnerText = 'id'
                        $propValue.InnerText = $index

                        $innerItem.AppendChild($propName) > $null
                        $innerItem.AppendChild($propValue) > $null
                        $item.AppendChild($innerItem) > $null
                        $rowList.AppendChild($item) > $null
                    }

                    $response.Envelope.Body.searchObjectsResponse.numOfResults = [string]$max
                    if (($max - $rowOffset) -gt $batchSize) {
                        $response.Envelope.Body.searchObjectsResponse.returnValue = 'OK'
                    }

                    $response
                } -ParameterFilter {
                    $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#searchObjects'
                }
            }

            It 'Calls searchObjects multiple times with increasing rowOffset values' {
                Get-AddressBookEntry @commonParameters

                $WithRowOffset = {
                    param(
                        [uint32]
                        [Parameter(Mandatory)]
                        $RowOffset
                    )

                    $request = [xml]'<?xml version="1.0" encoding="utf-8"?>
                    <s:Envelope
                            xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                            s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                        <s:Body>
                            <u:searchObjects
                                    xmlns:xs="http://www.w3.org/2001/XMLSchema"
                                    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                    xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                    xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                    xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                                <sessionId>12345</sessionId>
                                <selectProps
                                        xsi:type="soap-enc:Array"
                                        soap-enc:arrayType="xs:string[]">
                                    <item>id</item>
                                </selectProps>
                                <fromClass>entry</fromClass>
                                <orderBy xsi:type="soap-enc:Array" soap-enc:arrayType="ricoh:queryOrderBy[]">
                                    <item>
                                        <propName>index</propName>
                                        <isDescending>false</isDescending>
                                    </item>
                                </orderBy>
                                <rowOffset />
                                <rowCount>50</rowCount>
                            </u:searchObjects>
                        </s:Body>
                    </s:Envelope>'

                    $request.Envelope.Body.searchObjects.rowOffset = [string]$RowOffset

                    $request
                }

                Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 3 -ParameterFilter {
                    $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#searchObjects'
                }

                Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                    $Body.OuterXml -eq (& $WithRowOffset -RowOffset 0).OuterXml
                }

                Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                    $Body.OuterXml -eq (& $WithRowOffset -RowOffset 50).OuterXml
                }

                Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                    $Body.OuterXml -eq (& $WithRowOffset -RowOffset 100).OuterXml
                }
            }
        }
    }

    Context 'When called with -Id' {
        It 'Does not call searchObjects' {
            Get-AddressBookEntry @commonParameters -Id 1

            Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Times 0 -ParameterFilter {
                $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#searchObjects'
            }
        }

        It 'Calls getObjectsProps with the given IDs' {
            Get-AddressBookEntry @commonParameters -Id (1..5)

            Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#getObjectsProps' -and
                $Body.OuterXml -eq ([xml]'<?xml version="1.0" encoding="utf-8"?>
                <s:Envelope
                        xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                        s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                    <s:Body>
                        <u:getObjectsProps
                                xmlns:xs="http://www.w3.org/2001/XMLSchema"
                                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                            <sessionId>12345</sessionId>
                            <objectIdList xsi:type="soap-enc:Array" soap-enc:arrayType="xs:string[]">
                                <item>entry:1</item>
                                <item>entry:2</item>
                                <item>entry:3</item>
                                <item>entry:4</item>
                                <item>entry:5</item>
                            </objectIdList>
                            <selectProps xsi:type="soap-enc:Array" soap-enc:arrayType="xs:string[]">
                                <item>id</item>
                                <item>index</item>
                                <item>name</item>
                                <item>longName</item>
                                <item>displayedOrder</item>
                                <item>tagId</item>
                                <item>auth:</item>
                                <item>auth:name</item>
                                <item>lastAccessDateTime</item>
                                <item>remoteFolder:</item>
                                <item>remoteFolder:type</item>
                                <item>remoteFolder:path</item>
                                <item>remoteFolder:port</item>
                                <item>remoteFolder:accountName</item>
                                <item>mail:</item>
                                <item>mail:address</item>
                                <item>isSender</item>
                                <item>isDestination</item>
                            </selectProps>
                        </u:getObjectsProps>
                    </s:Body>
                </s:Envelope>').OuterXml
            }
        }
    }

    Context 'When called without -Id' {
        BeforeAll {
            Mock -ModuleName RicohAddressBook Search-AddressBookEntry { 1..10 }
        }

        It 'Calls getObjectsProps with the IDs returned from Search-AddressBookEntry' {
            Get-AddressBookEntry @commonParameters

            Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#getObjectsProps' -and
                $Body.OuterXml -eq ([xml]'<?xml version="1.0" encoding="utf-8"?>
                <s:Envelope
                        xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                        s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                    <s:Body>
                        <u:getObjectsProps
                                xmlns:xs="http://www.w3.org/2001/XMLSchema"
                                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                            <sessionId>12345</sessionId>
                            <objectIdList xsi:type="soap-enc:Array" soap-enc:arrayType="xs:string[]">
                                <item>entry:1</item>
                                <item>entry:2</item>
                                <item>entry:3</item>
                                <item>entry:4</item>
                                <item>entry:5</item>
                                <item>entry:6</item>
                                <item>entry:7</item>
                                <item>entry:8</item>
                                <item>entry:9</item>
                                <item>entry:10</item>
                            </objectIdList>
                            <selectProps xsi:type="soap-enc:Array" soap-enc:arrayType="xs:string[]">
                                <item>id</item>
                                <item>index</item>
                                <item>name</item>
                                <item>longName</item>
                                <item>displayedOrder</item>
                                <item>tagId</item>
                                <item>auth:</item>
                                <item>auth:name</item>
                                <item>lastAccessDateTime</item>
                                <item>remoteFolder:</item>
                                <item>remoteFolder:type</item>
                                <item>remoteFolder:path</item>
                                <item>remoteFolder:port</item>
                                <item>remoteFolder:accountName</item>
                                <item>mail:</item>
                                <item>mail:address</item>
                                <item>isSender</item>
                                <item>isDestination</item>
                            </selectProps>
                        </u:getObjectsProps>
                    </s:Body>
                </s:Envelope>').OuterXml
            }
        }
    }

    Context 'When called with more than 50 IDs' {
        BeforeAll {
            $range = 1..100

            Mock -ModuleName RicohAddressBook Search-AddressBookEntry { $range }
        }

        Context -ForEach @(
            @{ Label = 'With'; Arguments = @{Id = $range } }
            @{ Label = 'Without'; Arguments = @{} }
        ) '<Label> -Id' {
            It 'Calls getObjectsProps in batches of 50' {
                Get-AddressBookEntry @commonParameters @Arguments

                $expected = [xml]'<?xml version="1.0" encoding="utf-8"?>
            <s:Envelope
                    xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                <s:Body>
                    <u:getObjectsProps
                            xmlns:xs="http://www.w3.org/2001/XMLSchema"
                            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                            xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                            xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                            xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                        <sessionId>12345</sessionId>
                        <objectIdList xsi:type="soap-enc:Array" soap-enc:arrayType="xs:string[]" />
                        <selectProps xsi:type="soap-enc:Array" soap-enc:arrayType="xs:string[]">
                            <item>id</item>
                            <item>index</item>
                            <item>name</item>
                            <item>longName</item>
                            <item>displayedOrder</item>
                            <item>tagId</item>
                            <item>auth:</item>
                            <item>auth:name</item>
                            <item>lastAccessDateTime</item>
                            <item>remoteFolder:</item>
                            <item>remoteFolder:type</item>
                            <item>remoteFolder:path</item>
                            <item>remoteFolder:port</item>
                            <item>remoteFolder:accountName</item>
                            <item>mail:</item>
                            <item>mail:address</item>
                            <item>isSender</item>
                            <item>isDestination</item>
                        </selectProps>
                    </u:getObjectsProps>
                </s:Body>
            </s:Envelope>'

                Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 2 -ParameterFilter {
                    $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#getObjectsProps'
                }

                function Get-Expected {
                    param(
                        [uint32[]]
                        [Parameter(Mandatory)]
                        $range
                    )

                    [xml] $result = $expected.Clone()
                    foreach ($id in $range) {
                        $item = $result.CreateElement('item')
                        $item.InnerText = "entry:$id"
                        $result.Envelope.Body.getObjectsProps.objectIdList.AppendChild($item) > $null
                    }

                    $result.OuterXml
                }

                Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                    $Body.OuterXml -eq (Get-Expected (1..50))
                }

                Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
                    $Body.OuterXml -eq (Get-Expected (51..100))
                }
            }
        }
    }

    Context 'When called with -Name' {
        It 'Filters out the results that do not match' {
            $results = Get-AddressBookEntry @commonParameters -Name 'Matt*'

            $results | Should -HaveCount 0
        }

        It 'Includes the results that match' {
            $results = Get-AddressBookEntry @commonParameters -Name 'John*'

            $results | Should -HaveCount 1
        }
    }

    It 'Returns an object of the correct type' {
        $result = Get-AddressBookEntry @commonParameters

        $result.pstypenames[0] | Should -Be 'Ricoh.AddressBook.Entry'
    }

    It 'Returns an object with the parsed values' {
        $result = Get-AddressBookEntry @commonParameters

        $result | Should -MatchObject @([PSCustomObject]@{
                PSTypeName         = 'Ricoh.AddressBook.Entry'
                ID                 = 1
                RegistrationNumber = 1
                DisplayPriority    = 5
                Name               = 'John D'
                KeyDisplay         = 'John Doe'
                Frequent           = $true
                Title1             = 'IJK'
                Title2             = 2
                Title3             = 3
                LastUsed           = [datetime]'2023-12-21T13:26:18Z'
                UserCode           = '54321'
                FolderScanType     = 'smb'
                FolderScanPath     = '\\folder\path'
                FolderScanPort     = 21
                FolderScanAccount  = 'ScanAccount'
                EmailAddress       = 'john.doe@example.com'
                IsSender           = $false
                IsDestination      = $true
            })
    }

    It 'Does not contain values that are missing' {
        Mock -ModuleName RicohAddressBook Invoke-WebRequest {
            [xml]'<?xml version="1.0" encoding="UTF-8"?>
            <s:Envelope
                    xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                <s:Body>
                    <tns:getObjectsPropsResponse
                            xmlns:tns="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                        <returnValue
                                xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                soap-enc:arrayType="itt:property[][1]">
                            <item
                                    xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                    xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                    soap-enc:arrayType="itt:property[14]">
                                <item>
                                    <propName>id</propName>
                                    <propVal>1</propVal>
                                </item>
                                <item>
                                    <propName>index</propName>
                                    <propVal>1</propVal>
                                </item>
                                <item>
                                    <propName>displayedOrder</propName>
                                    <propVal>5</propVal>
                                </item>
                                <item>
                                    <propName>name</propName>
                                    <propVal>John D</propVal>
                                </item>
                                <item>
                                    <propName>longName</propName>
                                    <propVal>John Doe</propVal>
                                </item>
                                <item>
                                    <propName>tagId</propName>
                                    <propVal>1</propVal>
                                </item>
                                <item>
                                    <propName>lastAccessDateTime</propName>
                                    <propVal>1970-01-01T00:00:00Z</propVal>
                                </item>
                                <item>
                                    <propName>isDestination</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>isSender</propName>
                                    <propVal>false</propVal>
                                </item>
                                <item>
                                    <propName>auth:</propName>
                                    <propVal>false</propVal>
                                </item>
                                <item>
                                    <propName>auth:name</propName>
                                    <propVal>54321</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:</propName>
                                    <propVal>false</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:type</propName>
                                    <propVal>smb</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:path</propName>
                                    <propVal>\\folder\path</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:port</propName>
                                    <propVal>21</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:accountName</propName>
                                    <propVal>ScanAccount</propVal>
                                </item>
                                <item>
                                    <propName>mail:</propName>
                                    <propVal>false</propVal>
                                </item>
                                <item>
                                    <propName>mail:address</propName>
                                    <propVal>john.doe@example.com</propVal>
                                </item>
                            </item>
                        </returnValue>
                    </tns:getObjectsPropsResponse>
                </s:Body>
            </s:Envelope>'
        } -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#getObjectsProps'
        }

        $result = Get-AddressBookEntry @commonParameters
        $properties = $result.psobject.Properties.Name

        $properties | Should -Not -Contain 'Title1'
        $properties | Should -Not -Contain 'Title2'
        $properties | Should -Not -Contain 'Title3'
        $properties | Should -Not -Contain 'LastUsed'
        $properties | Should -Not -Contain 'UserCode'
        $properties | Should -Not -Contain 'RemoteFolderType'
        $properties | Should -Not -Contain 'RemoteFolderPath'
        $properties | Should -Not -Contain 'RemoteFolderPort'
        $properties | Should -Not -Contain 'RemoteFolderAccount'
        $properties | Should -Not -Contain 'EmailAddress'
    }

    It 'Has a false Frequent value if tagId does not contain 1' {
        Mock -ModuleName RicohAddressBook Invoke-WebRequest {
            [xml]'<?xml version="1.0" encoding="UTF-8"?>
            <s:Envelope
                    xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                <s:Body>
                    <tns:getObjectsPropsResponse
                            xmlns:tns="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                        <returnValue
                                xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                soap-enc:arrayType="itt:property[][1]">
                            <item
                                    xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                                    xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                                    soap-enc:arrayType="itt:property[14]">
                                <item>
                                    <propName>id</propName>
                                    <propVal>1</propVal>
                                </item>
                                <item>
                                    <propName>index</propName>
                                    <propVal>1</propVal>
                                </item>
                                <item>
                                    <propName>displayedOrder</propName>
                                    <propVal>5</propVal>
                                </item>
                                <item>
                                    <propName>name</propName>
                                    <propVal>John D</propVal>
                                </item>
                                <item>
                                    <propName>longName</propName>
                                    <propVal>John Doe</propVal>
                                </item>
                                <item>
                                    <propName>tagId</propName>
                                    <propVal>6</propVal>
                                </item>
                                <item>
                                    <propName>lastAccessDateTime</propName>
                                    <propVal>1970-01-01T00:00:00Z</propVal>
                                </item>
                                <item>
                                    <propName>isDestination</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>isSender</propName>
                                    <propVal>false</propVal>
                                </item>
                            </item>
                        </returnValue>
                    </tns:getObjectsPropsResponse>
                </s:Body>
            </s:Envelope>'
        } -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#getObjectsProps'
        }

        $result = Get-AddressBookEntry @commonParameters

        $result.Frequent | Should -BeFalse
    }
}

Describe 'Update-AddressBookEntry' {
    It 'Has the correct parameters' {
        $command = Get-Command Update-AddressBookEntry
        $tagId = & (Get-Module RicohAddressBook) { [TagId] }

        $command | Should -HaveParameter Hostname -Mandatory -Type [uri]
        $command | Should -HaveParameter Credential -Mandatory -Type [pscredential]
        $command | Should -HaveParameter Id -Mandatory -Type [uint32]
        $command | Should -HaveParameter Name -Not -Mandatory
        $command | Should -HaveParameter Name -Type [string]
        $command | Should -HaveParameter KeyDisplay -Not -Mandatory
        $command | Should -HaveParameter KeyDisplay -Type [string] -Alias LongName
        $command | Should -HaveParameter DisplayPriority -Not -Mandatory
        $command | Should -HaveParameter DisplayPriority -Type [byte]
        $command | Should -HaveParameter Frequent -Not -Mandatory
        $command | Should -HaveParameter Frequent -Type [switch]
        $command | Should -HaveParameter Title1 -Not -Mandatory
        $command | Should -HaveParameter Title1 -Type $tagId
        $command | Should -HaveParameter Title2 -Not -Mandatory
        $command | Should -HaveParameter Title2 -Type [byte]
        $command | Should -HaveParameter Title3 -Not -Mandatory
        $command | Should -HaveParameter Title3 -Type [byte]
        $command | Should -HaveParameter UserCode -Not -Mandatory
        $command | Should -HaveParameter UserCode -Type [string]
        $command | Should -HaveParameter FolderScanPath -Not -Mandatory
        $command | Should -HaveParameter FolderScanPath -Type [string]
        $command | Should -HaveParameter FolderScanAccount -Not -Mandatory
        $command | Should -HaveParameter FolderScanAccount -Type [pscredential]
        $command | Should -HaveParameter EmailAddress -Not -Mandatory
        $command | Should -HaveParameter EmailAddress -Type [string]
        $command | Should -HaveParameter IsSender -Not -Mandatory
        $command | Should -HaveParameter IsSender -Type [nullable[bool]]
        $command | Should -HaveParameter IsDestination -Not -Mandatory
        $command | Should -HaveParameter IsDestination -Type [nullable[bool]]
        $command | Should -HaveParameter ForceUserCode -Not -Mandatory
        $command | Should -HaveParameter ForceUserCode -Type [switch]
        $command | Should -HaveParameter ForceFolderScanPath -Not -Mandatory
        $command | Should -HaveParameter ForceFolderScanPath -Type [switch]
        $command | Should -HaveParameter ForceFolderScanAccount -Not -Mandatory
        $command | Should -HaveParameter ForceFolderScanAccount -Type [switch]
        $command | Should -HaveParameter ForceEmailAddress -Not -Mandatory
        $command | Should -HaveParameter ForceEmailAddress -Type [switch]
        $command | Should -HaveParameter PassThru -Not -Mandatory
        $command | Should -HaveParameter PassThru -Type [switch]
        $command | Should -HaveParameter SkipCertificateCheck -Not -Mandatory
        $command | Should -HaveParameter SkipCertificateCheck -Type [switch]
    }

    It 'Updates each element of the pipeline separately' {
        @(
            [PSCustomObject]@{
                Id   = 1
                Name = 'New Name'
            }
            [PSCustomObject]@{
                Id         = 2
                KeyDisplay = 'New Key Display'
            }
            [PSCustomObject]@{
                Id                  = 3
                FolderScanPath      = '\\new\folder\path'
                ForceFolderScanPath = $true
            }
            [PSCustomObject]@{
                Id                     = 4
                FolderScanAccount      = [pscredential]::new(
                    'NewScanAccount',
                    (ConvertTo-SecureString -String 'NewMockPassword' -AsPlainText -Force)
                )
                ForceFolderScanAccount = $true
            }
            [PSCustomObject]@{
                Id                = 5
                EmailAddress      = 'new@example.com'
                ForceEmailAddress = $true
            }
            [PSCustomObject]@{
                Id       = 6
                IsSender = $true
            }
            [PSCustomObject]@{
                Id       = 7
                IsSender = $false
            }
            [PSCustomObject]@{
                Id            = 8
                IsDestination = $true
            }
            [PSCustomObject]@{
                Id            = 9
                IsDestination = $false
            }
            [PSCustomObject]@{
                Id              = 10
                DisplayPriority = 1
            }
            [PSCustomObject]@{
                Id       = 11
                Frequent = $true
            }
            [PSCustomObject]@{
                Id     = 12
                Title1 = 'AB'
            }
            [PSCustomObject]@{
                Id     = 13
                Title2 = 10
            }
            [PSCustomObject]@{
                Id     = 14
                Title3 = 5
            }
            [PSCustomObject]@{
                Id            = 15
                UserCode      = '54321'
                ForceUserCode = $true
            }
        ) | Update-AddressBookEntry @commonParameters

        function Get-Expected {
            param(
                [uint32]
                [Parameter(Mandatory)]
                $Id,

                [System.Collections.IDictionary]
                $Properties
            )

            $expected = [xml]'<?xml version="1.0" encoding="utf-8"?>
            <s:Envelope
                    xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                <s:Body>
                    <u:putObjectProps
                            xmlns:xs="http://www.w3.org/2001/XMLSchema"
                            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                            xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                            xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                            xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                        <sessionId>12345</sessionId>
                        <objectId />
                        <propList xsi:type="soap-enc:Array" soap-enc:arrayType="ricoh:property[]" />
                    </u:putObjectProps>
                </s:Body>
            </s:Envelope>'
            $expected.Envelope.Body.putObjectProps.objectId = "entry:$Id"

            foreach ($pair in $Properties.GetEnumerator()) {
                $item = $expected.CreateElement('item')
                $propertyName = $expected.CreateElement('propName')
                $propValue = $expected.CreateElement('propVal')

                $propertyName.InnerText = $pair.Key
                $propValue.InnerText = $pair.Value

                $item.AppendChild($propertyName) > $null
                $item.AppendChild($propValue) > $null
                $expected.Envelope.Body.putObjectProps.propList.AppendChild($item) > $null
            }

            $expected.OuterXml
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 1 @{ name = 'New Name' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 2 @{ longName = 'New Key Display' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 3 ([ordered]@{
                        'remoteFolder:'     = 'true'
                        'remoteFolder:type' = 'smb'
                        'remoteFolder:path' = '\\new\folder\path'
                        'remoteFolder:port' = 21
                    })
            )
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 4 ([ordered]@{
                        'remoteFolder:select'      = 'private'
                        'remoteFolder:accountName' = 'NewScanAccount'
                        'remoteFolder:password'    = 'TmV3TW9ja1Bhc3N3b3Jk'
                    }))
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 5 ([ordered]@{
                        'mail:'        = 'true'
                        'mail:address' = 'new@example.com'
                    }))
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 6 @{ isSender = 'true' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 7 @{ isSender = 'false' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 8 @{ isDestination = 'true' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 9 @{ isDestination = 'false' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 10 @{ displayedOrder = '1' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 11 @{ tagId = '1' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 12 @{ tagId = '2' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 13 @{ tagId = '21' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 14 @{ tagId = '26' })
        }

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 15 ([ordered]@{
                        'auth:'     = 'true'
                        'auth:name' = '54321'
                    }))
        }
    }

    It 'Outputs nothing if -PassThru is not provided' {
        $result = Update-AddressBookEntry @commonParameters -Id 1

        $result | Should -BeNullOrEmpty
    }

    It 'Outputs an object with the correct values if -PassThru is provided' {
        $actual = @(
            [PSCustomObject]@{
                Id                = 1
                Name              = 'New Name'
                KeyDisplay        = 'New Key'
                DisplayPriority   = 4
                Frequent          = $true
                Title1            = 'LMN'
                Title2            = 2
                Title3            = 3
                UserCode          = '01234'
                FolderScanPath    = '\\new\folder\path'
                FolderScanAccount = [pscredential]::new(
                    'NewScanAccount',
                    (ConvertTo-SecureString -String 'NewMockPassword' -AsPlainText -Force)
                )
                EmailAddress      = 'new@example.com'
                IsSender          = $true
                IsDestination     = $true
            }
            [PSCustomObject]@{
                Id                = 2
                ForceUserCode     = $true
                ForceScanPath     = $true
                ForceScanAccount  = $true
                ForceEmailAddress = $true
            }
        ) | Update-AddressBookEntry @commonParameters -PassThru

        $actual | Should -MatchObject @(
            [PSCustomObject]@{
                PSTypeName        = 'Ricoh.AddressBook.Entry'
                ID                = 1
                Name              = 'New Name'
                KeyDisplay        = 'New Key'
                DisplayPriority   = 4
                Frequent          = $true
                Title1            = 'LMN'
                Title2            = 2
                Title3            = 3
                UserCode          = '01234'
                FolderScanType    = 'smb'
                FolderScanPath    = '\\new\folder\path'
                FolderScanPort    = 21
                FolderScanAccount = 'NewScanAccount'
                EmailAddress      = 'new@example.com'
                IsSender          = $true
                IsDestination     = $true
            }
            [PSCustomObject]@{
                PSTypeName = 'Ricoh.AddressBook.Entry'
                ID         = 2
            }
        )
    }

    It -ForEach @(
        @{ Switch = 'ForceUserCode'; Result = @{ 'auth:' = 'false' } }
        @{ Switch = 'ForceFolderScanPath'; Result = @{ 'remoteFolder:' = 'false' } }
        @{ Switch = 'ForceFolderScanAccount'; Result = @{ 'remoteFolder:select' = '' } }
        @{ Switch = 'ForceEmailAddress'; Result = @{ 'mail:' = 'false' } }
    ) 'Disables the proper settings when <Switch> is $true' {
        function Get-Expected {
            param(
                [uint32]
                [Parameter(Mandatory)]
                $Id,

                [System.Collections.IDictionary]
                $Properties
            )

            $expected = [xml]'<?xml version="1.0" encoding="utf-8"?>
            <s:Envelope
                    xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                <s:Body>
                    <u:putObjectProps
                            xmlns:xs="http://www.w3.org/2001/XMLSchema"
                            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                            xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                            xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                            xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                        <sessionId>12345</sessionId>
                        <objectId />
                        <propList xsi:type="soap-enc:Array" soap-enc:arrayType="ricoh:property[]" />
                    </u:putObjectProps>
                </s:Body>
            </s:Envelope>'
            $expected.Envelope.Body.putObjectProps.objectId = "entry:$Id"

            foreach ($pair in $Properties.GetEnumerator()) {
                $item = $expected.CreateElement('item')
                $propertyName = $expected.CreateElement('propName')
                $propValue = $expected.CreateElement('propVal')

                $propertyName.InnerText = $pair.Key
                $propValue.InnerText = $pair.Value

                $item.AppendChild($propertyName) > $null
                $item.AppendChild($propValue) > $null
                $expected.Envelope.Body.putObjectProps.propList.AppendChild($item) > $null
            }

            $expected.OuterXml
        }

        [PSCustomObject]@{
            Id        = 1
            ($Switch) = $true
        } | Update-AddressBookEntry @commonParameters

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq (Get-Expected -Id 1 $Result)
        }
    }

    It 'The -Force* switches do not reset values further along the pipeline' {
        @(
            [PSCustomObject]@{
                Id                     = 1
                ForceUserCode          = $true
                ForceFolderScanPath    = $true
                ForceFolderScanAccount = $true
                ForceEmailAddress      = $true
            }
            [PSCustomObject]@{
                Id = 2
            }
        ) | Update-AddressBookEntry @commonParameters

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $expected = [xml]'<?xml version="1.0" encoding="utf-8"?>
            <s:Envelope
                    xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                <s:Body>
                    <u:putObjectProps
                            xmlns:xs="http://www.w3.org/2001/XMLSchema"
                            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                            xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                            xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                            xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                        <sessionId>12345</sessionId>
                        <objectId>entry:2</objectId>
                        <propList xsi:type="soap-enc:Array" soap-enc:arrayType="ricoh:property[]" />
                    </u:putObjectProps>
                </s:Body>
            </s:Envelope>'

            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjectProps' -and
            $Body.OuterXml -eq $expected.OuterXml
        }
    }
}

Describe 'Add-AddressBookEntry' {
    It 'Has the correct parameters' {
        $command = Get-Command Add-AddressBookEntry
        $tagId = & (Get-Module RicohAddressBook) { [TagId] }

        $command | Should -HaveParameter Hostname -Mandatory -Type [uri]
        $command | Should -HaveParameter Credential -Mandatory -Type [pscredential]
        $command | Should -HaveParameter Name -Mandatory -Type [string]
        $command | Should -HaveParameter KeyDisplay -Mandatory -Type [string] -Alias LongName
        $command | Should -HaveParameter DisplayPriority -Not -Mandatory
        $command | Should -HaveParameter DisplayPriority -Type [byte] -DefaultValue 5
        $command | Should -HaveParameter Frequent -Not -Mandatory
        $command | Should -HaveParameter Frequent -Type [bool] -DefaultValue '$true'
        $command | Should -HaveParameter Title1 -Not -Mandatory
        $command | Should -HaveParameter Title1 -Type $tagId
        $command | Should -HaveParameter Title2 -Not -Mandatory
        $command | Should -HaveParameter Title2 -Type [byte]
        $command | Should -HaveParameter Title3 -Not -Mandatory
        $command | Should -HaveParameter Title3 -Type [byte]
        $command | Should -HaveParameter UserCode -Not -Mandatory
        $command | Should -HaveParameter UserCode -Type [string]
        $command | Should -HaveParameter FolderScanPath -Mandatory -Type [string] -InParameterSet 'Folder'
        $command | Should -HaveParameter FolderScanPath -Mandatory -Type [string] -InParameterSet 'FolderAndEmail'
        $command | Should -HaveParameter FolderScanAccount -Type [pscredential] -InParameterSet 'Folder'
        $command | Should -HaveParameter FolderScanAccount -Type [pscredential] -InParameterSet 'FolderAndEmail'
        $command | Should -HaveParameter EmailAddress -Mandatory -Type [string] -InParameterSet 'Email'
        $command | Should -HaveParameter EmailAddress -Mandatory -Type [string] -InParameterSet 'FolderAndEmail'
        $command | Should -HaveParameter IsSender -Type [nullable[bool]] -InParameterSet 'Email'
        $command | Should -HaveParameter IsSender -Type [nullable[bool]] -InParameterSet 'FolderAndEmail'
        $command | Should -HaveParameter IsDestination -Type [nullable[bool]]
        $command | Should -HaveParameter PassThru -Not -Mandatory
        $command | Should -HaveParameter PassThru -Type [switch]
        $command | Should -HaveParameter SkipCertificateCheck -Not -Mandatory
        $command | Should -HaveParameter SkipCertificateCheck -Type [switch]
    }

    It 'Adds each element of the pipeline separately' {
        @(
            [PSCustomObject]@{
                Name              = 'By Folder Path'
                KeyDisplay        = 'Folder'
                DisplayPriority   = 4
                Frequent          = $true
                Title1            = 'CD'
                UserCode          = 98765
                FolderScanPath    = '\\folder\path'
                FolderScanAccount = [pscredential]::new(
                    'ScanAccount',
                    (ConvertTo-SecureString -String 'MockPassword' -AsPlainText -Force)
                )
            }
            [PSCustomObject]@{
                Name         = 'By Email Address'
                KeyDisplay   = 'Email'
                Frequent     = $false
                Title2       = 9
                Title3       = 4
                EmailAddress = 'email@example.com'
            }
            [PSCustomObject]@{
                Name              = 'By Folder & Email'
                KeyDisplay        = 'Folder/Email'
                FolderScanPath    = '\\second\folder\path'
                FolderScanAccount = [pscredential]::new(
                    'ScanAccount2',
                    (ConvertTo-SecureString -String 'MockPassword2' -AsPlainText -Force)
                )
                EmailAddress      = 'email2@example.com'
                IsSender          = $true
                IsDestination     = $true
            }
            [PSCustomObject]@{
                Name           = 'Without ScanAccount'
                KeyDisplay     = 'No ScanAccount'
                Frequent       = $true
                Title1         = 'EF'
                FolderScanPath = '\\folder\path'
            }
            [PSCustomObject]@{
                Name          = 'Not Destination'
                KeyDisplay    = 'Not Destination'
                Frequent      = $false
                Title2        = 2
                EmailAddress  = 'email@example.com'
                IsDestination = $false
            }
            [PSCustomObject]@{
                Name         = 'With Default Values'
                KeyDisplay   = 'With Defaults'
                Title3       = 3
                EmailAddress = 'email@example.com'
            }
        ) | Add-AddressBookEntry @commonParameters

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $expected = [xml]'<?xml version="1.0" encoding="utf-8"?>
            <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                <s:Body>
                    <u:putObjects
                            xmlns:xs="http://www.w3.org/2001/XMLSchema"
                            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                            xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                            xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                            xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                        <sessionId>12345</sessionId>
                        <objectClass>entry</objectClass>
                        <propListList xsi:type="soap-enc:Array" soap-enc:arrayType="ricoh:propertyList[]">
                            <item>
                                <item>
                                    <propName>entryType</propName>
                                    <propVal>user</propVal>
                                </item>
                                <item>
                                    <propName>name</propName>
                                    <propVal>By Folder Path</propVal>
                                </item>
                                <item>
                                    <propName>longName</propName>
                                    <propVal>Folder</propVal>
                                </item>
                                <item>
                                    <propName>displayedOrder</propName>
                                    <propVal>4</propVal>
                                </item>
                                <item>
                                    <propName>tagId</propName>
                                    <propVal>1,3</propVal>
                                </item>
                                <item>
                                    <propName>auth:</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>auth:name</propName>
                                    <propVal>98765</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:type</propName>
                                    <propVal>smb</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:path</propName>
                                    <propVal>\\folder\path</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:port</propName>
                                    <propVal>21</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:select</propName>
                                    <propVal>private</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:accountName</propName>
                                    <propVal>ScanAccount</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:password</propName>
                                    <propVal>TW9ja1Bhc3N3b3Jk</propVal>
                                </item>
                                <item>
                                    <propName>isDestination</propName>
                                    <propVal>true</propVal>
                                </item>
                            </item>
                            <item>
                                <item>
                                    <propName>entryType</propName>
                                    <propVal>user</propVal>
                                </item>
                                <item>
                                    <propName>name</propName>
                                    <propVal>By Email Address</propVal>
                                </item>
                                <item>
                                    <propName>longName</propName>
                                    <propVal>Email</propVal>
                                </item>
                                <item>
                                    <propName>displayedOrder</propName>
                                    <propVal>5</propVal>
                                </item>
                                <item>
                                    <propName>tagId</propName>
                                    <propVal>20,25</propVal>
                                </item>
                                <item>
                                    <propName>mail:</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>mail:address</propName>
                                    <propVal>email@example.com</propVal>
                                </item>
                                <item>
                                    <propName>isSender</propName>
                                    <propVal>false</propVal>
                                </item>
                                <item>
                                    <propName>isDestination</propName>
                                    <propVal>true</propVal>
                                </item>
                            </item>
                            <item>
                                <item>
                                    <propName>entryType</propName>
                                    <propVal>user</propVal>
                                </item>
                                <item>
                                    <propName>name</propName>
                                    <propVal>By Folder &amp; Email</propVal>
                                </item>
                                <item>
                                    <propName>longName</propName>
                                    <propVal>Folder/Email</propVal>
                                </item>
                                <item>
                                    <propName>displayedOrder</propName>
                                    <propVal>5</propVal>
                                </item>
                                <item>
                                    <propName>tagId</propName>
                                    <propVal>1</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:type</propName>
                                    <propVal>smb</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:path</propName>
                                    <propVal>\\second\folder\path</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:port</propName>
                                    <propVal>21</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:select</propName>
                                    <propVal>private</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:accountName</propName>
                                    <propVal>ScanAccount2</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:password</propName>
                                    <propVal>TW9ja1Bhc3N3b3JkMg==</propVal>
                                </item>
                                <item>
                                    <propName>mail:</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>mail:address</propName>
                                    <propVal>email2@example.com</propVal>
                                </item>
                                <item>
                                    <propName>isSender</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>isDestination</propName>
                                    <propVal>true</propVal>
                                </item>
                            </item>
                            <item>
                                <item>
                                    <propName>entryType</propName>
                                    <propVal>user</propVal>
                                </item>
                                <item>
                                    <propName>name</propName>
                                    <propVal>Without ScanAccount</propVal>
                                </item>
                                <item>
                                    <propName>longName</propName>
                                    <propVal>No ScanAccount</propVal>
                                </item>
                                <item>
                                    <propName>displayedOrder</propName>
                                    <propVal>5</propVal>
                                </item>
                                <item>
                                    <propName>tagId</propName>
                                    <propVal>1,4</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:type</propName>
                                    <propVal>smb</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:path</propName>
                                    <propVal>\\folder\path</propVal>
                                </item>
                                <item>
                                    <propName>remoteFolder:port</propName>
                                    <propVal>21</propVal>
                                </item>
                                <item>
                                    <propName>isDestination</propName>
                                    <propVal>true</propVal>
                                </item>
                            </item>
                            <item>
                                <item>
                                    <propName>entryType</propName>
                                    <propVal>user</propVal>
                                </item>
                                <item>
                                    <propName>name</propName>
                                    <propVal>Not Destination</propVal>
                                </item>
                                <item>
                                    <propName>longName</propName>
                                    <propVal>Not Destination</propVal>
                                </item>
                                <item>
                                    <propName>displayedOrder</propName>
                                    <propVal>5</propVal>
                                </item>
                                <item>
                                    <propName>tagId</propName>
                                    <propVal>13</propVal>
                                </item>
                                <item>
                                    <propName>mail:</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>mail:address</propName>
                                    <propVal>email@example.com</propVal>
                                </item>
                                <item>
                                    <propName>isSender</propName>
                                    <propVal>false</propVal>
                                </item>
                                <item>
                                    <propName>isDestination</propName>
                                    <propVal>false</propVal>
                                </item>
                            </item>
                            <item>
                                <item>
                                    <propName>entryType</propName>
                                    <propVal>user</propVal>
                                </item>
                                <item>
                                    <propName>name</propName>
                                    <propVal>With Default Values</propVal>
                                </item>
                                <item>
                                    <propName>longName</propName>
                                    <propVal>With Defaults</propVal>
                                </item>
                                <item>
                                    <propName>displayedOrder</propName>
                                    <propVal>5</propVal>
                                </item>
                                <item>
                                    <propName>tagId</propName>
                                    <propVal>1,24</propVal>
                                </item>
                                <item>
                                    <propName>mail:</propName>
                                    <propVal>true</propVal>
                                </item>
                                <item>
                                    <propName>mail:address</propName>
                                    <propVal>email@example.com</propVal>
                                </item>
                                <item>
                                    <propName>isSender</propName>
                                    <propVal>false</propVal>
                                </item>
                                <item>
                                    <propName>isDestination</propName>
                                    <propVal>true</propVal>
                                </item>
                            </item>
                        </propListList>
                    </u:putObjects>
                </s:Body>
            </s:Envelope>'

            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putobjects' -and
            $Body.OuterXml -eq $expected.OuterXml
        }
    }

    It 'Throws an error if Frequent, Title1, Title2, and Title3 are not provided' {
        $incompleteEntry = @{
            Name           = 'Name'
            KeyDisplay     = 'KeyDisplay'
            FolderScanPath = '\\folder\path'
            Frequent       = $false
        }

        # Confirm that it is not a terminating error
        { Add-AddressBookEntry @commonParameters @incompleteEntry -ErrorAction SilentlyContinue } |
            Should -Not -Throw

        # Then if it throws with -ErrorAction 'Stop', it must be a
        # non-terminating error
        { Add-AddressBookEntry @commonParameters @incompleteEntry -ErrorAction Stop 2> $null } |
            Should -Throw
    }

    It -ForEach @(
        @{ Frequent = $true }
        @{ Title1 = 'AB' }
        @{ Title2 = 1 }
        @{ Title3 = 1 }
        @{ Frequent = $false; Title1 = 'CD' }
        @{ Frequent = $true; Title1 = 'EF' }
        @{ Frequent = $true; Title2 = 2 }
        @{ Frequent = $true; Title3 = 2 }
        @{ Title1 = 'GH'; Title2 = 3 }
        @{ Title1 = 'IJK'; Title3 = 3 }
        @{ Title2 = 4; Title3 = 4 }
    ) 'Does not throw if Frequent = <Frequent>, Title1 = <Title1>, Title2 = <Title2>, Title3 = <Title3>' {
        $entry = @{
            Name           = 'Name'
            KeyDisplay     = 'KeyDisplay'
            FolderScanPath = '\\folder\path'
        }

        foreach ($name in @('Frequent', 'Title1', 'Title2', 'Title3')) {
            $value = Get-Variable -Name $name -ValueOnly -ErrorAction SilentlyContinue
            if ($null -ne $value) {
                $entry[$name] = $value
            }
        }

        { Add-AddressBookEntry @commonParameters @entry } | Should -Not -Throw
    }

    It 'Outputs nothing if -PassThru is not provided' {
        $result = Add-AddressBookEntry @commonParameters -Name 'Name' -KeyDisplay 'KeyDisplay' -FolderScanPath '\\folder\path'

        $result | Should -BeNullOrEmpty
    }

    It 'Outputs an object with the correct values if -PassThru is provided' {
        $actual = @(
            [PSCustomObject]@{
                Name              = 'Name 1'
                KeyDisplay        = 'Key 1'
                DisplayPriority   = 6
                Frequent          = $false
                Title1            = 'IJK'
                Title2            = 9
                Title3            = 4
                UserCode          = '98765'
                FolderScanPath    = '\\folder\path'
                FolderScanAccount = [pscredential]::new(
                    'ScanAccount',
                    (ConvertTo-SecureString -String 'MockPassword' -AsPlainText -Force)
                )
            }
            [PSCustomObject]@{
                Name         = 'Name 2'
                KeyDisplay   = 'Key 2'
                EmailAddress = 'mail@example.com'
            }
        ) | Add-AddressBookEntry @commonParameters -PassThru

        $actual | Should -MatchObject @(
            [PSCustomObject]@{
                PSTypeName        = 'Ricoh.AddressBook.Entry'
                Name              = 'Name 1'
                KeyDisplay        = 'Key 1'
                DisplayPriority   = 6
                Frequent          = $false
                Title1            = 'IJK'
                Title2            = 9
                Title3            = 4
                UserCode          = '98765'
                FolderScanType    = 'smb'
                FolderScanPath    = '\\folder\path'
                FolderScanPort    = 21
                FolderScanAccount = 'ScanAccount'
                IsDestination     = $true
            }
            [PSCustomObject]@{
                PSTypeName      = 'Ricoh.AddressBook.Entry'
                Name            = 'Name 2'
                KeyDisplay      = 'Key 2'
                DisplayPriority = 5
                Frequent        = $true
                EmailAddress    = 'mail@example.com'
                IsSender        = $false
                IsDestination   = $true
            }
        )
    }

    It 'Does not send request if there are no additions' {
        $parameters = @{
            Name           = 'John Doe'
            KeyDisplay     = 'John D'
            Frequent       = $false
            FolderScanPath = '\\folder\path'
        }

        # This will error because there are no tags present (-Frequent is false
        # and no -Title1, -Title2, or -Title3 are provided). The command itself
        # will still complete because that is a non-terminating error, so the
        # end {} block still runs.

        # However, if the <propListList /> is empty (nothing was added, most
        # likely because they errored), the request will fail with a server
        # error stating COMMON_BAD_PARAMETER. When there are no new entries to
        # add, then this function should skip the final call where it commits
        # the new objects, which is what this test is trying to ensure.
        Add-AddressBookEntry @commonParameters @parameters -ErrorAction SilentlyContinue

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 0 -ParameterFilter {
            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjects'
        }
    }
}

Describe 'Remove-AddressBookEntry' {
    It 'Has the correct parameters' {
        $command = Get-Command Remove-AddressBookEntry

        $command | Should -HaveParameter Hostname -Mandatory -Type [uri]
        $command | Should -HaveParameter Credential -Mandatory -Type [pscredential]
        $command | Should -HaveParameter Id -Mandatory -Type [uint32[]]
        $command | Should -HaveParameter SkipCertificateCheck -Not -Mandatory
        $command | Should -HaveParameter SkipCertificateCheck -Type [switch]
    }

    It 'Removes entries with the given IDs' {
        1..5 | Remove-AddressBookEntry @commonParameters

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 1 -ParameterFilter {
            $expected = [xml]'<?xml version="1.0" encoding="utf-8"?>
            <s:Envelope
                    xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
                    s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
                <s:Body>
                    <u:deleteObjects
                            xmlns:xs="http://www.w3.org/2001/XMLSchema"
                            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                            xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/"
                            xmlns:ricoh="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes"
                            xmlns:u="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
                        <sessionId>12345</sessionId>
                        <objectIdList xsi:type="soap-enc:Array" soap-enc:arrayType="xs:string[]">
                            <item>entry:1</item>
                            <item>entry:2</item>
                            <item>entry:3</item>
                            <item>entry:4</item>
                            <item>entry:5</item>
                        </objectIdList>
                    </u:deleteObjects>
                </s:Body>
            </s:Envelope>'

            $Headers.SOAPAction -eq 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#deleteObjects' -and
            $Body.OuterXml -eq $expected.OuterXml
        }
    }
}

Describe 'Get-Title1Tag' {
    It -ForEach @(
        @{ Letter = 'A'; Value = 2 }
        @{ Letter = 'B'; Value = 2 }
        @{ Letter = 'C'; Value = 3 }
        @{ Letter = 'D'; Value = 3 }
        @{ Letter = 'E'; Value = 4 }
        @{ Letter = 'F'; Value = 4 }
        @{ Letter = 'G'; Value = 5 }
        @{ Letter = 'H'; Value = 5 }
        @{ Letter = 'I'; Value = 6 }
        @{ Letter = 'J'; Value = 6 }
        @{ Letter = 'K'; Value = 6 }
        @{ Letter = 'L'; Value = 7 }
        @{ Letter = 'M'; Value = 7 }
        @{ Letter = 'N'; Value = 7 }
        @{ Letter = 'O'; Value = 8 }
        @{ Letter = 'P'; Value = 8 }
        @{ Letter = 'Q'; Value = 8 }
        @{ Letter = 'R'; Value = 9 }
        @{ Letter = 'S'; Value = 9 }
        @{ Letter = 'T'; Value = 9 }
        @{ Letter = 'U'; Value = 10 }
        @{ Letter = 'V'; Value = 10 }
        @{ Letter = 'W'; Value = 10 }
        @{ Letter = 'X'; Value = 11 }
        @{ Letter = 'Y'; Value = 11 }
        @{ Letter = 'Z'; Value = 11 }
        @{ Letter = 'a'; Value = 2 }
        @{ Letter = 'b'; Value = 2 }
        @{ Letter = 'c'; Value = 3 }
        @{ Letter = 'd'; Value = 3 }
        @{ Letter = 'e'; Value = 4 }
        @{ Letter = 'f'; Value = 4 }
        @{ Letter = 'g'; Value = 5 }
        @{ Letter = 'h'; Value = 5 }
        @{ Letter = 'i'; Value = 6 }
        @{ Letter = 'j'; Value = 6 }
        @{ Letter = 'k'; Value = 6 }
        @{ Letter = 'l'; Value = 7 }
        @{ Letter = 'm'; Value = 7 }
        @{ Letter = 'n'; Value = 7 }
        @{ Letter = 'o'; Value = 8 }
        @{ Letter = 'p'; Value = 8 }
        @{ Letter = 'q'; Value = 8 }
        @{ Letter = 'r'; Value = 9 }
        @{ Letter = 's'; Value = 9 }
        @{ Letter = 't'; Value = 9 }
        @{ Letter = 'u'; Value = 10 }
        @{ Letter = 'v'; Value = 10 }
        @{ Letter = 'w'; Value = 10 }
        @{ Letter = 'x'; Value = 11 }
        @{ Letter = 'y'; Value = 11 }
        @{ Letter = 'z'; Value = 11 }
    ) 'Returns <Value> for "<Letter>"' {
        $result = Get-Title1Tag $Letter

        $result | Should -Be $Value
    }
}

Describe 'CmdletBinding' {
    It -ForEach @(
        'Get-Title1Tag'
        'Get-AddressBookEntry'
        'Update-AddressBookEntry'
        'Add-AddressBookEntry'
        'Remove-AddressBookEntry'
    ) '<_> has a [CmdletBinding()] attribute set' {
        $command = Get-Command $_
        [CmdletBinding] $attribute = $command.ScriptBlock.Attributes |
            Where-Object { $_.TypeId -eq [CmdletBinding] }

        $attribute | Should -Not -BeNullOrEmpty -Because "$_ should have [CmdletBinding()]"
    }
}

Describe 'SupportsShouldProcess' {
    It -ForEach @('Update', 'Add', 'Remove') '<_>-AddressBookEntry uses SupportsShouldProcess' {
        $commandName = "$_-AddressBookEntry"
        $command = Get-Command $commandName
        [CmdletBinding] $attribute = $command.ScriptBlock.Attributes |
            Where-Object { $_.TypeId -eq [CmdletBinding] }

        $attribute.SupportsShouldProcess | Should -BeTrue -Because "$commandName should have [CmdletBinding(SupportsShouldProcess)]"
    }

    It -ForEach @(
        @{ Function = 'Update'; Arguments = @{ Id = 1 } }
        @{ Function = 'Add'; Arguments = @{ Name = 'A'; KeyDisplay = 'B'; EmailAddress = 'email@example.com'; Frequent = $true } }
        @{ Function = 'Remove'; Arguments = @{ Id = @(1) } }
    ) '<Function>-AddressBookEntry does not call Invoke-WebRequest if -WhatIf is provided' {
        & "$Function-AddressBookEntry" @commonParameters @Arguments -WhatIf

        Should -Invoke Invoke-WebRequest -ModuleName RicohAddressBook -Exactly -Times 0 -ParameterFilter {
            $Headers.SOAPAction -ne 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#startSession' -and
            $Headers.SOAPAction -ne 'http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#terminateSession'
        }
    }
}

Describe 'PositionalBinding' {
    It -ForEach @('Get', 'Update', 'Add', 'Remove') '<_>-AddressBookEntry uses PositionalBinding = $false' {
        $commandName = "$_-AddressBookEntry"
        $command = Get-Command $commandName
        [CmdletBinding] $attribute = $command.ScriptBlock.Attributes |
            Where-Object { $_.TypeId -eq [CmdletBinding] }

        $attribute.PositionalBinding | Should -BeFalse -Because "$commandName should have [CmdletBinding(PositionalBinding = `$false)]"
    }
}

Describe 'Parameter validation' {
    Context -ForEach @('Update', 'Add') '<_>-AddressBookEntry' {
        BeforeAll {
            $functionName = $_
            $otherParameters = @{}

            if ('Update' -eq $functionName) {
                $otherParameters.Id = 1
            } else {
                $otherParameters.Name = '.'
                $otherParameters.KeyDisplay = '.'
                $otherParameters.Frequent = $true
                $otherParameters.FolderScanPath = '.'
                $otherParameters.EmailAddress = '.'
            }
        }

        It -ForEach @(
            @{ ParameterName = 'Name'; MaximumLength = 20 }
            @{ ParameterName = 'KeyDisplay'; MaximumLength = 16 }
            @{ ParameterName = 'FolderScanPath'; MaximumLength = 256 }
            @{ ParameterName = 'EmailAddress'; MaximumLength = 128 }
            @{ ParameterName = 'UserCode'; MaximumLength = 8; Character = '0' }
        ) 'Throws a validation expection when <ParameterName> is longer than <MaximumLength> characters' {
            if ($null -eq $Character) {
                $Character = '.'
            }

            $parameters = $otherParameters.Clone()
            $parameters.Remove($ParameterName)

            {
                $validParameters = @{
                    $ParameterName = [string]::new($Character, $MaximumLength)
                }

                & "$functionName-AddressBookEntry" @commonParameters @parameters @validParameters
            } | Should -Not -Throw

            {
                $invalidParameters = @{
                    $ParameterName = [string]::new($Character, $MaximumLength + 1)
                }

                & "$functionName-AddressBookEntry" @commonParameters @parameters @invalidParameters -ErrorAction SilentlyContinue
            } | Should -Throw "Cannot validate argument on parameter '$ParameterName'. The character length of the $($MaximumLength + 1) argument is too long. Shorten the character length of the argument so it is fewer than or equal to `"$MaximumLength`" characters, and then try the command again."
        }

        It -ForEach @(
            @{ ParameterName = 'DisplayPriority'; MaximumValue = 10 }
            @{ ParameterName = 'Title2'; MaximumValue = 10 }
            @{ ParameterName = 'Title3'; MaximumValue = 5 }
        ) 'Throws a validation exception when <ParameterName> is greater than <MaximumValue>' {
            {
                $validParameters = @{
                    $ParameterName = $MaximumValue
                }

                & "$functionName-AddressBookEntry" @commonParameters @otherParameters @validParameters
            } | Should -Not -Throw

            {
                $invalidParameters = @{
                    $ParameterName = $MaximumValue + 1
                }

                & "$functionName-AddressBookEntry" @commonParameters @otherParameters @invalidParameters -ErrorAction SilentlyContinue
            } | Should -Throw "Cannot validate argument on parameter '$ParameterName'. The $($MaximumValue + 1) argument is greater than the maximum allowed range of $MaximumValue. Supply an argument that is less than or equal to $MaximumValue and then try the command again."
        }

        It -ForEach @('DisplayPriority', 'Title2', 'Title3') 'Throws a validation exception when <_> is less than 1' {
            {
                $validParameters = @{
                    $_ = 1
                }

                & "$functionName-AddressBookEntry" @commonParameters @otherParameters @validParameters
            } | Should -Not -Throw

            {
                $invalidParameters = @{
                    $_ = 0
                }

                & "$functionName-AddressBookEntry" @commonParameters @otherParameters @invalidParameters -ErrorAction SilentlyContinue
            } | Should -Throw "Cannot validate argument on parameter '$_'. The 0 argument is less than the minimum allowed range of 1. Supply an argument that is greater than or equal to 1 and then try the command again."
        }

        It 'Does not throw a validation exception when UserCode is number-like' {
            {
                & "$functionName-AddressBookEntry" @commonParameters @otherParameters -UserCode '01234567'
            } | Should -Not -Throw
        }

        It -ForEach @('xyz', '0123ABCD', 'ABCD0123') 'Throws a validation exception when UserCode is <_>, not number-like' {
            {
                & "$functionName-AddressBookEntry" @commonParameters @otherParameters -UserCode $_ -ErrorAction SilentlyContinue
            } | Should -Throw "Cannot validate argument on parameter 'UserCode'. The argument `"$_`" does not match the `"^\d+$`" pattern. Supply an argument that matches `"^\d+$`" and try the command again."
        }
    }
}
