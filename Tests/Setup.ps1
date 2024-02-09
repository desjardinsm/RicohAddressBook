& {
    function Get-RelativePath {
        param(
            [string]
            [Parameter(Mandatory)]
            $Path
        )

        $result = Join-Path $PSScriptRoot $Path

        $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
            $result
        )
    }

    Import-Module (Get-RelativePath './CustomAssertions')
    Import-Module (Get-RelativePath './Mocks')

    Remove-Module [R]icohAddressBook
    Import-Module (Get-RelativePath '../Module/RicohAddressBook.psd1')
}

function Cleanup {
    $modules = Get-Module -All CustomAssertions, Mocks, RicohAddressBook
    Remove-Module -ModuleInfo $modules
}
