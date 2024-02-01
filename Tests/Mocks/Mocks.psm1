$mocks = Import-PowerShellDataFile (
    Join-Path $PSScriptRoot 'ResponseMocks.psd1'
)

function Get-SoapResponseMock {
    param(
        [string]
        [ValidateSet('startSession', 'searchObjects', 'getObjectsProps')]
        $SoapAction
    )

    $mock = $mocks.'Invoke-WebRequest'.$SoapAction

    { [xml] $mock }.GetNewClosure()
}
