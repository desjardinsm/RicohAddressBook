#Requires -Module psake
Task Test {
    $configuration = New-PesterConfiguration

    $configuration.Run.Path = './Tests/'
    $configuration.Run.Throw = $true
    $configuration.Output.Verbosity = if ($Detailed) {
        'Detailed'
    } else {
        'Normal'
    }

    if ('True' -eq $env:CI -or $Results) {
        $resultsDirectory = 'TestResults'

        $configuration.TestResult.Enabled = $true
        $configuration.TestResult.OutputFormat = 'NUnit3'
        $configuration.TestResult.OutputPath = Join-Path $resultsDirectory 'testResults.xml'

        $configuration.CodeCoverage.Enabled = $true
        $configuration.CodeCoverage.Path = './Module/'
        $configuration.CodeCoverage.OutputPath = Join-Path $resultsDirectory 'coverage.xml'
    }

    Invoke-Pester -Configuration $configuration
}
