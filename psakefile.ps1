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

Task TagRelease {
    $module = Test-ModuleManifest -Path (Join-Path 'Module' 'RicohAddressBook.psd1')
    $tagName = "v$($module.Version)"

    if ($PreRelease) {
        $lastRelease = git tag --list "$tagName-pre.*" |
            Where-Object { $_ -match '\.(\d+)$' } |
            ForEach-Object { [uint32] $Matches[1] } |
            Measure-Object -Maximum |
            Select-Object -ExpandProperty Maximum

        $next = [string] ($lastRelease + 1)
        $tagName += "-pre.$next"
    }

    git tag -a $tagName -m $tagName
}
