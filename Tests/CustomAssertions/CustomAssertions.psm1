$files = @{
    File        = $true
    Recurse     = $true
    LiteralPath = $PSScriptRoot
    Filter      = '*.ps1'
}

foreach ($script in Get-ChildItem @files) {
    . $script.FullName
}
