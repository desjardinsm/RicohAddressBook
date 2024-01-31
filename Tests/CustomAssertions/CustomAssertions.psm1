foreach ($script in
  (Get-ChildItem -File -Recurse -LiteralPath $PSScriptRoot -Filter *.ps1)
) {
    . $script.FullName
}
