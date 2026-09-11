[CmdletBinding()]
$ErrorActionPreference = 'Stop'

$destination = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\TN-Montants-V3.4.dotm'

if (Test-Path -LiteralPath $destination) {
    Remove-Item -LiteralPath $destination -Force
    Write-Host "TN-Montants supprimé de : $destination"
} else {
    Write-Host 'TN-Montants n''est pas installé dans le dossier Startup de l''utilisateur courant.'
}

Write-Host 'Redémarrez Word si nécessaire.'
