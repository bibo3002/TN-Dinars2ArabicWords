[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DotmPath
)

$ErrorActionPreference = 'Stop'

$source = (Resolve-Path -LiteralPath $DotmPath).Path
$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'

New-Item -ItemType Directory -Path $startup -Force | Out-Null
$destination = Join-Path $startup 'TN-Montants-V3.4.dotm'

Copy-Item -LiteralPath $source -Destination $destination -Force

Write-Host "TN-Montants installé dans : $destination"
Write-Host 'Redémarrez Word pour charger le modèle.'
