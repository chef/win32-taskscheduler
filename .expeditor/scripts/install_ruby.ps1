param(
    [Parameter(Mandatory = $true)]
    [string]$RubyVersion
)

$ErrorActionPreference = "Stop"

Write-Output "--- Installing Ruby $RubyVersion using Chocolatey"
choco install ruby --version $RubyVersion -y --force
if (-not $?) { throw "Failed to install Ruby $RubyVersion." }

Write-Output "Updating PATH using refreshenv"
Import-Module $env:ChocolateyInstall\helpers\chocolateyProfile.psm1
refreshenv
