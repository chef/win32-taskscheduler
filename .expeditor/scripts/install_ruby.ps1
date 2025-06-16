param(
    [Parameter(Mandatory = $true)]
    [string]$RubyVersion
)

$ErrorActionPreference = "Stop"

Write-Output "--- Downloading and installing Ruby"
$installerUrl = "https://github.com/oneclick/rubyinstaller2/releases/download/RubyInstaller-$RubyVersion-1/rubyinstaller-devkit-$RubyVersion-1-x64.exe"
$installerPath = "$env:TEMP\rubyinstaller.exe"

Write-Output "Downloading Ruby $RubyVersion from $installerUrl"
Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
if (-not $?) { throw "Could not download ruby with devkit from https://github.com/oneclick/rubyinstaller2" }

Write-Output "Running Ruby installer..."
Start-Process -FilePath $installerPath -ArgumentList "/verysilent","/tasks=modpath", "/currentuser" -Wait
if (-not $?) { throw "Failed to install Ruby." }
