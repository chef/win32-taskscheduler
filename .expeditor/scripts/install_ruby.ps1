param(
    [Parameter(Mandatory = $true)]
    [string]$RubyVersion
)

$ErrorActionPreference = "Stop"

# RubyInstaller release tags/assets use a hyphen before the build revision
# (e.g. "3.1.6-1"), while this script's caller passes the Chocolatey-style
# dotted version (e.g. "3.1.6.1"). Convert by replacing the final dot.
$lastDot = $RubyVersion.LastIndexOf(".")
$InstallerVersion = $RubyVersion.Remove($lastDot, 1).Insert($lastDot, "-")

# Use the Devkit variant so native extensions (e.g. ffi, used by this gem)
# can be compiled via the bundled MSYS2/MinGW toolchain.
$InstallerFile = "rubyinstaller-devkit-$InstallerVersion-x64.exe"
$DownloadUrl = "https://github.com/oneclick/rubyinstaller2/releases/download/RubyInstaller-$InstallerVersion/$InstallerFile"
$InstallerPath = Join-Path $env:TEMP $InstallerFile

$versionParts = $InstallerVersion.Split("-")[0].Split(".")
$InstallDir = "C:\Ruby$($versionParts[0])$($versionParts[1])-x64"

Write-Output "--- Downloading Ruby $InstallerVersion from $DownloadUrl"
Invoke-WebRequest -Uri $DownloadUrl -OutFile $InstallerPath -UseBasicParsing

Write-Output "--- Installing Ruby $InstallerVersion to $InstallDir"
$process = Start-Process -FilePath $InstallerPath `
    -ArgumentList "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/TASKS=`"modpath`"", "/DIR=`"$InstallDir`"" `
    -Wait -PassThru
if ($process.ExitCode -ne 0) { throw "Failed to install Ruby $InstallerVersion. Installer exited with code $($process.ExitCode)." }

Write-Output "--- Updating PATH for this session"
$env:PATH = "$InstallDir\bin;$env:PATH"

Write-Output "--- Installing MSYS2/MinGW toolchain via ridk install"
"1,3" | & "$InstallDir\bin\ridk.cmd" install
if (-not $?) { throw "ridk install failed." }
