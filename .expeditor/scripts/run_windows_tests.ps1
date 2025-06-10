# Stop script execution when a non-terminating error occurs
$ErrorActionPreference = "Stop"

# Find the latest Ruby install directory
$RubyDir = Get-ChildItem -Directory "C:\Ruby*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($RubyDir) {
    $RubyBin = "$($RubyDir.FullName)\bin"
    Write-Output "Ruby installed at $RubyBin"
} else {
    throw "Could not find Ruby installation directory."
}
$env:PATH = "$env:PATH;$RubyBin"

Write-Output "--- Installed Ruby version"
ruby --version

Write-Output "--- Bundle install"

bundle config --local path vendor/bundle
bundle install --jobs=7 --retry=3
if (-not $?) { throw "bundle install failed" }

Write-Output "--- Running Cookstyle"
gem install cookstyle
cookstyle --chefstyle -c .rubocop.yml
if (-not $?) { throw "Cookstyle failed." }

Write-Output "--- Bundle Execute"

bundle exec rake
if (-not $?) { throw "Rake task failed." }