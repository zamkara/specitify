$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$spicetifyFolderPath = "$env:LOCALAPPDATA\spicetify"
$spicetifyOldFolderPath = "$HOME\spicetify-cli"
$spotifyRoaming = "$env:APPDATA\Spotify"
$spotifyLocal = "$env:LOCALAPPDATA\Spotify"
$spotifyExe = "$spotifyRoaming\Spotify.exe"
$spotifyInstaller = "$env:TEMP\SpotifySetup.exe"

function Write-Success { Write-Host ' > OK' -ForegroundColor Green }
function Write-Unsuccess { Write-Host ' > ERROR' -ForegroundColor Red }

function Test-PowerShellVersion {
    Write-Host 'Checking PowerShell version...' -NoNewline
    $PSVersionTable.PSVersion -ge [version]'5.1'
}

function Test-Admin {
    Write-Host 'Checking if not running as administrator...' -NoNewline
    $u = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    -not $u.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-SpotifyStoreBuild {
    (Get-AppxPackage -Name Spotify.Spotify -ErrorAction SilentlyContinue) -ne $null
}

function Remove-SpotifyStore {
    Write-Host 'Removing Microsoft Store Spotify...' -NoNewline
    Get-AppxPackage Spotify.Spotify -AllUsers | Remove-AppxPackage -ErrorAction SilentlyContinue
    Write-Success
}

function Remove-SpotifyFolders {
    Write-Host 'Cleaning Spotify folders...' -NoNewline
    Remove-Item $spotifyLocal -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $spotifyRoaming -Recurse -Force -ErrorAction SilentlyContinue
    Write-Success
}

function Install-SpotifyRelease {
    Write-Host 'Installing Spotify Release build...' -NoNewline
    Invoke-WebRequest 'https://download.scdn.co/SpotifySetup.exe' -OutFile $spotifyInstaller -UseBasicParsing
    Start-Process $spotifyInstaller -Wait
    Write-Success
}

function Wait-SpotifyReady {
    Write-Host 'Waiting for Spotify login...' -NoNewline
    while (-not (Test-Path $spotifyExe)) { Start-Sleep 1 }
    Write-Success
}

function Move-OldSpicetifyFolder {
    if (Test-Path $spicetifyOldFolderPath) {
        Write-Host 'Migrating old spicetify folder...' -NoNewline
        Copy-Item "$spicetifyOldFolderPath\*" $spicetifyFolderPath -Recurse -Force
        Remove-Item $spicetifyOldFolderPath -Recurse -Force
        Write-Success
    }
}

function Get-Spicetify {
    if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { $arch = 'arm64' }
    elseif ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') { $arch = 'x64' }
    else { $arch = 'x32' }

    Write-Host 'Fetching latest spicetify...' -NoNewline
    $r = Invoke-RestMethod 'https://api.github.com/repos/spicetify/cli/releases/latest'
    $v = $r.tag_name -replace 'v',''
    Write-Success

    $zip = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'spicetify.zip')

    Write-Host "Downloading spicetify v$v..." -NoNewline
    Invoke-WebRequest "https://github.com/spicetify/cli/releases/download/v$v/spicetify-$v-windows-$arch.zip" -OutFile $zip -UseBasicParsing
    Write-Success

    $zip
}

function Add-SpicetifyToPath {
    Write-Host 'Adding spicetify to PATH...' -NoNewline
    $u = [EnvironmentVariableTarget]::User
    $p = [Environment]::GetEnvironmentVariable('PATH', $u)
    $p = $p -replace "$([regex]::Escape($spicetifyOldFolderPath))\\*;*", ''
    if ($p -notlike "*$spicetifyFolderPath*") { $p = "$p;$spicetifyFolderPath" }
    [Environment]::SetEnvironmentVariable('PATH', $p, $u)
    $env:PATH = $p
    Write-Success
}

function Install-Spicetify {
    Write-Host 'Installing spicetify...'
    $zip = Get-Spicetify
    Write-Host 'Extracting spicetify...' -NoNewline
    Expand-Archive $zip $spicetifyFolderPath -Force
    Write-Success
    Add-SpicetifyToPath
    Remove-Item $zip -Force -ErrorAction SilentlyContinue
    Write-Host 'spicetify installed successfully!' -ForegroundColor Green
}

if (-not (Test-PowerShellVersion)) { Write-Unsuccess } else { Write-Success }
if (-not (Test-Admin)) { Write-Unsuccess } else { Write-Success }

if (Test-SpotifyStoreBuild) {
    Remove-SpotifyStore
    Remove-SpotifyFolders
}

if (-not (Test-Path $spotifyExe)) {
    Install-SpotifyRelease
    Wait-SpotifyReady
}

Move-OldSpicetifyFolder
Install-Spicetify

Write-Host "`nRun after Spotify login:" -ForegroundColor Cyan
Write-Host 'spicetify backup apply' -ForegroundColor Cyan
Write-Host 'spicetify apply' -ForegroundColor Cyan

$choices = @(
    (New-Object Management.Automation.Host.ChoiceDescription '&Yes'),
    (New-Object Management.Automation.Host.ChoiceDescription '&No')
)

$choice = $Host.UI.PromptForChoice('', 'Install Spicetify Marketplace?', $choices, 0)

if ($choice -eq 0) {
    Invoke-WebRequest 'https://raw.githubusercontent.com/spicetify/spicetify-marketplace/main/resources/install.ps1' -UseBasicParsing | Invoke-Expression
}
