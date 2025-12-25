$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$spicetifyFolderPath = "$env:LOCALAPPDATA\spicetify"
$spicetifyOldFolderPath = "$HOME\spicetify-cli"
$spotifyAppPath = "C:\Users\zam\AppData\Roaming\Spotify"
$spotifyExePathLocal = "C:\Users\zam\AppData\Local\Spotify\Spotify.exe"
$spotifyExePathRoaming = "C:\Users\zam\AppData\Roaming\Spotify\Spotify.exe"

function Write-Success {
    Write-Host ' > OK' -ForegroundColor Green
}

function Write-Unsuccess {
    Write-Host ' > ERROR' -ForegroundColor Red
}

function Test-SpotifyInstalled {
    if (Test-Path $spotifyAppPath) {
        Write-Host 'Spotify is already installed.' -ForegroundColor Green
        $true
    } else {
        Write-Host 'Spotify is not installed.' -ForegroundColor Red
        $false
    }
}

function Test-SpotifyRunning {
    (Get-Process Spotify -ErrorAction SilentlyContinue) -ne $null
}

function Test-SpotifyStoreBuild {
    (Get-AppxPackage -Name Spotify.Spotify -ErrorAction SilentlyContinue) -ne $null
}

function Install-Spotify {
    Write-Host 'Spotify is not installed. Installing Spotify...' -NoNewline
    try {
        Start-Process winget -ArgumentList 'install -e --id Spotify.Spotify --source winget' -Wait
        Write-Success
    } catch {
        Write-Unsuccess
        Write-Host 'Failed to install Spotify.' -ForegroundColor Red
    }
}

function Test-Admin {
    Write-Host 'Checking if the script is not being run as administrator...' -NoNewline
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    -not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-PowerShellVersion {
    $PSMinVersion = [version]'5.1'
    Write-Host 'Checking if your PowerShell version is compatible...' -NoNewline
    $PSVersionTable.PSVersion -ge $PSMinVersion
}

function Move-OldSpicetifyFolder {
    if (Test-Path $spicetifyOldFolderPath) {
        Write-Host 'Moving the old spicetify folder...' -NoNewline
        Copy-Item "$spicetifyOldFolderPath\*" $spicetifyFolderPath -Recurse -Force
        Remove-Item $spicetifyOldFolderPath -Recurse -Force
        Write-Success
    }
}

function Get-Spicetify {
    if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
        $architecture = 'x64'
    } elseif ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
        $architecture = 'arm64'
    } else {
        $architecture = 'x32'
    }

    Write-Host 'Fetching the latest spicetify version...' -NoNewline
    $latestRelease = Invoke-RestMethod 'https://api.github.com/repos/spicetify/cli/releases/latest'
    $targetVersion = $latestRelease.tag_name -replace 'v', ''
    Write-Success

    $archivePath = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'spicetify.zip')

    Write-Host "Downloading spicetify v$targetVersion..." -NoNewline
    Invoke-WebRequest -Uri "https://github.com/spicetify/cli/releases/download/v$targetVersion/spicetify-$targetVersion-windows-$architecture.zip" -OutFile $archivePath -UseBasicParsing
    Write-Success

    $archivePath
}

function Add-SpicetifyToPath {
    Write-Host 'Making spicetify available in the PATH...' -NoNewline
    $user = [EnvironmentVariableTarget]::User
    $path = [Environment]::GetEnvironmentVariable('PATH', $user)
    $path = $path -replace "$([regex]::Escape($spicetifyOldFolderPath))\\*;*", ''
    if ($path -notlike "*$spicetifyFolderPath*") {
        $path = "$path;$spicetifyFolderPath"
    }
    [Environment]::SetEnvironmentVariable('PATH', $path, $user)
    $env:PATH = $path
    Write-Success
}

function Install-Spicetify {
    Write-Host 'Installing spicetify...'
    $archivePath = Get-Spicetify
    Write-Host 'Extracting spicetify...' -NoNewline
    Expand-Archive $archivePath $spicetifyFolderPath -Force
    Write-Success
    Add-SpicetifyToPath
    Remove-Item $archivePath -Force -ErrorAction SilentlyContinue
    Write-Host 'spicetify was successfully installed!' -ForegroundColor Green
}

if (-not (Test-SpotifyInstalled)) {
    Install-Spotify
}

if (-not (Test-SpotifyRunning)) {
    Write-Host 'Starting Spotify...'
    if (Test-Path $spotifyExePathLocal) {
        Start-Process $spotifyExePathLocal '--minimized'
    } elseif (Test-Path $spotifyExePathRoaming) {
        Start-Process $spotifyExePathRoaming '--minimized'
    }
    Start-Sleep 5
}

if (-not (Test-PowerShellVersion)) {
    Write-Unsuccess
    Write-Warning 'PowerShell 5.1 or higher is required'
} else {
    Write-Success
}

if (-not (Test-Admin)) {
    Write-Unsuccess
    Write-Warning 'The script was run as administrator'
} else {
    Write-Success
}

if (Test-SpotifyStoreBuild) {
    Write-Warning 'Detected Microsoft Store / Master build of Spotify'
    Write-Warning 'Spicetify only works with Release build'
    Write-Warning 'Uninstall Spotify and reinstall from https://www.spotify.com/download/windows/'
} else {
    Move-OldSpicetifyFolder
    Install-Spicetify

    Write-Host "`nRun" -NoNewline
    Write-Host ' spicetify backup apply ' -NoNewline -ForegroundColor Cyan
    Write-Host 'then'
    Write-Host ' spicetify apply ' -ForegroundColor Cyan

    $Host.UI.RawUI.FlushInputBuffer()
    $choices = @(
        (New-Object Management.Automation.Host.ChoiceDescription '&Yes'),
        (New-Object Management.Automation.Host.ChoiceDescription '&No')
    )
    $choice = $Host.UI.PromptForChoice('', 'Do you also want to install Spicetify Marketplace?', $choices, 0)

    if ($choice -eq 0) {
        Write-Host 'Starting the spicetify Marketplace installation script..'
        Invoke-WebRequest 'https://raw.githubusercontent.com/spicetify/spicetify-marketplace/main/resources/install.ps1' -UseBasicParsing | Invoke-Expression
    } else {
        Write-Host 'spicetify Marketplace installation aborted' -ForegroundColor Yellow
    }
}
