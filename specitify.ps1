$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$spicetifyFolderPath = "$env:LOCALAPPDATA\spicetify"
$spicetifyOldFolderPath = "$HOME\spicetify-cli"
$spotifyAppPath = "C:\Users\F15\AppData\Roaming\Spotify"
$spotifyExePathLocal = "C:\Users\F15\AppData\Local\Spotify\Spotify.exe"
$spotifyExePathRoaming = "C:\Users\F15\AppData\Roaming\Spotify\Spotify.exe"
$marketAppPath = "$spotifyAppPath\Spicetify Marketplace"
$marketThemePath = "$marketAppPath\themes"

function Test-SpotifyInstalled {
    [CmdletBinding()]
    param ()
    process {
        if (Test-Path -Path $spotifyAppPath) {
            Write-Host -Object 'Spotify is already installed.' -ForegroundColor 'Green'
            return $true
        } else {
            Write-Host -Object 'Spotify is not installed.' -ForegroundColor 'Red'
            return $false
        }
    }
}

function Test-SpotifyRunning {
    [CmdletBinding()]
    param ()
    process {
        $spotifyProcess = Get-Process -Name "Spotify" -ErrorAction SilentlyContinue
        return $spotifyProcess -ne $null
    }
}

function Install-Spotify {
    [CmdletBinding()]
    param ()
    process {
        Write-Host -Object 'Spotify is not installed. Installing Spotify...' -NoNewline
        try {
            Start-Process "winget" -ArgumentList "install -e --id Spotify.Spotify" -Wait
            Write-Success
        } catch {
            Write-Unsuccess
            Write-Host -Object 'Failed to install Spotify.' -ForegroundColor 'Red'
            exit
        }
    }
}

function Write-Success {
    [CmdletBinding()]
    param ()
    process {
        Write-Host -Object ' > OK' -ForegroundColor 'Green'
    }
}

function Write-Unsuccess {
    [CmdletBinding()]
    param ()
    process {
        Write-Host -Object ' > ERROR' -ForegroundColor 'Red'
    }
}

function Test-Admin {
    [CmdletBinding()]
    param ()
    begin {
        Write-Host -Object "Checking if the script is not being run as administrator..." -NoNewline
    }
    process {
        $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        -not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
}

function Test-PowerShellVersion {
    [CmdletBinding()]
    param ()
    begin {
        $PSMinVersion = [version]'5.1'
    }
    process {
        Write-Host -Object 'Checking if your PowerShell version is compatible...' -NoNewline
        $PSVersionTable.PSVersion -ge $PSMinVersion
    }
}

function Move-OldSpicetifyFolder {
    [CmdletBinding()]
    param ()
    process {
        if (Test-Path -Path $spicetifyOldFolderPath) {
            Write-Host -Object 'Moving the old spicetify folder...' -NoNewline
            Copy-Item -Path "$spicetifyOldFolderPath\*" -Destination $spicetifyFolderPath -Recurse -Force
            Remove-Item -Path $spicetifyOldFolderPath -Recurse -Force
            Write-Success
        }
    }
}

function Get-Spicetify {
    [CmdletBinding()]
    param ()
    begin {
        if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
            $architecture = 'x64'
        }
        elseif ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
            $architecture = 'arm64'
        }
        else {
            $architecture = 'x32'
        }
        if ($v) {
            if ($v -match '^\d+\.\d+\.\d+$') {
                $targetVersion = $v
            }
            else {
                Write-Warning -Message "You have spicefied an invalid spicetify version: $v `nThe version must be in the following format: 1.2.3"
                Pause
                exit
            }
        }
        else {
            Write-Host -Object 'Fetching the latest spicetify version...' -NoNewline
            $latestRelease = Invoke-RestMethod -Uri 'https://api.github.com/repos/spicetify/cli/releases/latest'
            $targetVersion = $latestRelease.tag_name -replace 'v', ''
            Write-Success
        }
        $archivePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "spicetify.zip")
    }
    process {
        Write-Host -Object "Downloading spicetify v$targetVersion..." -NoNewline
        $Parameters = @{
            Uri            = "https://github.com/spicetify/cli/releases/download/v$targetVersion/spicetify-$targetVersion-windows-$architecture.zip"
            UseBasicParsin = $true
            OutFile        = $archivePath
        }
        Invoke-WebRequest @Parameters
        Write-Success
    }
    end {
        $archivePath
    }
}

function Add-SpicetifyToPath {
    [CmdletBinding()]
    param ()
    begin {
        Write-Host -Object 'Making spicetify available in the PATH...' -NoNewline
        $user = [EnvironmentVariableTarget]::User
        $path = [Environment]::GetEnvironmentVariable('PATH', $user)
    }
    process {
        $path = $path -replace "$([regex]::Escape($spicetifyOldFolderPath))\\*;*", ''
        if ($path -notlike "*$spicetifyFolderPath*") {
            $path = "$path;$spicetifyFolderPath"
        }
    }
    end {
        [Environment]::SetEnvironmentVariable('PATH', $path, $user)
        $env:PATH = $path
        Write-Success
    }
}

function Install-Spicetify {
    [CmdletBinding()]
    param ()
    begin {
        Write-Host -Object 'Installing spicetify...'
    }
    process {
        $archivePath = Get-Spicetify
        Write-Host -Object 'Extracting spicetify...' -NoNewline
        Expand-Archive -Path $archivePath -DestinationPath $spicetifyFolderPath -Force
        Write-Success
        Add-SpicetifyToPath
    }
    end {
        Remove-Item -Path $archivePath -Force -ErrorAction 'SilentlyContinue'
        Write-Host -Object 'spicetify was successfully installed!' -ForegroundColor 'Green'
    }
}

if (-not (Test-SpotifyInstalled)) {
    Install-Spotify
}

if (-not (Test-SpotifyRunning)) {
    Write-Host 'Starting Spotify...'
    if (Test-Path -Path $spotifyExePathLocal) {
        Start-Process $spotifyExePathLocal -ArgumentList "--minimized"
    } elseif (Test-Path -Path $spotifyExePathRoaming) {
        Start-Process $spotifyExePathRoaming -ArgumentList "--minimized"
    }
    Start-Sleep -Seconds 5
}

if (-not (Test-PowerShellVersion)) {
    Write-Unsuccess
    Write-Warning -Message 'PowerShell 5.1 or higher is required to run this script'
    Write-Warning -Message "You are running PowerShell $($PSVersionTable.PSVersion)"
    Write-Host -Object 'PowerShell 5.1 install guide:'
    Write-Host -Object 'https://learn.microsoft.com/skypeforbusiness/set-up-your-computer-for-windows-powershell/download-and-install-windows-powershell-5-1'
    Write-Host -Object 'PowerShell 7 install guide:'
    Write-Host -Object 'https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows'
    Pause
    exit
} else {
    Write-Success
}

if (-not (Test-Admin)) {
    Write-Unsuccess
    Write-Warning -Message "The script was run as administrator. This can result in problems with the installation process or unexpected behavior. Do not continue if you do not know what you are doing."
    $Host.UI.RawUI.Flushinputbuffer()
    $choices = [System.Management.Automation.Host.ChoiceDescription[]] @(
        (New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Abort installation.'),
        (New-Object System.Management.Automation.Host.ChoiceDescription '&No', 'Resume installation.')
    )
    $choice = $Host.UI.PromptForChoice('', 'Do you want to abort the installation process?', $choices, 0)
    if ($choice -eq 0) {
        Write-Host -Object 'spicetify installation aborted' -ForegroundColor 'Yellow'
        Pause
        exit
    }
} else {
    Write-Success
}

Move-OldSpicetifyFolder
Install-Spicetify
Write-Host -Object "`nRun" -NoNewline
Write-Host -Object ' spicetify -h ' -NoNewline -ForegroundColor 'Cyan'
Write-Host -Object 'to get started'
$Host.UI.RawUI.Flushinputbuffer()
$choices = [System.Management.Automation.Host.ChoiceDescription[]] @(
    (New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Install Spicetify Marketplace."),
    (New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do not install Spicetify Marketplace.")
)
$choice = $Host.UI.PromptForChoice('', "`nDo you also want to install Spicetify Marketplace? It will become available within the Spotify client, where you can easily install themes and extensions.", $choices, 0)
if ($choice -eq 1) {
    Write-Host -Object 'spicetify Marketplace installation aborted' -ForegroundColor 'Yellow'
} else {
    Write-Host -Object 'Starting the spicetify Marketplace installation script..'
    $Parameters = @{
        Uri             = 'https://raw.githubusercontent.com/spicetify/spicetify-marketplace/main/resources/install.ps1'
        UseBasicParsing = $true
    }
    Invoke-WebRequest @Parameters | Invoke-Expression
}