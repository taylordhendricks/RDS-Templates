<#
.SYNOPSIS
    Install and setup the latest version of 7-Zip.

.DESCRIPTION
    This script downloads and installs the latest 7-Zip from GitHub releases.

.EXAMPLE
    .\Install-7Zip.ps1 -InstallPath "C:\Program Files\7-Zip" -LogRetentionDays 60 -Verbose

.NOTES
    Author: Taylor Hendricks
    Date: 10/01/2024
    Version: 1.3
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "C:\Program Files\7-Zip",
    [int]$LogRetentionDays = 30,
    [switch]$NoCleanup
)

# Variables
$githubReleasesApiUrl = "https://api.github.com/repos/ip7z/7zip/releases/latest"
$localTemp            = "C:\temp\7Zip"
$timestamp            = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFilePath          = "C:\Windows\Logs"
$logFileName          = "7Zip_Install_$timestamp.log"
$logFileFullPath      = Join-Path -Path $logFilePath -ChildPath $logFileName
$architecture         = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Create-Directories {
    param([string[]]$Paths)
    foreach ($path in $Paths) {
        if (-not (Test-Path -Path $path)) {
            Write-Verbose "Creating directory: $path"
            New-Item -Path $path -ItemType Directory | Out-Null
        } else {
            Write-Verbose "Directory already exists: $path"
        }
    }
}

function Get-Latest7ZipRelease {
    try {
        Write-Verbose "Fetching latest 7-Zip release info from GitHub."
        return Invoke-RestMethod -Uri $githubReleasesApiUrl -Headers @{"User-Agent"="PowerShell"}
    } catch {
        Write-Error "Failed to fetch latest release info: $_"
        throw
    }
}

function Download-Installer {
    param($Url, $Destination)
    try {
        Write-Verbose "Downloading installer from $Url"
        Invoke-WebRequest -Uri $Url -OutFile $Destination -ErrorAction Stop
    } catch {
        Write-Error "Failed to download installer: $_"
        throw
    }
}

function Install-Software {
    param($InstallerPath, $LogFile)
    $msiArguments = @(
        "/i", "`"$InstallerPath`"",
        "/qn",
        "/norestart",
        "/log", "`"$LogFile`""
    )
    try {
        Write-Verbose "Starting installation with arguments: $msiArguments"
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArguments -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -ne 0) {
            Write-Error "Installation failed with exit code $($process.ExitCode)"
            throw
        }
    } catch {
        Write-Error "Installation failed: $_"
        throw
    }
}

function Validate-Installation {
    param($ExpectedVersion)
    $registryPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    $softwareKey = Get-ItemProperty -Path $registryPath | Where-Object {
        $_.DisplayName -like '7-Zip*'
    }
    if ($softwareKey) {
        $installedVersion = $softwareKey.DisplayVersion

        # Split the version strings into arrays
        $expectedVersionParts = $ExpectedVersion -split '\.'
        $installedVersionParts = $installedVersion -split '\.'

        # Ensure we have at least two parts
        if ($expectedVersionParts.Length -ge 2 -and $installedVersionParts.Length -ge 2) {
            # Take the first two components
            $expectedVersionShort = ($expectedVersionParts[0..1] -join '.')
            $installedVersionShort = ($installedVersionParts[0..1] -join '.')

            # Compare the shortened versions
            if ($installedVersionShort -eq $expectedVersionShort) {
                Write-Verbose "7-Zip version $ExpectedVersion installed successfully."
            } else {
                Write-Error "Installed version ($installedVersion) does not match expected version ($ExpectedVersion)."
                throw
            }
        } else {
            Write-Error "Version strings are not in the expected format."
            throw
        }
    } else {
        Write-Error "7-Zip is not installed."
        throw
    }
}

function Cleanup-Files {
    param($Paths)
    if ($NoCleanup) {
        Write-Verbose "No cleanup performed due to NoCleanup switch."
        return
    }
    foreach ($path in $Paths) {
        try {
            Write-Verbose "Cleaning up: $path"
            Remove-Item -Path $path -Force -Recurse
        } catch {
            Write-Warning "Could not clean up $path - $_"
        }
    }
    # Delete old log files
    try {
        Get-ChildItem -Path $logFilePath -Filter '7Zip_Install_*.log' | Where-Object {
            $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays)
        } | Remove-Item -Force
    } catch {
        Write-Warning "Could not clean up old log files: $_"
    }
}

# Main Script Execution
try {
    Create-Directories -Paths @($localTemp, $logFilePath)

    $latestRelease = Get-Latest7ZipRelease
    $downloadUrl = $latestRelease.assets | Where-Object { $_.name -like "*-$architecture.msi" } | Select-Object -ExpandProperty browser_download_url

    if (-not $downloadUrl) {
        Write-Error "No MSI installer found for architecture $architecture."
        exit 1
    }

    $sevenZip_Version = $latestRelease.tag_name.TrimStart('v')
    $installerName = "7Zip_$sevenZip_Version-$architecture.msi"
    $installerPath = Join-Path -Path $localTemp -ChildPath $installerName

    Download-Installer -Url $downloadUrl -Destination $installerPath
    Install-Software -InstallerPath $installerPath -LogFile $logFileFullPath
    Validate-Installation -ExpectedVersion $sevenZip_Version
    Cleanup-Files -Paths @($localTemp)

    Write-Host "7-Zip installation script completed successfully."
} catch {
    Write-Error "Script failed: $_"
    exit 1
}
