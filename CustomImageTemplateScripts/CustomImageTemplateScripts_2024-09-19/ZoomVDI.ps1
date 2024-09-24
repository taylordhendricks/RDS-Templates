<# 
.SYNOPSIS
    Install and setup Zoom VDI.

.DESCRIPTION
    This script downloads and installs Zoom VDI with specified configurations.

.PARAMETER zoomVDI_Version
    The version of Zoom VDI to install.
    Reference: [Zoom VDI Releases](https://support.zoom.com/hc/en/article?id=zm_kb&sysparm_article=KB0063810)

.PARAMETER zSSOHost
    The SSO host for Zoom configuration.
    Reference: [Zoom VDI CLI Configuration](https://support.zoom.com/hc/en/article?id=zm_kb&sysparm_article=KB0064484#h_01FW2JNVWMVKJX9A10EV3TRTMB)

.EXAMPLE
    .\Install-ZoomVDI.ps1 -zoomVDI_Version "6.1.10.25260" -zSSOHost "hyland.com"

.NOTES
    Author: Taylor Hendricks
    Date: 09/19/2024
    Version: 1.1

#>

param(
    [string]$zoomVDI_Version = "6.1.10.25260",
    [string]$zSSOHost = "hyland.com"
)

# Variables
$zoomVDI_URL        = "https://zoom.us/download/vdi/$zoomVDI_Version/ZoomInstallerVDI.msi?archType=x64"
$localTemp          = "C:\temp"
$localPath          = Join-Path -Path $localTemp -ChildPath 'zoomVDI'
$zoomVDIInstaller   = 'ZoomInstallerVDI.msi'
$timestamp          = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFilePath        = "C:\Windows\Logs\"
$logFileName        = "zoomVDI_MSI_$timestamp.log"
$logFileFullPath    = Join-Path -Path $logFilePath -ChildPath $logFileName

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Create Temp Directories if they don't exist
foreach ($path in @($localTemp, $localPath)) {
    if (-not (Test-Path -Path $path)) {
        Write-Host "AVD AIB Customization:- Install ZoomVDI: Creating directory: $path"
        New-Item -Path $path -ItemType Directory | Out-Null
    } else {
        Write-Host "AVD AIB Customization:- Install ZoomVDI: Directory already exists: $path"
    }
}

# Download Zoom VDI Installer
$installerPath = Join-Path -Path $localPath -ChildPath $zoomVDIInstaller
try {
    Write-Host "AVD AIB Customization:- Install ZoomVDI: Downloading Zoom VDI installer from $zoomVDI_URL"
    #Invoke-WebRequest -Uri $zoomVDI_URL -OutFile $installerPath -ErrorAction Stop
    $wc = new-object System.Net.WebClient
    $wc.DownloadFile("$zoomVDI_URL","$installerPath")
} catch {
    Write-Host "AVD AIB Customization:- Install ZoomVDI: ERROR, Failed to download Zoom VDI installer: $_"
    exit 1
}

# Install Zoom VDI
Write-Host "AVD AIB Customization:- Install ZoomVDI: Starting Zoom VDI installation."
$msiArguments = @(
    "/i `"$installerPath`""
    "/qn"
    "/norestart"
    "/log `"$logFileFullPath`""
    "ZNODESKTOPSHORTCUT=True"
    "ZSILENTSTART=1"
    "ZSSOHOST=$zSSOHost"
    "ZCONFIG=noGoogle=1;noFacebook=1;enableAppleLogin=0;AutoSSOLogin=1"
)

try {
    Write-Host "Installing Zoom with the following Arguments $msiArguments"
    $zoomVDI_deploy_status = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArguments -Wait -PassThru -ErrorAction Stop
} catch {
    Write-Host "AVD AIB Customization:- Install ZoomVDI: ERROR Zoom VDI installation failed: $_"
    exit 1
}

# Validate Installation
$registryPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
$zoomKey = Get-ItemProperty -Path $registryPath | Where-Object {
    $_.DisplayName -like 'Zoom*VDI*'
}

if ($zoomKey) {
    $installedVersion = $zoomKey.DisplayVersion
    if ($installedVersion -eq $zoomVDI_Version) {
        Write-Host "AVD AIB Customization:- Install ZoomVDI: Zoom VDI version $zoomVDI_Version installed successfully."
    } else {
        Write-Host "AVD AIB Customization:- Install ZoomVDI: ERROR Installed Zoom VDI version ($installedVersion) does not match expected version ($zoomVDI_Version)."
        exit 1
    }
} else {
    Write-Host "AVD AIB Customization:- Install ZoomVDI: ERROR Zoom VDI is not installed."
    exit 1
}

# Cleanup
try {
    Write-Host "AVD AIB Customization:- Install ZoomVDI: Cleaning up log and temporary files."
    Remove-Item -Path $localPath -Force -Recurse
    # Define how many days of logs to keep
    $logRetentionDays = 30
    # Delete log files older than retention period
    Get-ChildItem -Path $logFilePath -Filter 'zoomVDI_MSI_*.log' | Where-Object {
        $_.LastWriteTime -lt (Get-Date).AddDays(-$logRetentionDays)
    } | Remove-Item -Force  
} catch {
    Write-Host "AVD AIB Customization:- Install ZoomVDI: WARNING Could not cleanup: $_"
}

Write-Host "AVD AIB Customization:- Install ZoomVDI: Zoom VDI installation script completed successfully."
