# Define variables
$logFile = "C:\Logs\UpdateInstallLog.txt"
$msuUrl = "https://www.catalog.update.microsoft.com/Search.aspx?q=KB5058411"  # Replace with the actual download URL for the latest 24H2 .msu file
$msuFile = "C:\Updates\KB5058411.msu"
$updateName = "KB5058411"

# Create log directory if it doesn't exist
if (-not (Test-Path "C:\Logs")) {
    New-Item -ItemType Directory -Path "C:\Logs"
}

# Create update directory if it doesn't exist
if (-not (Test-Path "C:\Updates")) {
    New-Item -ItemType Directory -Path "C:\Updates"
}

# Function to log messages
function Log-Message {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFile -Value $logMessage
    Write-Host $logMessage
}

# Download the .msu file
Log-Message "Starting download of the .msu file from $msuUrl"
Invoke-WebRequest -Uri $msuUrl -OutFile $msuFile
if ($?) {
    Log-Message "Download completed successfully."
} else {
    Log-Message "Download failed."
    exit 1
}

# Install the .msu file
Log-Message "Starting installation of the .msu file."
Start-Process -FilePath "wusa.exe" -ArgumentList "$msuFile /quiet /norestart" -Wait
if ($?) {
    Log-Message "Installation completed successfully."
} else {
    Log-Message "Installation failed."
    exit 1
}

# Verify the installation
Log-Message "Verifying the installation of the update."
$installedUpdates = Get-WmiObject -Query "Select * from Win32_QuickFixEngineering where HotFixID='$updateName'"
if ($installedUpdates) {
    Log-Message "Update $updateName installed successfully."
} else {
    Log-Message "Update $updateName installation verification failed."
    exit 1
}

Log-Message "Script completed."
