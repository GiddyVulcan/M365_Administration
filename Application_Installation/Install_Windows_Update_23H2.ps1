# Define variables
$logFile = "C:\UpdateLogs\FeatureUpdate_23H2.log"
$updateURL = "https://www.microsoft.com/en-us/software-download/windows23H2" # Placeholder URL, replace with actual
$updateFilePath = "C:\Updates\WindowsFeatureUpdate23H2.exe"

# Function to log messages
function Write-Log {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$timestamp - $message"
}

# Create log directory if it doesn’t exist
if (!(Test-Path "C:\UpdateLogs")) {
    New-Item -ItemType Directory -Path "C:\UpdateLogs"
}

Write-Log "Starting Windows Feature Update 23H2 process."

# Download the update
Write-Log "Downloading the update from $updateURL"
Invoke-WebRequest -Uri $updateURL -OutFile $updateFilePath
Write-Log "Download completed."

# Install the update
Write-Log "Starting installation."
Start-Process -FilePath $updateFilePath -ArgumentList "/quiet /norestart" -Wait
Write-Log "Installation complete."

# Verify installation
$installedVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId
if ($installedVersion -eq "23H2") {
    Write-Log "Installation confirmed: Windows Feature Update 23H2 is installed successfully."
} else {
    Write-Log "Installation verification failed."
}

Write-Log "Update process finished."
