# Define a log file
$logFile = "Install-GIMP-Log.txt"
Function Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp - $message"
    Write-Output $entry
    $entry | Out-File $logFile -Append
}

# Start logging
Log "Script started."

# Check if Chocolatey is installed
Try {
    Log "Checking if Chocolatey is installed..."
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Log "Chocolatey is not installed. Installing Chocolatey..."
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        Log "Chocolatey installation completed successfully."
    } else {
        Log "Chocolatey is already installed."
    }
} Catch {
    Log "Error during Chocolatey installation: $_"
    Log "Script terminated due to error."
    Exit 1
}

# Install GIMP
Try {
    Log "Starting GIMP installation..."
    choco install gimp -y
    Log "GIMP installation completed successfully."
} Catch {
    Log "Error during GIMP installation: $_"
    Log "Script terminated due to error."
    Exit 1
}

Log "Script finished successfully."

