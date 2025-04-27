# Define a log file for tracking progress
$logFile = "Install-NotepadPlusPlus-Log.txt"
Function Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp - $message"
    Write-Output $entry
    $entry | Out-File $logFile -Append
}

# Start logging
Log "Script started."

# Install Notepad++ using winget
Try {
    Log "Starting installation of Notepad++..."
    winget install --id Notepad++.Notepad++ --silent --accept-source-agreements --accept-package-agreements
    Log "Notepad++ installation completed successfully."
} Catch {
    Log "Error during Notepad++ installation: $_"
}

Log "Script finished."
