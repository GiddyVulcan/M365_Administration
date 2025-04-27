# Define a log file for tracking progress
$logFile = "Check-Updates-Log.txt"
Function Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp - $message"
    Write-Output $entry
    $entry | Out-File $logFile -Append
}

# Start logging
Log "Script started."

# Function to check and install Windows Updates
Function Check-WindowsUpdates {
    Try {
        Log "Checking for Windows OS updates..."
        Install-WindowsUpdate -AcceptAll -AutoReboot -ErrorAction Stop
        Log "Windows OS updates check completed successfully."
    } Catch {
        Log "Error during Windows OS updates check: $_"
    }
}

# Function to check for updates for installed apps using winget
Function Check-WingetUpdates {
    Try {
        Log "Checking for application updates using winget..."
        $updates = winget list --upgradable
        Log "Available updates: $updates"
        
        if ($updates -ne $null) {
            $updates | ForEach-Object {
                Try {
                    $appId = $_.Id
                    Log "Updating $($_.Name)..."
                    winget upgrade --id $appId --silent --accept-source-agreements --accept-package-agreements
                    Log "$($_.Name) updated successfully."
                } Catch {
                    Log "Error updating $($_.Name): $_"
                }
            }
        } else {
            Log "No updates available for installed applications."
        }
    } Catch {
        Log "Error during winget updates check: $_"
    }
}

# Ensure Windows Update module is loaded
Try {
    Log "Loading Windows Update module..."
    Import-Module PSWindowsUpdate -ErrorAction Stop
    Log "Windows Update module loaded successfully."
} Catch {
    Log "Error loading Windows Update module: $_"
    Log "Please ensure the PSWindowsUpdate module is installed."
    Exit 1
}

# Execute update checks
Check-WindowsUpdates
Check-WingetUpdates

Log "Script finished successfully."
