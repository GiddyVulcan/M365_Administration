# Define a log file for tracking progress
$logFile = "Install-Apps-Log.txt"
Function Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp - $message"
    Write-Output $entry
    $entry | Out-File $logFile -Append
}

# Start logging
Log "Script started."

# Define the list of applications to install
$appList = @(
    @{ Name = "Google Chrome"; Id = "Google.Chrome" },
    @{ Name = "Mozilla Firefox"; Id = "Mozilla.Firefox" },
    @{ Name = "GIMP"; Id = "GIMP.GIMP" },
    @{ Name = "Inkscape"; Id = "Inkscape.Inkscape" },
    @{ Name = "Visual Studio Code"; Id = "Microsoft.VisualStudioCode" },
    @{ Name = "Python"; Id = "Python.Python.3" },
    @{ Name = "Microsoft Office 365"; Id = "Microsoft.Office" },
    @{ Name = "PowerShell 7.5"; Id = "Microsoft.Powershell" },
    @{ Name = "Adobe Reader"; Id = "Adobe.Acrobat.Reader.64-bit" }
    @{ Name = "NotePad++"; Id = "Notepad++.Notepad++" }
    @{ Name = "VLC media player"; Id = "VideoLAN.VLC" }
)

# Function to install an application using winget
Function Install-App {
    param([string]$name, [string]$id)

    Try {
        Log "Starting installation of $name..."
        winget install --id=$id --silent --accept-source-agreements --accept-package-agreements --force --scope machine 
        Log "$name installation completed successfully."
    } Catch {
        Log "Error during $name installation: $_"
    }
}

# Install each application in the list
foreach ($app in $appList) {
    Install-App -name $app.Name -id $app.Id
}

Log "Script finished. All applications have been processed."
