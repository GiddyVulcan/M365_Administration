<#
.SYNOPSIS
    Downloads Notepad++ and packages it as an .intunewin file for Microsoft Intune deployment with comprehensive logging
    
.DESCRIPTION
    This script:
    1. Downloads the latest version of Notepad++
    2. Downloads the Microsoft Win32 Content Prep Tool if not present
    3. Creates an .intunewin package with proper naming convention (VendorName_SoftwareName_VersionNumber.intunewin)
    4. Logs all activities and errors to both console and log file
    
.PARAMETER WorkingDirectory
    Directory where downloads and packaging will occur. Default is user's downloads folder.
    
.PARAMETER ContentPrepUtilPath
    Path where the Content Prep Tool should be downloaded. Default is within working directory.
    
.PARAMETER VendorName
    Vendor name to use in the .intunewin filename. Default is "Notepad++"

.PARAMETER LogPath
    Path where log files will be stored. Default is within working directory.
    
.EXAMPLE
    .\Package-NotepadPlusPlus.ps1
    
.NOTES
    Author: Claude
    Version: 1.1
#>

param (
    [string]$WorkingDirectory = (Join-Path $env:USERPROFILE "Downloads\NotepadPlusPlusPackaging"),
    [string]$ContentPrepUtilPath = (Join-Path $WorkingDirectory "IntuneWinAppUtil"),
    [string]$VendorName = "NotepadPlusPlus",
    [string]$LogPath = (Join-Path $WorkingDirectory "Logs")
)

#Region Functions
function Initialize-Logging {
    param (
        [Parameter(Mandatory=$true)]
        [string]$LogPath
    )
    
    # Create log directory if it doesn't exist
    if (-not (Test-Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }
    
    # Create log file with timestamp
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $global:LogFile = Join-Path $LogPath "NPP_Packaging_$timestamp.log"
    
    # Create the log file and add a header
    $logHeader = @"
======================================================
Notepad++ IntuneWin Packaging Log
Started: $(Get-Date)
Working Directory: $WorkingDirectory
======================================================

"@
    $logHeader | Out-File -FilePath $global:LogFile -Encoding utf8
    
    # Return the log file path
    return $global:LogFile
}

function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory=$false)]
        [switch]$NoConsole
    )
    
    # Define color mapping for console output
    $colorMap = @{
        "INFO" = "White"
        "WARNING" = "Yellow"
        "ERROR" = "Red"
        "SUCCESS" = "Green"
    }
    
    # Create timestamp and format log entry
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    $logEntry | Out-File -FilePath $global:LogFile -Encoding utf8 -Append
    
    # Write to console if not suppressed
    if (-not $NoConsole) {
        Write-Host $logEntry -ForegroundColor $colorMap[$Level]
    }
}

function Test-AdminPrivileges {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Initialize-Directory {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
            Write-Log -Message "Created directory: $Path" -Level "INFO"
        } else {
            Write-Log -Message "Directory already exists: $Path" -Level "INFO"
        }
    }
    catch {
        Write-Log -Message "Failed to create directory $Path`: $_" -Level "ERROR"
        throw $_
    }
}

function Get-NotepadPlusPlusLatestVersion {
    try {
        Write-Log -Message "Determining latest Notepad++ version..." -Level "INFO"
        
        # Get the Notepad++ GitHub releases page
        Write-Log -Message "Connecting to GitHub to check latest release..." -Level "INFO"
        $releasesPage = Invoke-WebRequest -Uri "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/latest" -UseBasicParsing
        
        # Extract the latest version from the redirect URL
        $latestVersion = ($releasesPage.Links | Where-Object { $_.href -match '/tag/v([0-9.]+)' } | Select-Object -First 1).href
        $latestVersion = $latestVersion -replace '.*\/tag\/v', ''
        
        Write-Log -Message "Latest Notepad++ version detected: $latestVersion" -Level "SUCCESS"
        return $latestVersion
    }
    catch {
        Write-Log -Message "Error determining latest Notepad++ version: $_" -Level "ERROR"
        throw $_
    }
}

function Get-NotepadPlusPlusInstaller {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Version,
        
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )
    
    try {
        # Determine architecture
        $arch = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }
        $installerFilename = "npp.$Version.Installer.$arch.exe"
        $downloadUrl = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v$Version/$installerFilename"
        $destinationFile = Join-Path $DestinationPath $installerFilename
        
        Write-Log -Message "System architecture detected: $arch" -Level "INFO"
        Write-Log -Message "Preparing to download: $installerFilename" -Level "INFO"
        Write-Log -Message "Download URL: $downloadUrl" -Level "INFO"
        Write-Log -Message "Destination: $destinationFile" -Level "INFO"
        
        # Check if the file already exists
        if (Test-Path $destinationFile) {
            Write-Log -Message "Installer file already exists. Checking file size..." -Level "INFO"
            
            try {
                $response = Invoke-WebRequest -Uri $downloadUrl -Method Head -UseBasicParsing
                $expectedSize = $response.Headers.'Content-Length'
                $actualSize = (Get-Item $destinationFile).Length
                
                if ($expectedSize -eq $actualSize) {
                    Write-Log -Message "Existing installer file verified. Skipping download." -Level "SUCCESS"
                    return $destinationFile
                } else {
                    Write-Log -Message "Existing installer file size mismatch. Re-downloading..." -Level "WARNING"
                }
            } catch {
                Write-Log -Message "Couldn't verify existing file. Will re-download: $_" -Level "WARNING"
            }
        }
        
        # Download the file
        Write-Log -Message "Downloading Notepad++ installer..." -Level "INFO"
        $progressPreference = 'SilentlyContinue'  # Hide progress bar for faster downloads
        Invoke-WebRequest -Uri $downloadUrl -OutFile $destinationFile -UseBasicParsing
        $progressPreference = 'Continue'  # Restore progress bar
        
        if (Test-Path $destinationFile) {
            $fileSize = [math]::Round((Get-Item $destinationFile).Length / 1MB, 2)
            Write-Log -Message "Notepad++ installer downloaded successfully! Size: $fileSize MB" -Level "SUCCESS"
            return $destinationFile
        }
        else {
            Write-Log -Message "Failed to download Notepad++ installer." -Level "ERROR"
            throw "Download failed: File not found at $destinationFile"
        }
    }
    catch {
        Write-Log -Message "Error downloading Notepad++ installer: $_" -Level "ERROR"
        throw $_
    }
}

function Get-IntuneWinAppUtil {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )
    
    try {
        $intuneWinAppUtilExe = Join-Path $DestinationPath "IntuneWinAppUtil.exe"
        
        # Check if the utility already exists
        if (Test-Path $intuneWinAppUtilExe) {
            Write-Log -Message "IntuneWinAppUtil.exe already exists at: $intuneWinAppUtilExe" -Level "SUCCESS"
            return $intuneWinAppUtilExe
        }
        
        # Download URL for the Content Prep Tool
        $downloadUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
        
        Write-Log -Message "Downloading Microsoft Win32 Content Prep Tool..." -Level "INFO"
        Write-Log -Message "Download URL: $downloadUrl" -Level "INFO"
        Write-Log -Message "Destination: $intuneWinAppUtilExe" -Level "INFO"
        
        # Download the file directly
        $progressPreference = 'SilentlyContinue'  # Hide progress bar for faster downloads
        Invoke-WebRequest -Uri $downloadUrl -OutFile $intuneWinAppUtilExe -UseBasicParsing
        $progressPreference = 'Continue'  # Restore progress bar
        
        if (Test-Path $intuneWinAppUtilExe) {
            $fileSize = [math]::Round((Get-Item $intuneWinAppUtilExe).Length / 1KB, 2)
            Write-Log -Message "Microsoft Win32 Content Prep Tool downloaded successfully! Size: $fileSize KB" -Level "SUCCESS"
            return $intuneWinAppUtilExe
        }
        else {
            Write-Log -Message "Failed to download Microsoft Win32 Content Prep Tool." -Level "ERROR"
            throw "Download failed: File not found at $intuneWinAppUtilExe"
        }
    }
    catch {
        Write-Log -Message "Error downloading Microsoft Win32 Content Prep Tool: $_" -Level "ERROR"
        throw $_
    }
}

function New-IntuneWinPackage {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ContentPrepToolPath,
        
        [Parameter(Mandatory=$true)]
        [string]$SourceFolder,
        
        [Parameter(Mandatory=$true)]
        [string]$SetupFile,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFolder,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFileName
    )
    
    try {
        Write-Log -Message "Creating .intunewin package..." -Level "INFO"
        Write-Log -Message "Content Prep Tool: $ContentPrepToolPath" -Level "INFO"
        Write-Log -Message "Source Folder: $SourceFolder" -Level "INFO"
        Write-Log -Message "Setup File: $SetupFile" -Level "INFO"
        Write-Log -Message "Output Folder: $OutputFolder" -Level "INFO"
        Write-Log -Message "Output File Name: $OutputFileName" -Level "INFO"
        
        # Setup file name only (no path)
        $setupFileName = Split-Path $SetupFile -Leaf
        
        # Use Start-Process to properly handle the IntuneWinAppUtil tool
        $arguments = "-c `"$SourceFolder`" -s `"$setupFileName`" -o `"$OutputFolder`" -q"
        
        Write-Log -Message "Executing IntuneWinAppUtil with arguments: $arguments" -Level "INFO"
        
        # Change to the source folder first (required by IntuneWinAppUtil)
        $currentLocation = Get-Location
        Set-Location -Path $SourceFolder
        
        try {
            $process = Start-Process -FilePath $ContentPrepToolPath -ArgumentList $arguments -NoNewWindow -Wait -PassThru
            $exitCode = $process.ExitCode
            Write-Log -Message "IntuneWinAppUtil process completed with exit code: $exitCode" -Level "INFO"
        }
        finally {
            # Return to original location
            Set-Location -Path $currentLocation
        }
        
        # Check the process exit code
        if ($exitCode -eq 0) {
            # The tool creates files with .intunewin extension but with a different name format
            $defaultOutputFile = Join-Path $OutputFolder "$([System.IO.Path]::GetFileNameWithoutExtension($setupFileName)).intunewin"
            
            # Check if the file was created
            if (Test-Path $defaultOutputFile) {
                Write-Log -Message "IntuneWin package created: $defaultOutputFile" -Level "SUCCESS"
                
                # Rename to our desired format if needed
                $desiredOutputFile = Join-Path $OutputFolder "$OutputFileName.intunewin"
                
                if ($defaultOutputFile -ne $desiredOutputFile) {
                    Write-Log -Message "Renaming output file to match naming convention..." -Level "INFO"
                    Write-Log -Message "New filename: $desiredOutputFile" -Level "INFO"
                    
                    Move-Item -Path $defaultOutputFile -Destination $desiredOutputFile -Force
                    
                    if (Test-Path $desiredOutputFile) {
                        $fileSize = [math]::Round((Get-Item $desiredOutputFile).Length / 1MB, 2)
                        Write-Log -Message "Package renamed successfully. Final size: $fileSize MB" -Level "SUCCESS"
                    } else {
                        Write-Log -Message "Failed to rename the package file." -Level "ERROR"
                        throw "Rename operation failed"
                    }
                }
                
                return $desiredOutputFile
            }
            else {
                Write-Log -Message "IntuneWinAppUtil ran successfully but output file not found at: $defaultOutputFile" -Level "ERROR"
                throw "Output file not found after IntuneWinAppUtil execution"
            }
        }
        else {
            Write-Log -Message "Failed to create .intunewin package. Process exit code: $exitCode" -Level "ERROR"
            throw "IntuneWinAppUtil failed with exit code $exitCode"
        }
    }
    catch {
        Write-Log -Message "Error creating .intunewin package: $_" -Level "ERROR"
        throw $_
    }
}
#EndRegion Functions

#Region Main Script
try {
    # Start logging
    $logFilePath = Initialize-Logging -LogPath $LogPath
    Write-Log -Message "Script execution started" -Level "INFO"
    Write-Log -Message "Log file created at: $logFilePath" -Level "INFO"
    Write-Log -Message "PowerShell Version: $($PSVersionTable.PSVersion)" -Level "INFO"
    Write-Log -Message "Operating System: $([Environment]::OSVersion.VersionString)" -Level "INFO"
    
    # Check for administrator privileges
    Write-Log -Message "Checking administrator privileges..." -Level "INFO"
    if (-not (Test-AdminPrivileges)) {
        Write-Log -Message "This script requires administrator privileges. Please run as administrator." -Level "ERROR"
        exit 1
    }
    Write-Log -Message "Administrator privileges confirmed." -Level "SUCCESS"
    
    # Initialize directories
    Write-Log -Message "Setting up working directories..." -Level "INFO"
    Initialize-Directory -Path $WorkingDirectory
    Initialize-Directory -Path $ContentPrepUtilPath
    Initialize-Directory -Path $LogPath
    
    # Get the latest Notepad++ version
    Write-Log -Message "Finding latest Notepad++ version..." -Level "INFO"
    $notepadPlusPlusVersion = Get-NotepadPlusPlusLatestVersion
    Write-Log -Message "Will package Notepad++ version: $notepadPlusPlusVersion" -Level "INFO"
    
    # Download Notepad++ installer
    Write-Log -Message "Starting Notepad++ installer download..." -Level "INFO"
    $installerPath = Get-NotepadPlusPlusInstaller -Version $notepadPlusPlusVersion -DestinationPath $WorkingDirectory
    Write-Log -Message "Notepad++ installer ready at: $installerPath" -Level "SUCCESS"
    
    # Get the IntuneWinAppUtil tool
    Write-Log -Message "Preparing Content Prep Tool..." -Level "INFO"
    $intuneWinAppUtilPath = Get-IntuneWinAppUtil -DestinationPath $ContentPrepUtilPath
    Write-Log -Message "Content Prep Tool ready at: $intuneWinAppUtilPath" -Level "SUCCESS"
    
    # Create the output filename in the specified format
    $outputFileName = "${VendorName}_NotepadPlusPlus_$notepadPlusPlusVersion"
    Write-Log -Message "Output file name will be: $outputFileName.intunewin" -Level "INFO"
    
    # Create the .intunewin package
    Write-Log -Message "Starting .intunewin package creation..." -Level "INFO"
    $intuneWinPackagePath = New-IntuneWinPackage -ContentPrepToolPath $intuneWinAppUtilPath `
                                                 -SourceFolder $WorkingDirectory `
                                                 -SetupFile $installerPath `
                                                 -OutputFolder $WorkingDirectory `
                                                 -OutputFileName $outputFileName
    Write-Log -Message "IntuneWin package created successfully at: $intuneWinPackagePath" -Level "SUCCESS"
    
    # Display final information
    Write-Log -Message "============= Summary =============" -Level "INFO"
    Write-Log -Message "Notepad++ Version: $notepadPlusPlusVersion" -Level "INFO"
    Write-Log -Message "Installer Location: $installerPath" -Level "INFO"
    Write-Log -Message "IntuneWin Package: $intuneWinPackagePath" -Level "INFO"
    Write-Log -Message "==================================" -Level "INFO"
    
    Write-Log -Message "Process completed successfully!" -Level "SUCCESS"
    Write-Log -Message "Script execution completed" -Level "INFO"
    
    # Return success
    exit 0
}
catch {
    Write-Log -Message "Script failed with unhandled exception: $_" -Level "ERROR"
    Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
    Write-Log -Message "Script execution terminated due to error" -Level "ERROR"
    
    # Return failure
    exit 1
}
#EndRegion Main Script