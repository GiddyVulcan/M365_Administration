<#
.SYNOPSIS
    Downloads Google Chrome and packages it as an .intunewin file for Microsoft Intune deployment with comprehensive logging
    
.DESCRIPTION
    This script:
    1. Downloads the latest version of Google Chrome Enterprise installer
    2. Downloads the Microsoft Win32 Content Prep Tool if not present
    3. Creates an .intunewin package with proper naming convention (VendorName_SoftwareName_VersionNumber.intunewin)
    4. Logs all activities and errors to both console and log file
    
.PARAMETER WorkingDirectory
    Directory where downloads and packaging will occur. Default is user's downloads folder.
    
.PARAMETER ContentPrepUtilPath
    Path where the Content Prep Tool should be downloaded. Default is within working directory.
    
.PARAMETER VendorName
    Vendor name to use in the .intunewin filename. Default is "Google"

.PARAMETER LogPath
    Path where log files will be stored. Default is within working directory.
    
.EXAMPLE
    .\Package-GoogleChrome.ps1
    
.NOTES
    Author: Claude
    Version: 1.0
#>

param (
    [string]$WorkingDirectory = (Join-Path $env:USERPROFILE "Downloads\GoogleChromePackaging"),
    [string]$ContentPrepUtilPath = (Join-Path $WorkingDirectory "IntuneWinAppUtil"),
    [string]$VendorName = "Google",
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
    $global:LogFile = Join-Path $LogPath "Chrome_Packaging_$timestamp.log"
    
    # Create the log file and add a header
    $logHeader = @"
======================================================
Google Chrome IntuneWin Packaging Log
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

function Get-ChromeVersion {
    try {
        Write-Log -Message "Determining latest Google Chrome version..." -Level "INFO"
        
        # Chrome version API
        $versionUrl = "https://chromereleases.googleblog.com/feeds/posts/default"
        
        Write-Log -Message "Connecting to Google Chrome releases feed..." -Level "INFO"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $webClient = New-Object System.Net.WebClient
        $feed = [xml]$webClient.DownloadString($versionUrl)
        
        # Look for the latest stable version in the feed
        $latestPost = $feed.feed.entry | Select-Object -First 1
        $postContent = $latestPost.content.'#text'
        
        # Extract version number using regex
        $versionPattern = "Chrome ([\d\.]+)"
        $versionMatch = [regex]::Match($postContent, $versionPattern)
        
        if ($versionMatch.Success) {
            $chromeVersion = $versionMatch.Groups[1].Value
            Write-Log -Message "Latest Chrome version detected: $chromeVersion" -Level "SUCCESS"
            return $chromeVersion
        } else {
            # Fallback: If we can't find the version, use direct download and extract from URL
            Write-Log -Message "Version not found in feed, trying direct download URL..." -Level "WARNING"
            
            $downloadUrl = "https://dl.google.com/chrome/install/GoogleChromeStandaloneEnterprise64.msi"
            
            try {
                $request = [System.Net.WebRequest]::Create($downloadUrl)
                $request.Method = "HEAD"
                $response = $request.GetResponse()
                $realURL = $response.ResponseUri.AbsoluteUri
                $response.Close()
                
                # Try to extract version from filename or URL
                $versionPattern = "(\d+\.\d+\.\d+\.\d+)"
                $versionMatch = [regex]::Match($realURL, $versionPattern)
                
                if ($versionMatch.Success) {
                    $chromeVersion = $versionMatch.Groups[1].Value
                    Write-Log -Message "Chrome version from download URL: $chromeVersion" -Level "SUCCESS"
                    return $chromeVersion
                } else {
                    # If still no version, use a placeholder
                    $chromeVersion = "latest"
                    Write-Log -Message "Unable to determine exact version, using '$chromeVersion' as placeholder" -Level "WARNING"
                    return $chromeVersion
                }
            } catch {
                Write-Log -Message "Failed to get version from URL: $_" -Level "WARNING"
                $chromeVersion = "latest"
                Write-Log -Message "Using '$chromeVersion' as placeholder version" -Level "WARNING"
                return $chromeVersion
            }
        }
    }
    catch {
        Write-Log -Message "Error determining Chrome version: $_" -Level "ERROR"
        throw $_
    }
}

function Get-ChromeInstaller {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory=$false)]
        [switch]$X64 = $true
    )
    
    try {
        # Determine architecture
        $arch = if ($X64) { "64" } else { "32" }
        $installerFilename = "GoogleChromeStandaloneEnterprise$arch.msi"
        $downloadUrl = "https://dl.google.com/chrome/install/$installerFilename"
        $destinationFile = Join-Path $DestinationPath $installerFilename
        
        Write-Log -Message "Architecture: $($arch)-bit" -Level "INFO"
        Write-Log -Message "Preparing to download: $installerFilename" -Level "INFO"
        Write-Log -Message "Download URL: $downloadUrl" -Level "INFO"
        Write-Log -Message "Destination: $destinationFile" -Level "INFO"
        
        # Check if the file already exists
        if (Test-Path $destinationFile) {
            Write-Log -Message "Installer file already exists. Checking file age..." -Level "INFO"
            
            $fileAge = (Get-Date) - (Get-Item $destinationFile).LastWriteTime
            # If file is older than 7 days, re-download
            if ($fileAge.TotalDays -gt 7) {
                Write-Log -Message "Existing installer is more than 7 days old. Re-downloading..." -Level "WARNING"
            } else {
                Write-Log -Message "Existing installer is recent. Using cached version." -Level "SUCCESS"
                return $destinationFile
            }
        }
        
        # Download the file
        Write-Log -Message "Downloading Google Chrome installer..." -Level "INFO"
        $progressPreference = 'SilentlyContinue'  # Hide progress bar for faster downloads
        Invoke-WebRequest -Uri $downloadUrl -OutFile $destinationFile -UseBasicParsing
        $progressPreference = 'Continue'  # Restore progress bar
        
        if (Test-Path $destinationFile) {
            $fileSize = [math]::Round((Get-Item $destinationFile).Length / 1MB, 2)
            Write-Log -Message "Google Chrome installer downloaded successfully! Size: $fileSize MB" -Level "SUCCESS"
            return $destinationFile
        }
        else {
            Write-Log -Message "Failed to download Google Chrome installer." -Level "ERROR"
            throw "Download failed: File not found at $destinationFile"
        }
    }
    catch {
        Write-Log -Message "Error downloading Google Chrome installer: $_" -Level "ERROR"
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

function New-ChromeDetectionScript {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory=$true)]
        [string]$Version
    )
    
    try {
        $scriptContent = @"
# Chrome version detection script for Intune
# Detects if Chrome is installed with version $Version or higher

try {
    # Check if Chrome is installed in the default location
    `$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    if (-not (Test-Path `$chromePath)) {
        `$chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    }
    
    if (Test-Path `$chromePath) {
        `$chromeVersion = (Get-Item `$chromePath).VersionInfo.ProductVersion
        `$installedVersion = [version](`$chromeVersion)
        `$requiredVersion = [version]("$Version")
        
        # Return true if installed version is greater than or equal to required version
        if (`$installedVersion -ge `$requiredVersion) {
            Write-Host "Chrome version `$chromeVersion is installed (required: $Version)"
            exit 0
        } else {
            Write-Host "Chrome version `$chromeVersion is installed but does not meet required version $Version"
            exit 1
        }
    } else {
        Write-Host "Chrome is not installed"
        exit 1
    }
} catch {
    Write-Host "Error checking Chrome version: `$_"
    exit 1
}
"@

        $scriptPath = Join-Path $DestinationPath "Detect-Chrome.ps1"
        $scriptContent | Out-File -FilePath $scriptPath -Encoding utf8 -Force
        
        Write-Log -Message "Created Chrome detection script at: $scriptPath" -Level "SUCCESS"
        return $scriptPath
    }
    catch {
        Write-Log -Message "Error creating Chrome detection script: $_" -Level "ERROR"
        throw $_
    }
}

function New-ChromeInstallScript {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )
    
    try {
        $installScript = @"
# Chrome Enterprise installation script for Intune
# Installs Chrome silently with enterprise preferences

# Detect processor architecture
`$Is64Bit = [Environment]::Is64BitOperatingSystem

# Install parameters
`$InstallerName = if (`$Is64Bit) { "GoogleChromeStandaloneEnterprise64.msi" } else { "GoogleChromeStandaloneEnterprise.msi" }
`$InstallArgs = "/quiet /norestart ALLUSERS=1"

# Log file path
`$LogPath = "`$env:TEMP\ChromeInstall_`$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Log function
function Write-Log {
    param (
        [Parameter(Mandatory=`$true)]
        [string]`$Message
    )
    
    `$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "`$Timestamp - `$Message" | Out-File -FilePath `$LogPath -Append -Force
}

try {
    Write-Log "Starting Chrome Enterprise installation"
    Write-Log "Architecture: `$(`$Is64Bit ? '64-bit' : '32-bit')"
    Write-Log "Installer: `$InstallerName"
    
    # Get the directory where this script is located
    `$ScriptDir = Split-Path -Parent `$MyInvocation.MyCommand.Path
    `$InstallerPath = Join-Path `$ScriptDir `$InstallerName
    
    Write-Log "Installer path: `$InstallerPath"
    
    if (Test-Path `$InstallerPath) {
        Write-Log "Found installer at `$InstallerPath"
        
        # Start installation
        Write-Log "Starting installation with arguments: `$InstallArgs"
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"`$InstallerPath`" `$InstallArgs" -Wait -NoNewWindow
        
        # Check if Chrome is installed
        `$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        if (-not (Test-Path `$ChromePath)) {
            `$ChromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        }
        
        if (Test-Path `$ChromePath) {
            `$ChromeVersion = (Get-Item `$ChromePath).VersionInfo.ProductVersion
            Write-Log "Chrome installed successfully: Version `$ChromeVersion"
            exit 0
        } else {
            Write-Log "Installation completed but Chrome executable not found"
            exit 1
        }
    } else {
        Write-Log "ERROR: Installer not found at `$InstallerPath"
        exit 1
    }
} catch {
    Write-Log "ERROR: `$_"
    exit 1
}
"@

        $installScriptPath = Join-Path $DestinationPath "Install-Chrome.ps1"
        $installScript | Out-File -FilePath $installScriptPath -Encoding utf8 -Force
        
        Write-Log -Message "Created Chrome installation script at: $installScriptPath" -Level "SUCCESS"
        return $installScriptPath
    }
    catch {
        Write-Log -Message "Error creating Chrome installation script: $_" -Level "ERROR"
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
    
    # Get the latest Chrome version
    Write-Log -Message "Finding latest Chrome version..." -Level "INFO"
    $chromeVersion = Get-ChromeVersion
    Write-Log -Message "Will package Chrome version: $chromeVersion" -Level "INFO"
    
    # Download Chrome installer (64-bit by default)
    Write-Log -Message "Starting Chrome installer download..." -Level "INFO"
    $installerPath = Get-ChromeInstaller -DestinationPath $WorkingDirectory -X64
    Write-Log -Message "Chrome installer ready at: $installerPath" -Level "SUCCESS"
    
    # Create detection script
    Write-Log -Message "Creating detection script..." -Level "INFO"
    $detectionScriptPath = New-ChromeDetectionScript -DestinationPath $WorkingDirectory -Version $chromeVersion
    
    # Create install script
    Write-Log -Message "Creating installation script..." -Level "INFO"
    $installScriptPath = New-ChromeInstallScript -DestinationPath $WorkingDirectory
    
    # Get the IntuneWinAppUtil tool
    Write-Log -Message "Preparing Content Prep Tool..." -Level "INFO"
    $intuneWinAppUtilPath = Get-IntuneWinAppUtil -DestinationPath $ContentPrepUtilPath
    Write-Log -Message "Content Prep Tool ready at: $intuneWinAppUtilPath" -Level "SUCCESS"
    
    # Create the output filename in the specified format
    $outputFileName = "${VendorName}_Chrome_$chromeVersion"
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
    Write-Log -Message "Chrome Version: $chromeVersion" -Level "INFO"
    Write-Log -Message "Installer Location: $installerPath" -Level "INFO"
    Write-Log -Message "Detection Script: $detectionScriptPath" -Level "INFO"
    Write-Log -Message "Installation Script: $installScriptPath" -Level "INFO"
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