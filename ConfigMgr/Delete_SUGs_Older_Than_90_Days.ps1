# Delete-OldUpdateGroups.ps1
# Purpose: Deletes Software Update Groups and their deployments in ConfigMgr that are older than 90 days
# Requirements: 
#   - PowerShell 5.1 or later
#   - ConfigMgr Console installed
#   - ConfigMgr PowerShell module
#   - Run with appropriate permissions to modify ConfigMgr

# Parameters
param (
    [Parameter(Mandatory = $false)]
    [string]$SiteCode = "AUTO", # Will auto-detect if not provided
    
    [Parameter(Mandatory = $false)]
    [string]$SiteServer = "AUTO", # Will auto-detect if not provided
    
    [Parameter(Mandatory = $false)]
    [int]$DaysToKeep = 90, # Number of days to keep SUGs
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$($env:TEMP)\Delete-OldUpdateGroups_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf = $false, # Run in simulation mode without deleting
    
    [Parameter(Mandatory = $false)]
    [switch]$Force = $false # Force deletion even if SUGs are still deployed
)

#region Functions

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$NoConsole
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $LogPath -Value $logEntry
    
    # Write to console with color coding based on level
    if (-not $NoConsole) {
        switch ($Level) {
            "ERROR" { Write-Host $logEntry -ForegroundColor Red }
            "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
            default { Write-Host $logEntry }
        }
    }
}

function Initialize-ConfigMgrConnection {
    Write-Log "Initializing ConfigMgr connection..."
    
    try {
        # Auto-detect site code and server if not provided
        if ($SiteCode -eq "AUTO" -or $SiteServer -eq "AUTO") {
            Write-Log "Attempting to auto-detect Site Code and/or Site Server..."
            
            try {
                $SMSProvider = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ErrorAction Stop
                
                if ($SMSProvider) {
                    if ($SiteCode -eq "AUTO") { 
                        $script:SiteCode = $SMSProvider.SiteCode 
                        Write-Log "Auto-detected Site Code: $SiteCode" 
                    }
                    
                    if ($SiteServer -eq "AUTO") { 
                        $script:SiteServer = $SMSProvider.Machine 
                        Write-Log "Auto-detected Site Server: $SiteServer" 
                    }
                }
                else {
                    throw "Failed to auto-detect Site Code and Site Server"
                }
            }
            catch {
                Write-Log "WMI detection failed. Trying PowerShell module detection..." -Level "WARNING"
                
                # Try to get from existing ConfigMgr module connection
                try {
                    $CMPSDrive = Get-PSDrive -PSProvider CMSite -ErrorAction SilentlyContinue
                    if ($CMPSDrive) {
                        if ($SiteCode -eq "AUTO") {
                            $script:SiteCode = $CMPSDrive.Name
                            Write-Log "Auto-detected Site Code from PSDrive: $SiteCode"
                        }
                        
                        if ($SiteServer -eq "AUTO") {
                            $script:SiteServer = $CMPSDrive.Root
                            Write-Log "Auto-detected Site Server from PSDrive: $SiteServer"
                        }
                    }
                    else {
                        throw "No existing ConfigMgr drive found"
                    }
                }
                catch {
                    Write-Log "Failed to auto-detect ConfigMgr site information: $_" -Level "ERROR"
                    return $false
                }
            }
        }
        
        # Validate site code and server
        if ([string]::IsNullOrEmpty($SiteCode) -or [string]::IsNullOrEmpty($SiteServer)) {
            Write-Log "Site Code or Site Server is empty after auto-detection" -Level "ERROR"
            return $false
        }
        
        # Import ConfigMgr module
        if (-not (Get-Module ConfigurationManager)) {
            $configManagerPath = Join-Path $(Split-Path $ENV:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1
            
            if (Test-Path $configManagerPath) {
                Import-Module $configManagerPath -ErrorAction Stop
                Write-Log "ConfigMgr PowerShell module imported successfully"
            }
            else {
                Write-Log "ConfigMgr console not found. Please make sure the ConfigMgr console is installed." -Level "ERROR"
                return $false
            }
        }
        
        # Connect to site
        if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
            Write-Log "Connected to ConfigMgr site $SiteCode"
        }
        
        # Set current location to the site drive
        Set-Location "$($SiteCode):" -ErrorAction Stop
        Write-Log "Current location set to $($SiteCode):"
        return $true
    }
    catch {
        Write-Log "Error initializing ConfigMgr connection: $_" -Level "ERROR"
        return $false
    }
}

function Get-OldSoftwareUpdateGroups {
    param (
        [int]$DaysOld
    )
    
    Write-Log "Finding Software Update Groups older than $DaysOld days..."
    
    try {
        $cutoffDate = (Get-Date).AddDays(-$DaysOld)
        Write-Log "Using cutoff date: $($cutoffDate.ToString('yyyy-MM-dd'))"
        
        # Get software update groups
        $updateGroups = Get-CMSoftwareUpdateGroup | Where-Object { $_.DateCreated -lt $cutoffDate }
        
        if ($updateGroups -eq $null) {
            Write-Log "No Software Update Groups found older than $DaysOld days" -Level "WARNING"
            return @()
        }
        
        $updateGroups = $updateGroups | Select-Object LocalizedDisplayName, CI_ID, DateCreated, DateLastModified, IsDeployed
        
        Write-Log "Found $($updateGroups.Count) Software Update Groups older than $DaysOld days"
        return $updateGroups
    }
    catch {
        Write-Log "Error finding old Software Update Groups: $_" -Level "ERROR"
        return @()
    }
}

function Get-UpdateGroupDeployments {
    param (
        $UpdateGroup
    )
    
    try {
        # Get all deployments for this update group
        $deployments = Get-CMDeployment -SoftwareUpdateGroupId $UpdateGroup.CI_ID -ErrorAction SilentlyContinue
        
        if ($deployments -and $deployments.Count -gt 0) {
            Write-Log "Update Group '$($UpdateGroup.LocalizedDisplayName)' has $($deployments.Count) active deployments" -NoConsole
            return $deployments
        }
        
        return $null
    }
    catch {
        Write-Log "Error checking deployments for Update Group '$($UpdateGroup.LocalizedDisplayName)': $_" -Level "ERROR" -NoConsole
        return $null
    }
}

function Remove-UpdateGroupDeployments {
    param (
        $UpdateGroup,
        $Deployments,
        [switch]$WhatIf
    )
    
    $groupName = $UpdateGroup.LocalizedDisplayName
    $deploymentsRemoved = 0
    
    foreach ($deployment in $Deployments) {
        try {
            $collectionName = "Unknown"
            try {
                $collection = Get-CMCollection -Id $deployment.CollectionID -ErrorAction SilentlyContinue
                if ($collection) {
                    $collectionName = $collection.Name
                }
            }
            catch {
                # Continue even if we can't get the collection name
            }
            
            if ($WhatIf) {
                Write-Log "WhatIf: Would remove deployment of '$groupName' to collection '$collectionName' (ID: $($deployment.CollectionID))" -Level "WARNING"
                $deploymentsRemoved++
            }
            else {
                Write-Log "Removing deployment of '$groupName' to collection '$collectionName' (ID: $($deployment.CollectionID))..."
                
                # Remove the deployment
                Remove-CMDeployment -DeploymentId $deployment.DeploymentID -Force -ErrorAction Stop
                
                Write-Log "Successfully removed deployment of '$groupName' to collection '$collectionName'" -Level "SUCCESS" -NoConsole
                $deploymentsRemoved++
            }
        }
        catch {
            Write-Log "Error removing deployment for '$groupName' to collection ID $($deployment.CollectionID): $_" -Level "ERROR"
        }
    }
    
    return $deploymentsRemoved
}

function Remove-SoftwareUpdateGroupSafely {
    param (
        $UpdateGroup,
        [switch]$WhatIf,
        [switch]$Force
    )
    
    $groupName = $UpdateGroup.LocalizedDisplayName
    $groupID = $UpdateGroup.CI_ID
    $creationDate = $UpdateGroup.DateCreated.ToString('yyyy-MM-dd')
    
    try {
        # Check for deployments
        $deployments = Get-UpdateGroupDeployments -UpdateGroup $UpdateGroup
        
        # If we have deployments, remove them first
        if ($deployments) {
            Write-Log "Software Update Group '$groupName' (ID: $groupID) has deployments that need to be removed first"
            
            $deploymentsRemoved = Remove-UpdateGroupDeployments -UpdateGroup $UpdateGroup -Deployments $deployments -WhatIf:$WhatIf
            
            if ($WhatIf) {
                Write-Log "WhatIf: Would have removed $deploymentsRemoved deployments for '$groupName'" -Level "WARNING"
            }
            else {
                Write-Log "Removed $deploymentsRemoved deployments for '$groupName'" -Level "SUCCESS"
            }
        }
        
        # Perform the deletion of the SUG
        if ($WhatIf) {
            Write-Log "WhatIf: Would delete Software Update Group '$groupName' (ID: $groupID) created on $creationDate" -Level "WARNING"
            return @{
                Success = $true
                DeploymentsRemoved = if ($deployments) { $deployments.Count } else { 0 }
            }
        }
        else {
            Write-Log "Deleting Software Update Group '$groupName' (ID: $groupID) created on $creationDate..."
            Remove-CMSoftwareUpdateGroup -Id $groupID -Force -ErrorAction Stop
            Write-Log "Successfully deleted Software Update Group '$groupName'" -Level "SUCCESS"
            return @{
                Success = $true
                DeploymentsRemoved = if ($deployments) { $deployments.Count } else { 0 }
            }
        }
    }
    catch {
        Write-Log "Error deleting Software Update Group '$groupName': $_" -Level "ERROR"
        return @{
            Success = $false
            DeploymentsRemoved = 0
        }
    }
}

#endregion Functions

#region Main Script

# Initialize log file
if (-not (Test-Path (Split-Path -Path $LogPath -Parent))) {
    New-Item -Path (Split-Path -Path $LogPath -Parent) -ItemType Directory -Force | Out-Null
}

# Start time for execution tracking
$scriptStartTime = Get-Date
Write-Log "========================================================================"
Write-Log "Delete-OldUpdateGroups script started at $scriptStartTime"
Write-Log "Parameter - DaysToKeep: $DaysToKeep"
Write-Log "Parameter - LogPath: $LogPath"
Write-Log "Parameter - WhatIf: $WhatIf"
Write-Log "Parameter - Force: $Force"
Write-Log "========================================================================"

# Initialize ConfigMgr connection
if (-not (Initialize-ConfigMgrConnection)) {
    Write-Log "Failed to initialize ConfigMgr connection. Exiting script." -Level "ERROR"
    exit 1
}

try {
    # Find old Software Update Groups
    $oldUpdateGroups = Get-OldSoftwareUpdateGroups -DaysOld $DaysToKeep
    
    if ($oldUpdateGroups.Count -eq 0) {
        Write-Log "No Software Update Groups found that are older than $DaysToKeep days. Nothing to delete."
        exit 0
    }
    
    # Process each update group
    $deletedCount = 0
    $skippedCount = 0
    $totalDeploymentsRemoved = 0
    
    Write-Log "========================================================================"
    Write-Log "Starting to process $($oldUpdateGroups.Count) Software Update Groups..."
    Write-Log "========================================================================"
    
    foreach ($updateGroup in $oldUpdateGroups) {
        $result = Remove-SoftwareUpdateGroupSafely -UpdateGroup $updateGroup -WhatIf:$WhatIf -Force:$Force
        
        if ($result.Success) {
            $deletedCount++
            $totalDeploymentsRemoved += $result.DeploymentsRemoved
        }
        else {
            $skippedCount++
        }
    }
    
    # Summary
    Write-Log "========================================================================"
    if ($WhatIf) {
        Write-Log "WHATIF SUMMARY: Would have deleted $deletedCount Software Update Groups with $totalDeploymentsRemoved deployments, skipped $skippedCount" -Level "SUCCESS"
    }
    else {
        Write-Log "EXECUTION SUMMARY: Successfully deleted $deletedCount Software Update Groups and $totalDeploymentsRemoved deployments, skipped $skippedCount" -Level "SUCCESS"
    }
}
catch {
    Write-Log "Unexpected error during script execution: $_" -Level "ERROR"
    exit 1
}
finally {
    # Disconnect from ConfigMgr
    Set-Location $env:SystemDrive
    Write-Log "Disconnected from ConfigMgr site"
    
    $scriptEndTime = Get-Date
    $executionTime = $scriptEndTime - $scriptStartTime
    Write-Log "Script completed in $($executionTime.TotalSeconds.ToString('0.00')) seconds"
    Write-Log "Log file saved to: $LogPath"
    Write-Log "========================================================================"
}

#endregion Main Script