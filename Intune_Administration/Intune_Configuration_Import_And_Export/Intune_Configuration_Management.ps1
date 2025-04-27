# Intune Configuration Profile Extractor and Importer
# This script extracts all configuration profiles from an Intune tenant and saves them as JSON files
# It can also import configuration profiles from JSON files back into Intune
# Prerequisites: 
# - PowerShell 5.1 or higher
# - Microsoft Graph PowerShell modules (Install-Module Microsoft.Graph)

# Parameters
param (
    [Parameter(Mandatory = $false)]
    [ValidateSet("Export", "Import")]
    [string]$Mode = "Export",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = ".\IntuneConfigProfiles",
    
    [Parameter(Mandatory = $false)]
    [string]$ImportFolder = ".\IntuneConfigProfiles",
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId = "",
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId = "",
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret = "",
    
    [Parameter(Mandatory = $false)]
    [switch]$Force = $false
)

# Function to check if module is installed and import it
function Import-RequiredModule {
    param (
        [string]$ModuleName
    )
    
    if (-not (Get-Module -Name $ModuleName -ListAvailable)) {
        Write-Host "Module $ModuleName is required but not installed. Installing now..."
        Install-Module -Name $ModuleName -Force -Scope CurrentUser
    }
    
    Import-Module -Name $ModuleName -Force
}

# Import required modules
Import-RequiredModule -ModuleName "Microsoft.Graph.Authentication"
Import-RequiredModule -ModuleName "Microsoft.Graph.DeviceManagement"

# Function to connect to Microsoft Graph
function Connect-ToMicrosoftGraph {
    try {
        if ($TenantId -and $ClientId -and $ClientSecret) {
            # App-only authentication
            $securePassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            $clientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $securePassword
            
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $clientSecretCredential -NoWelcome
        } else {
            # Interactive authentication with appropriate scopes
            if ($Mode -eq "Export") {
                Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All" -NoWelcome
            } else {
                Connect-MgGraph -Scopes "DeviceManagementConfiguration.ReadWrite.All" -NoWelcome
            }
        }
        
        Write-Host "Successfully connected to Microsoft Graph API"
        
        # Set the API version to beta to access all profile types
        Select-MgProfile -Name "beta"
    } catch {
        Write-Error "Failed to connect to Microsoft Graph API: $_"
        exit 1
    }
}

# Function to export profiles
function Export-IntuneProfiles {
    # Create output folder if it doesn't exist
    if (-not (Test-Path -Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory | Out-Null
        Write-Host "Created output folder: $OutputFolder"
    }
    
    # Array of profile types to extract
    $profileTypes = @(
        @{Name = "Device Configuration Profiles"; Endpoint = "deviceManagement/deviceConfigurations"; FolderName = "DeviceConfigurationProfiles" },
        @{Name = "Settings Catalog Profiles"; Endpoint = "deviceManagement/configurationPolicies"; FolderName = "SettingsCatalogProfiles" },
        @{Name = "Compliance Policies"; Endpoint = "deviceManagement/deviceCompliancePolicies"; FolderName = "CompliancePolicies" },
        @{Name = "Administrative Templates"; Endpoint = "deviceManagement/groupPolicyConfigurations"; FolderName = "AdministrativeTemplates" },
        @{Name = "Security Baselines"; Endpoint = "deviceManagement/templates"; FolderName = "SecurityBaselines" },
        @{Name = "Endpoint Security Policies"; Endpoint = "deviceManagement/intents"; FolderName = "EndpointSecurityPolicies" }
    )
    
    # Extract each profile type
    foreach ($profileType in $profileTypes) {
        $typeFolderPath = Join-Path -Path $OutputFolder -ChildPath $profileType.FolderName
        
        if (-not (Test-Path -Path $typeFolderPath)) {
            New-Item -Path $typeFolderPath -ItemType Directory | Out-Null
        }
        
        Write-Host "Extracting $($profileType.Name)..."
        
        try {
            # Call Microsoft Graph API to get profiles
            $uri = "https://graph.microsoft.com/beta/$($profileType.Endpoint)"
            $profiles = Invoke-MgGraphRequest -Uri $uri -Method GET
            
            if ($profiles.value) {
                Write-Host "  Found $($profiles.value.Count) profiles"
                
                # Process each profile
                foreach ($profile in $profiles.value) {
                    # Create a valid filename
                    $fileName = "$($profile.id)_$($profile.displayName -replace '[\\\/\:\*\?\"\<\>\|]', '_').json"
                    $filePath = Join-Path -Path $typeFolderPath -ChildPath $fileName
                    
                    # Get detailed profile information
                    $detailUri = "$uri/$($profile.id)"
                    $detailedProfile = Invoke-MgGraphRequest -Uri $detailUri -Method GET
                    
                    # Get assignments for this profile if available
                    try {
                        $assignmentsUri = "$detailUri/assignments"
                        $assignments = Invoke-MgGraphRequest -Uri $assignmentsUri -Method GET
                        if ($assignments.value) {
                            $detailedProfile | Add-Member -MemberType NoteProperty -Name "assignments" -Value $assignments.value -Force
                        }
                    } catch {
                        Write-Verbose "No assignments found for profile: $($profile.displayName)"
                    }
                    
                    # Save profile to JSON file
                    $detailedProfile | ConvertTo-Json -Depth 20 | Out-File -FilePath $filePath -Encoding utf8
                    
                    Write-Host "  Exported: $($profile.displayName)"
                }
            } else {
                Write-Host "  No profiles found"
            }
        } catch {
            Write-Warning "Failed to extract $($profileType.Name): $_"
        }
    }
    
    # Export additional Intune policy types based on assignments
    try {
        $assignmentsFolderPath = Join-Path -Path $OutputFolder -ChildPath "PolicyAssignments"
        
        if (-not (Test-Path -Path $assignmentsFolderPath)) {
            New-Item -Path $assignmentsFolderPath -ItemType Directory | Out-Null
        }
        
        Write-Host "Extracting policy assignments..."
        
        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts"
        $assignments = Invoke-MgGraphRequest -Uri $assignmentsUri -Method GET
        
        if ($assignments.value) {
            foreach ($assignment in $assignments.value) {
                $fileName = "$($assignment.id)_$($assignment.displayName -replace '[\\\/\:\*\?\"\<\>\|]', '_').json"
                $filePath = Join-Path -Path $assignmentsFolderPath -ChildPath $fileName
                
                # Get assignment details
                $detailUri = "$assignmentsUri/$($assignment.id)"
                $detailedAssignment = Invoke-MgGraphRequest -Uri $detailUri -Method GET
                
                # Get the assignments for this policy
                $assignmentDetailsUri = "$detailUri/assignments"
                try {
                    $assignmentDetails = Invoke-MgGraphRequest -Uri $assignmentDetailsUri -Method GET
                    $detailedAssignment | Add-Member -MemberType NoteProperty -Name "assignmentDetails" -Value $assignmentDetails.value
                } catch {
                    Write-Warning "  Failed to get assignment details for $($assignment.displayName): $_"
                }
                
                # Save assignment to JSON file
                $detailedAssignment | ConvertTo-Json -Depth 20 | Out-File -FilePath $filePath -Encoding utf8
                
                Write-Host "  Exported assignment: $($assignment.displayName)"
            }
        } else {
            Write-Host "  No assignments found"
        }
    } catch {
        Write-Warning "Failed to extract policy assignments: $_"
    }
    
    # Output summary
    Write-Host "`nExtraction completed. Files saved to: $OutputFolder"
    Write-Host "Summary:"
    Get-ChildItem -Path $OutputFolder -Directory | ForEach-Object {
        $count = (Get-ChildItem -Path $_.FullName -File).Count
        Write-Host "  $($_.Name): $count profiles"
    }
}

# Function to import profiles
function Import-IntuneProfiles {
    if (-not (Test-Path -Path $ImportFolder)) {
        Write-Error "Import folder does not exist: $ImportFolder"
        exit 1
    }
    
    # Profile type mappings
    $profileTypeMappings = @{
        "DeviceConfigurationProfiles" = @{
            Endpoint = "deviceManagement/deviceConfigurations"
            IdField = "id"
            PropertiesToRemove = @("id", "@odata.type", "createdDateTime", "lastModifiedDateTime", "version")
        }
        "SettingsCatalogProfiles" = @{
            Endpoint = "deviceManagement/configurationPolicies"
            IdField = "id"
            PropertiesToRemove = @("id", "createdDateTime", "lastModifiedDateTime", "version")
        }
        "CompliancePolicies" = @{
            Endpoint = "deviceManagement/deviceCompliancePolicies"
            IdField = "id"
            PropertiesToRemove = @("id", "@odata.type", "createdDateTime", "lastModifiedDateTime", "version")
        }
        "AdministrativeTemplates" = @{
            Endpoint = "deviceManagement/groupPolicyConfigurations"
            IdField = "id"
            PropertiesToRemove = @("id", "createdDateTime", "lastModifiedDateTime", "version")
        }
    }
    
    # Get profile folders
    $profileFolders = Get-ChildItem -Path $ImportFolder -Directory
    
    foreach ($folder in $profileFolders) {
        $folderName = $folder.Name
        
        # Check if we have mapping for this profile type
        if ($profileTypeMappings.ContainsKey($folderName)) {
            $mapping = $profileTypeMappings[$folderName]
            $profileFiles = Get-ChildItem -Path $folder.FullName -Filter "*.json"
            
            Write-Host "Importing $($folderName) ($($profileFiles.Count) files)..."
            
            foreach ($file in $profileFiles) {
                try {
                    # Read the JSON file
                    $profileJson = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
                    $profileName = $profileJson.displayName
                    
                    Write-Host "  Processing: $profileName"
                    
                    # Check if profile already exists
                    $existingProfiles = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/$($mapping.Endpoint)?`$filter=displayName eq '$($profileName -replace "'", "''")'" -Method GET
                    
                    if ($existingProfiles.value -and $existingProfiles.value.Count -gt 0) {
                        if (-not $Force) {
                            Write-Warning "  Profile '$profileName' already exists. Use -Force to overwrite."
                            continue
                        }
                        
                        # Delete existing profile
                        $existingId = $existingProfiles.value[0].id
                        Write-Host "  Replacing existing profile with ID: $existingId"
                        
                        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/$($mapping.Endpoint)/$existingId" -Method DELETE
                    }
                    
                    # Prepare profile for import by removing unnecessary properties
                    $importProfile = $profileJson | ConvertTo-Json -Depth 20 | ConvertFrom-Json
                    
                    # Remove properties that should not be included in creation
                    foreach ($prop in $mapping.PropertiesToRemove) {
                        if ($importProfile.PSObject.Properties.Name -contains $prop) {
                            $importProfile.PSObject.Properties.Remove($prop)
                        }
                    }
                    
                    # Handle assignments separately
                    $assignments = $null
                    if ($importProfile.PSObject.Properties.Name -contains "assignments") {
                        $assignments = $importProfile.assignments
                        $importProfile.PSObject.Properties.Remove("assignments")
                    }
                    
                    # Create new profile
                    $importProfileJson = $importProfile | ConvertTo-Json -Depth 20
                    $newProfile = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/$($mapping.Endpoint)" -Method POST -Body $importProfileJson -ContentType "application/json"
                    
                    Write-Host "  Created profile: $($newProfile.displayName) with ID: $($newProfile.id)"
                    
                    # Apply assignments if available
                    if ($assignments) {
                        Write-Host "  Applying assignments..."
                        
                        foreach ($assignment in $assignments) {
                            # Remove properties that should not be included
                            $assignment.PSObject.Properties.Remove("id")
                            
                            $assignmentJson = $assignment | ConvertTo-Json -Depth 20
                            $assignmentUrl = "https://graph.microsoft.com/beta/$($mapping.Endpoint)/$($newProfile.id)/assignments"
                            
                            try {
                                $result = Invoke-MgGraphRequest -Uri $assignmentUrl -Method POST -Body $assignmentJson -ContentType "application/json"
                                Write-Host "    Applied assignment to: $($assignment.target.groupId)"
                            } catch {
                                Write-Warning "    Failed to apply assignment: $_"
                            }
                        }
                    }
                } catch {
                    Write-Warning "  Failed to import profile $($file.Name): $_"
                }
            }
        } else {
            Write-Warning "Skipping folder $folderName - import not supported for this profile type."
        }
    }
    
    Write-Host "`nImport completed."
}

# Main script execution
Write-Host "Intune Configuration Profile $Mode script"
Write-Host "----------------------------------------"

# Connect to Microsoft Graph
Connect-ToMicrosoftGraph

# Run the appropriate mode
if ($Mode -eq "Export") {
    Export-IntuneProfiles
} else {
    Import-IntuneProfiles
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from Microsoft Graph API"