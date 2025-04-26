#Requires -Modules Microsoft.Graph.Intune, ImportExcel, PSWriteWord

<#
.SYNOPSIS
    Creates a comprehensive and visually appealing report for Microsoft Intune tenant.

.DESCRIPTION
    This script connects to Microsoft Graph API, retrieves data about your Intune tenant,
    and generates a Word document with cover page, table of contents, and detailed information 
    about profiles, applications, policies, compliance baselines and other Intune configurations.

.PARAMETER OutputPath
    The path where the report will be saved. Default is the current directory.

.PARAMETER CompanyName
    Your organization name to be displayed on the report cover page.

.PARAMETER ReportTitle
    Custom title for the report. Defaults to "Intune Tenant Configuration Report"

.EXAMPLE
    .\IntuneReportGenerator.ps1 -OutputPath "C:\Reports" -CompanyName "Contoso" -ReportTitle "Quarterly Intune Configuration Report"

.NOTES
    Required modules: Microsoft.Graph.Intune, ImportExcel, PSWriteWord
    Author: Claude
    Version: 1.0
#>

param (
    [string]$OutputPath = (Get-Location).Path,
    [string]$CompanyName = "Your Organization",
    [string]$ReportTitle = "Intune Tenant Configuration Report"
)

# Check for required modules
$requiredModules = @("Microsoft.Graph.Intune", "ImportExcel", "PSWriteWord")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Required module $module is not installed. Installing..." -ForegroundColor Yellow
        Install-Module -Name $module -Force -Scope CurrentUser
    }
}

Import-Module Microsoft.Graph.Intune
Import-Module PSWriteWord
Import-Module ImportExcel

# Create timestamp for the report file
$timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$reportFile = Join-Path -Path $OutputPath -ChildPath "$($CompanyName -replace '\s+', '')_IntuneReport_$timestamp.docx"

# Connect to Microsoft Graph
try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MSGraph | Out-Null
    Write-Host "Connected successfully!" -ForegroundColor Green
}
catch {
    Write-Host "Error connecting to Microsoft Graph: $_" -ForegroundColor Red
    exit
}

# Initialize Word document
Write-Host "Initializing report document..." -ForegroundColor Cyan
$word = New-WordDocument -FilePath $reportFile

# Function to add a section header
function Add-SectionHeader {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [int]$Level = 1
    )
    
    Add-WordText -WordDocument $word -Text $Title -HeadingType "Heading$Level" -Bold $true
    if ($Level -eq 1) {
        Add-WordLine -WordDocument $word -LineType Single -LineColor ([System.Drawing.Color]::FromArgb(0, 113, 197))
    }
}

# Add cover page
$logoPath = $null
# Try to get the company logo if available
# If you have a logo file, uncomment the next line and specify the path
# $logoPath = "C:\Path\To\YourCompanyLogo.png"

Write-Host "Creating cover page..." -ForegroundColor Cyan
Add-WordPageBreak -WordDocument $word
Add-WordText -WordDocument $word -Text $ReportTitle -FontSize 28 -Bold $true -FontColor ([System.Drawing.Color]::FromArgb(0, 113, 197)) -Alignment center
Add-WordText -WordDocument $word -Text $CompanyName -FontSize 20 -Bold $true -Alignment center
Add-WordText -WordDocument $word -Text "Generated on: $(Get-Date -Format 'dddd, MMMM dd, yyyy HH:mm')" -FontSize 14 -Alignment center

# If logo exists, add it to the cover page
if ($logoPath -and (Test-Path $logoPath)) {
    $picture = Add-WordPicture -WordDocument $word -ImagePath $logoPath -Alignment center -Width 300
}

Add-WordText -WordDocument $word -Text "CONFIDENTIAL" -FontSize 12 -Bold $true -FontColor ([System.Drawing.Color]::Red) -Alignment center
Add-WordPageBreak -WordDocument $word

# Add Table of Contents
Write-Host "Adding table of contents..." -ForegroundColor Cyan
Add-WordText -WordDocument $word -Text "Table of Contents" -FontSize 16 -Bold $true
Add-WordTOC -WordDocument $word
Add-WordPageBreak -WordDocument $word

# Add Introduction Section
Add-SectionHeader -Title "Introduction"
Add-WordText -WordDocument $word -Text "This document provides a comprehensive overview of the Intune tenant configuration for $CompanyName. It includes details about device compliance policies, configuration profiles, applications, security baselines, and other key components of the Microsoft Endpoint Manager environment." -FontSize 11
Add-WordText -WordDocument $word -Text "Report generated using automated PowerShell script via Microsoft Graph API." -FontSize 11
Add-WordPageBreak -WordDocument $word

# Gather Intune data
Write-Host "Gathering Intune tenant data..." -ForegroundColor Cyan

# Get tenant details
try {
    $organization = Invoke-MSGraphRequest -HttpMethod GET -Url "organization" | Get-MSGraphAllPages
    $tenantName = $organization.displayName
    $tenantId = $organization.id
    
    Add-SectionHeader -Title "Tenant Information"
    Add-WordText -WordDocument $word -Text "Tenant Name: $tenantName" -FontSize 11
    Add-WordText -WordDocument $word -Text "Tenant ID: $tenantId" -FontSize 11
    Add-WordPageBreak -WordDocument $word
}
catch {
    Write-Host "Error retrieving tenant information: $_" -ForegroundColor Yellow
    Add-WordText -WordDocument $word -Text "Error retrieving tenant information." -FontSize 11 -FontColor ([System.Drawing.Color]::Red)
}

# Get Device Configuration Profiles
try {
    Write-Host "Retrieving device configuration profiles..." -ForegroundColor Cyan
    $deviceConfigs = Invoke-MSGraphRequest -HttpMethod GET -Url "deviceManagement/deviceConfigurations" | Get-MSGraphAllPages
    
    Add-SectionHeader -Title "Device Configuration Profiles"
    Add-WordText -WordDocument $word -Text "Total Profiles: $($deviceConfigs.Count)" -FontSize 11
    
    if ($deviceConfigs.Count -gt 0) {
        $configTable = New-Object System.Collections.ArrayList
        foreach ($config in $deviceConfigs) {
            $null = $configTable.Add([PSCustomObject]@{
                Name = $config.displayName
                Type = $config.'@odata.type' -replace '#microsoft.graph.', ''
                Platform = if ($config.platformType) { $config.platformType } else { "Multiple" }
                CreatedDateTime = $config.createdDateTime
                LastModifiedDateTime = $config.lastModifiedDateTime
            })
        }
        
        Add-WordTable -WordDocument $word -DataTable $configTable -Design LightGridAccent5 -AutoFit Window
        
        # Add detail section for each configuration profile
        foreach ($config in $deviceConfigs) {
            Add-SectionHeader -Title $config.displayName -Level 2
            Add-WordText -WordDocument $word -Text "Profile Type: $($config.'@odata.type' -replace '#microsoft.graph.', '')" -FontSize 11
            Add-WordText -WordDocument $word -Text "Description: $($config.description)" -FontSize 11
            Add-WordText -WordDocument $word -Text "Created: $($config.createdDateTime)" -FontSize 11
            Add-WordText -WordDocument $word -Text "Last Modified: $($config.lastModifiedDateTime)" -FontSize 11
            
            # You can add more detailed settings here based on the profile type
        }
    }
    else {
        Add-WordText -WordDocument $word -Text "No configuration profiles found." -FontSize 11
    }
    Add-WordPageBreak -WordDocument $word
}
catch {
    Write-Host "Error retrieving device configuration profiles: $_" -ForegroundColor Yellow
    Add-WordText -WordDocument $word -Text "Error retrieving device configuration profiles." -FontSize 11 -FontColor ([System.Drawing.Color]::Red)
}

# Get Compliance Policies
try {
    Write-Host "Retrieving device compliance policies..." -ForegroundColor Cyan
    $compliancePolicies = Invoke-MSGraphRequest -HttpMethod GET -Url "deviceManagement/deviceCompliancePolicies" | Get-MSGraphAllPages
    
    Add-SectionHeader -Title "Device Compliance Policies"
    Add-WordText -WordDocument $word -Text "Total Policies: $($compliancePolicies.Count)" -FontSize 11
    
    if ($compliancePolicies.Count -gt 0) {
        $complianceTable = New-Object System.Collections.ArrayList
        foreach ($policy in $compliancePolicies) {
            $null = $complianceTable.Add([PSCustomObject]@{
                Name = $policy.displayName
                Type = $policy.'@odata.type' -replace '#microsoft.graph.', ''
                Created = $policy.createdDateTime
                LastModified = $policy.lastModifiedDateTime
            })
        }
        
        Add-WordTable -WordDocument $word -DataTable $complianceTable -Design LightGridAccent5 -AutoFit Window
        
        # Add detail section for each compliance policy
        foreach ($policy in $compliancePolicies) {
            Add-SectionHeader -Title $policy.displayName -Level 2
            Add-WordText -WordDocument $word -Text "Policy Type: $($policy.'@odata.type' -replace '#microsoft.graph.', '')" -FontSize 11
            Add-WordText -WordDocument $word -Text "Description: $($policy.description)" -FontSize 11
            Add-WordText -WordDocument $word -Text "Created: $($policy.createdDateTime)" -FontSize 11
            Add-WordText -WordDocument $word -Text "Last Modified: $($policy.lastModifiedDateTime)" -FontSize 11
            
            # Add specific compliance settings based on policy type
        }
    }
    else {
        Add-WordText -WordDocument $word -Text "No compliance policies found." -FontSize 11
    }
    Add-WordPageBreak -WordDocument $word
}
catch {
    Write-Host "Error retrieving compliance policies: $_" -ForegroundColor Yellow
    Add-WordText -WordDocument $word -Text "Error retrieving compliance policies." -FontSize 11 -FontColor ([System.Drawing.Color]::Red)
}

# Get Applications
try {
    Write-Host "Retrieving mobile applications..." -ForegroundColor Cyan
    $applications = Invoke-MSGraphRequest -HttpMethod GET -Url "deviceAppManagement/mobileApps" | Get-MSGraphAllPages
    
    Add-SectionHeader -Title "Applications"
    Add-WordText -WordDocument $word -Text "Total Applications: $($applications.Count)" -FontSize 11
    
    if ($applications.Count -gt 0) {
        # Filter out system apps and group by type
        $userApps = $applications | Where-Object { $_.isAssigned -eq $true -or $null -eq $_.isAssigned }
        
        $appsByType = $userApps | Group-Object -Property '@odata.type'
        
        foreach ($appType in $appsByType) {
            $friendlyTypeName = $appType.Name -replace '#microsoft.graph.', ''
            
            Add-SectionHeader -Title "Application Type: $friendlyTypeName" -Level 2
            
            $appsTable = New-Object System.Collections.ArrayList
            foreach ($app in $appType.Group) {
                $null = $appsTable.Add([PSCustomObject]@{
                    Name = $app.displayName
                    Publisher = $app.publisher
                    Featured = if ($app.isFeatured) { "Yes" } else { "No" }
                    Version = if ($app.version) { $app.version } else { "N/A" }
                })
            }
            
            Add-WordTable -WordDocument $word -DataTable $appsTable -Design LightGridAccent5 -AutoFit Window
        }
    }
    else {
        Add-WordText -WordDocument $word -Text "No applications found." -FontSize 11
    }
    Add-WordPageBreak -WordDocument $word
}
catch {
    Write-Host "Error retrieving applications: $_" -ForegroundColor Yellow
    Add-WordText -WordDocument $word -Text "Error retrieving applications." -FontSize 11 -FontColor ([System.Drawing.Color]::Red)
}

# Get Security Baselines
try {
    Write-Host "Retrieving security baselines..." -ForegroundColor Cyan
    $securityBaselines = Invoke-MSGraphRequest -HttpMethod GET -Url "deviceManagement/templates?`$filter=startswith(id,'securityBaseline')" | Get-MSGraphAllPages
    
    Add-SectionHeader -Title "Security Baselines"
    
    if ($securityBaselines.Count -gt 0) {
        Add-WordText -WordDocument $word -Text "Total Security Baselines: $($securityBaselines.Count)" -FontSize 11
        
        $baselineTable = New-Object System.Collections.ArrayList
        foreach ($baseline in $securityBaselines) {
            $null = $baselineTable.Add([PSCustomObject]@{
                Name = $baseline.displayName
                ID = $baseline.id
                Version = $baseline.versionInfo
            })
        }
        
        Add-WordTable -WordDocument $word -DataTable $baselineTable -Design LightGridAccent5 -AutoFit Window
    }
    else {
        Add-WordText -WordDocument $word -Text "No security baselines found." -FontSize 11
    }
    Add-WordPageBreak -WordDocument $word
}
catch {
    Write-Host "Error retrieving security baselines: $_" -ForegroundColor Yellow
    Add-WordText -WordDocument $word -Text "Error retrieving security baselines." -FontSize 11 -FontColor ([System.Drawing.Color]::Red)
}

# Get Device Enrollment Configurations
try {
    Write-Host "Retrieving device enrollment configurations..." -ForegroundColor Cyan
    $enrollmentConfigs = Invoke-MSGraphRequest -HttpMethod GET -Url "deviceManagement/deviceEnrollmentConfigurations" | Get-MSGraphAllPages
    
    Add-SectionHeader -Title "Device Enrollment Configurations"
    
    if ($enrollmentConfigs.Count -gt 0) {
        Add-WordText -WordDocument $word -Text "Total Enrollment Configurations: $($enrollmentConfigs.Count)" -FontSize 11
        
        $enrollmentTable = New-Object System.Collections.ArrayList
        foreach ($config in $enrollmentConfigs) {
            $null = $enrollmentTable.Add([PSCustomObject]@{
                Name = $config.displayName
                Type = $config.'@odata.type' -replace '#microsoft.graph.', ''
                Priority = $config.priority
                Created = $config.createdDateTime
            })
        }
        
        Add-WordTable -WordDocument $word -DataTable $enrollmentTable -Design LightGridAccent5 -AutoFit Window
    }
    else {
        Add-WordText -WordDocument $word -Text "No enrollment configurations found." -FontSize 11
    }
    Add-WordPageBreak -WordDocument $word
}
catch {
    Write-Host "Error retrieving enrollment configurations: $_" -ForegroundColor Yellow
    Add-WordText -WordDocument $word -Text "Error retrieving enrollment configurations." -FontSize 11 -FontColor ([System.Drawing.Color]::Red)
}

# Get Conditional Access Policies (if permissions allow)
try {
    Write-Host "Retrieving conditional access policies..." -ForegroundColor Cyan
    $conditionalAccessPolicies = Invoke-MSGraphRequest -HttpMethod GET -Url "identity/conditionalAccess/policies" | Get-MSGraphAllPages
    
    Add-SectionHeader -Title "Conditional Access Policies"
    
    if ($conditionalAccessPolicies.Count -gt 0) {
        Add-WordText -WordDocument $word -Text "Total Conditional Access Policies: $($conditionalAccessPolicies.Count)" -FontSize 11
        
        $caTable = New-Object System.Collections.ArrayList
        foreach ($policy in $conditionalAccessPolicies) {
            $null = $caTable.Add([PSCustomObject]@{
                Name = $policy.displayName
                State = $policy.state
                CreatedDateTime = $policy.createdDateTime
                ModifiedDateTime = $policy.modifiedDateTime
            })
        }
        
        Add-WordTable -WordDocument $word -DataTable $caTable -Design LightGridAccent5 -AutoFit Window
    }
    else {
        Add-WordText -WordDocument $word -Text "No conditional access policies found." -FontSize 11
    }
    Add-WordPageBreak -WordDocument $word
}
catch {
    # This might fail if the user doesn't have permissions to CA policies
    Write-Host "Error or insufficient permissions for retrieving conditional access policies: $_" -ForegroundColor Yellow
    Add-WordText -WordDocument $word -Text "Error or insufficient permissions for retrieving conditional access policies." -FontSize 11 -FontColor ([System.Drawing.Color]::Red)
}

# Add summary section
Add-SectionHeader -Title "Summary and Recommendations"
Add-WordText -WordDocument $word -Text "This report provides a comprehensive overview of your Intune tenant configuration as of $(Get-Date). Review the configurations regularly to ensure they align with your organization's security policies and business requirements." -FontSize 11

Add-WordText -WordDocument $word -Text "Key statistics:" -FontSize 11 -Bold $true

# Create a summary list
$summaryStats = @()
if (Get-Variable -Name deviceConfigs -ErrorAction SilentlyContinue) { $summaryStats += "Device Configuration Profiles: $($deviceConfigs.Count)" }
if (Get-Variable -Name compliancePolicies -ErrorAction SilentlyContinue) { $summaryStats += "Compliance Policies: $($compliancePolicies.Count)" }
if (Get-Variable -Name applications -ErrorAction SilentlyContinue) { $summaryStats += "Applications: $($applications.Count)" }
if (Get-Variable -Name securityBaselines -ErrorAction SilentlyContinue) { $summaryStats += "Security Baselines: $($securityBaselines.Count)" }
if (Get-Variable -Name enrollmentConfigs -ErrorAction SilentlyContinue) { $summaryStats += "Enrollment Configurations: $($enrollmentConfigs.Count)" }
if (Get-Variable -Name conditionalAccessPolicies -ErrorAction SilentlyContinue) { $summaryStats += "Conditional Access Policies: $($conditionalAccessPolicies.Count)" }

# Add the summary as a bulleted list
foreach ($stat in $summaryStats) {
    Add-WordText -WordDocument $word -Text $stat -FontSize 11 -Bullet
}

# Save the document
try {
    Write-Host "Saving report to $reportFile..." -ForegroundColor Cyan
    Save-WordDocument -WordDocument $word -FilePath $reportFile -Supress $true
    Write-Host "Report saved successfully!" -ForegroundColor Green
    Write-Host "Report location: $reportFile" -ForegroundColor Green
    
    # Open the report
    Invoke-Item $reportFile
}
catch {
    Write-Host "Error saving report: $_" -ForegroundColor Red
}

# Disconnect from Microsoft Graph
Disconnect-Graph | Out-Null
Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Cyan