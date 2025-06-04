# Install and import the required module
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Import-Module Microsoft.Graph

# Authenticate
Connect-MgGraph -Scopes "Device.ReadWrite.All"


$listOfHostNames = Get-Content -Path "C:\temp\hostnames.txt"
$GroupId = "your-group-id" # Replace with your Entra ID group ID



function Check-DeviceName {
# Define variables
$MatchString = $name # Replace with the ending characters you want to search for
$Devices = Get-MgDevice | Where-Object { $_.DisplayName -like "*$MatchString" }

if ($Devices.Count -eq 1) {
    Write-Host "There are no duplicates for the hostname ending in $($MatchString)!" -ForegroundColor Green
    #exit
}

# Sort devices by registration date (oldest first)
$SortedDevices = $Devices | Sort-Object RegisteredDateTime

# Compare registration dates and prompt for deletion
for ($i = 0; $i -lt $SortedDevices.Count - 1; $i++) {
    $OlderDevice = $SortedDevices[$i]
    $NewerDevice = $SortedDevices[$i + 1]

    Write-Host "Older device found: $($OlderDevice.DisplayName) registered on $($OlderDevice.RegisteredDateTime)" -ForegroundColor Yellow
    $Response = Read-Host "Do you want to delete this device? (Y/N)"

    if ($Response -eq "Y") {
        Remove-MgDevice -DeviceId $OlderDevice.Id
        Write-Host "$($OlderDevice.DisplayName) has been deleted." -ForegroundColor Green
    }
}

Write-Output "Process complete!"


}

function Add-DeviceNameToGroup {
# Define variables
# $GroupId = "your-group-id" # Replace with your Entra ID group ID
# $Hostnames = @("Host1", "Host2", "Host3") # Replace with your list of hostnames

# Retrieve device IDs corresponding to hostnames
$DeviceIds = @()
foreach ($Hostname in $listOfHostName) {
    $Device = Get-MgDevice -Filter "displayName -like '*$Hostname'"
    if ($Device) {
        $DeviceIds += $Device.Id
    }
}

# Add devices to the Entra ID group
foreach ($DeviceId in $DeviceIds) {
    New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $DeviceId
}

Write-Output "Devices added successfully!"



}


foreach($name in $listOfHostNames){

Check-DeviceName


Add-DeviceNameToGroup

}