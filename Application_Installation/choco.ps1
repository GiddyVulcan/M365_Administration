if (Get-Command choco -ErrorAction SilentlyContinue) {
    $chocoPath = (Get-Command choco).Path
    Write-Host "Chocolatey is installed at: $chocoPath"
} else {
    Write-Host "Chocolatey is not installed."
    # Exit or handle the case where Chocolatey is not installed


    Write-Host "Installing Chocolatey..."
    Set-ExecutionPolicy Bypass -Scope Process -Force    
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))  
    Write-Host "Chocolatey installation completed successfully."
    # Check if Chocolatey is installed successfully
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Host "Chocolatey is now installed."
        $chocoPath = (Get-Command choco).Path
        Write-Host "Chocolatey is installed at: $chocoPath" 
        $envPath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
            if ($envPath -notcontains $chocoInstallDir) {
                [Environment]::SetEnvironmentVariable('Path', "$envPath;$chocoInstallDir", 'Machine')
            Write-Host "Chocolatey added to system PATH."
            $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine')
            (Get-Item Env:Path).Value -split ";" | Select-String "chocolatey"
            } else {
                Write-Host "Chocolatey is already in system PATH."
            }
    } else {
        Write-Host "Chocolatey installation failed."
        Exit 1
    }       
    exit
}