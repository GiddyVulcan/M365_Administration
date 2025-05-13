# Define the path to the .inf file
$infPath = "C:\Path\To\Your\Driver.inf"

# Add the driver package using pnputil
pnputil /add-driver $infPath /install

# Define the printer driver name and printer name
$printerDriverName = "Your Printer Driver Name"
$printerName = "Your Printer Name"

# Install the printer driver
Add-PrinterDriver -Name $printerDriverName

# Add the printer
Add-Printer -Name $printerName -DriverName $printerDriverName -PortName "PORTNAME"

Write-Output "Printer driver and printer installed successfully."
