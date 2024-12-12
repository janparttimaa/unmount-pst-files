# PowerShell equivalent of the VBScript code converted by AI
# PowerShell-script unmounts PST-files from Classic Outlook

# Suppress error messages
$ErrorActionPreference = "SilentlyContinue"

# Create an Outlook application object
$objOutlook = New-Object -ComObject Outlook.Application

# Get the list of Outlook stores
$Stores = $objOutlook.Session.Stores

# Loop through the stores
for ($i = $Stores.Count; $i -ge 0; $i--) {
    if ($Stores[$i].ExchangeStoreType -eq 3) {
        # Remove the store if it is an Exchange store
        $objFolder = $Stores[$i].GetRootFolder()
        $objOutlook.Session.RemoveStore($objFolder)
    }
}