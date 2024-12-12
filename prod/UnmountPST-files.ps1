<#
.SYNOPSIS
    Unmount PST-files from Microsoft Outlook Classic

.DESCRIPTION
    This PowerShell-script is forked from original VBS-script provided by Diane Poremsky. 
    This script have been converted to PowerShell-format because Microsoft is deprecating VBScript.
    Script run will happens per user context.
    NOTE: This scripts have prerequirements that needs to be applied before deploying this script. Please check instructions from GitHub.

.VERSION
    1.0.0

.AUTHOR
    Converted to PowerShell: 2024 Jan Parttimaa (https://github.com/janparttimaa/apple-devices)
    Original script and author: 2018 Diane Poremsky (https://www.slipstick.com/exchange/script-remove-pst-file-profile/)

.COPYRIGHT
    © 2018-2024 Diane Poremsky & Jan Parttimaa. All rights reserved.

.LICENSE
    This script is licensed under the MIT License.
    You may obtain a copy of the License at https://opensource.org/licenses/MIT

.RELEASE NOTES
    1.0.0 - Initial release for PowerShell-format.

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\UnmountPST-files.ps1

    This example is how to run this script running Windows PowerShell. Command needs to be run without admin rights on user context.
#>

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
