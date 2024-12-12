# Unmount PST-files for Microsoft Outlook Classic

This PowerShell-script is forked from original VBS-script provided by [Diane Poremsky](https://www.slipstick.com/exchange/script-remove-pst-file-profile/). This script have been converted to PowerShell-format because [Microsoft is deprecating VBScript](https://techcommunity.microsoft.com/blog/windows-itpro-blog/vbscript-deprecation-timelines-and-next-steps/4148301).

You can deploy this script using either Microsoft Intune or Configuration Manager. In this article we will guide how to deploy this using Intune.

## Prerequisites
You have configured these policies via Intune to Microsoft 365 Apps:
- [Prevent users from adding new content to existing PST files](https://admx.help/?Category=Office2016&Policy=outlk16.Office.Microsoft.Policies.Windows::L_Preventusersfromaddingnewcontentto): Enabled
- [Prevent users from adding PSTs to Outlook profiles and/or prevent using Sharing-Exclusive PSTs](https://admx.help/?Category=Office2016&Policy=outlk16.Office.Microsoft.Policies.Windows::L_Preventusersfromaddingpsts): Enabled (No PSTs can be added)

## Deployment instructions
Deploy the PowerShell-script from Intune using Intune's own script deployment function. You need to deploy the script to Security group that includes specific devices that needs this application.

You need to set following settings:
![Screenshot](/img/img01.png)

## Known issues
If running this script multiple times to same Windows-device, it will exchaust Outlook so much that user might get error that says Outlook ran out of resources.
Therefore, run this script only once for every user from the device.
