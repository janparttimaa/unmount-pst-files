# Unmount PST-files

This PowerShell-script is forked from original VBS-script provided by [Diane Poremsky](https://www.slipstick.com/exchange/script-remove-pst-file-profile/). This script have been converted to PowerShell-format because [Microsoft is deprecating VBScript](https://techcommunity.microsoft.com/blog/windows-itpro-blog/vbscript-deprecation-timelines-and-next-steps/4148301).

You can deploy this script using either Microsoft Intune or Configuration Manager. In this article we will guide how to deploy this using Intune.

## Known issues
If running this script multiple times to same Windows-device, it will exchaust Outlook so much that user might get error that says Outlook ran out of resources.
Therefore, run this script only once for every user from the device.
