# Unmount PST-files

This PowerShell-script if forked up from original VBS-script provided by [Diane Poremsky](https://www.slipstick.com/exchange/script-remove-pst-file-profile/). This script have been converted to PowerShell-format because [Microsoft is deprecating VBScript](https://techcommunity.microsoft.com/blog/windows-itpro-blog/vbscript-deprecation-timelines-and-next-steps/4148301).

## Known issues
If running this multiple times to same Windows-device, it will exchaust Outlook so much that user might get error that says Outlook ran out of resources.
