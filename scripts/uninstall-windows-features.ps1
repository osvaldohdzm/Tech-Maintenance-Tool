dism /online /get-features

dism /online /Disable-Feature /featurename:MicrosoftWindowsPowerShellV2
dism /online /Disable-Feature /featurename:Printing-XPSServices-Features
dism /online /Disable-Feature /featurename:Printing-Foundation-Features
dism /online /Disable-Feature /featurename:WCF-Services45
dism /online /Disable-Feature /featurename:Printing-Foundation-InternetPrinting-Client
dism /online /Disable-Feature /featurename:WindowsMediaPlayer
dism /online /Disable-Feature /featurename:Internet-Explorer-Optional-amd64
dism /online /Disable-Feature /featurename:MSRDC-Infrastructure
dism /online /Disable-Feature /featurename:WorkFolders-Client
dism /online /Disable-Feature /featurename:MicrosoftWindowsPowerShellV2Root
dism /online /Disable-Feature /featurename:MicrosoftWindowsPowerShellV2

[Console]::ReadKey()