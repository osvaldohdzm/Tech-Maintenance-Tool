Get-CimInstance -ClassName Win32_StartupCommand | Select-Object -Property Command, Description, User, Location | Format-List

Get-Item -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run | Select-Object -ExpandProperty Property
Get-Item -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run | Select-Object -ExpandProperty Property
Get-Item -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run | Select-Object -ExpandProperty Property
Get-Item -Path Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce | Select-Object -ExpandProperty Property
Get-Item -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce | Select-Object -ExpandProperty Property
Get-Item -Path Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices | Select-Object -ExpandProperty Property
Get-Item -Path Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce | Select-Object -ExpandProperty Property
Get-Item -Path Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce\Setup | Select-Object -ExpandProperty Property


[Console]::ReadKey()