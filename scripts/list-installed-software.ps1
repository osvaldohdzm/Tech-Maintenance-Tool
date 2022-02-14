
$InstalledSoftware = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
 foreach($obj in $InstalledSoftware){
 $str1 = $obj.GetValue('DisplayName')
 $str2 = $obj.GetValue('DisplayVersion')
 if ($str1) {  Write-Output "$($str1) - $($str2)" >> List2.txt }
 }

 $InstalledSoftware = Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
 foreach($obj in $InstalledSoftware){
 $str1 = $obj.GetValue('DisplayName')
 $str2 = $obj.GetValue('DisplayVersion')
 if ($str1) {  Write-Output "$($str1) - $($str2)" >> List2.txt }
 }


[System.IO.File]::ReadLines('List2.txt') | sort -u | out-file List3.txt -encoding ascii
Remove-Item List2.txt