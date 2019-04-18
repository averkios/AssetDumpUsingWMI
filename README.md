# AssetDumpUsingWMI

## SystemInfo.ps1

SystemInfo.ps1 is a script that automatically gets workstation specs using computer names dump from Microsoft Active Directory. The output is in json format.

Prerequisites :  
If you want to run this script on a workstation you are going to need RSAT  
If you want to run this script on a server you are going to need "Active Directory PowerShell module"  

The script has the following switches :

  - Just run the script without any switches to run it for the whole domain
  - Switch "-Computer" is used to target only specific computers separated with commas followed by spacess
	  - i.e. .\SystemInfo.ps1 -Computer Workstation1 Workstation2
  - Switch "-NotFound" to scan only the workstations that was not found in a previous run
	  - i.e. .\SystemInfo.ps1 -NotFound

When the script runs It creates an error.log file in the  script's root directory that contains any exceptions thrown from the WMI requests.

It also creates the following directories in the script's root location :
```
ScriptRootFolder
	└───SystemInfo
		 ├───Found
		 ├───NotFound
		 └───Notify
```
Directories explained:
```
Found : Contains files for each workstation that has been successfully logged
NotFound : Contains files for each workstation that has not been found, also contains the reason
Notify : Contains files for each workstation that thrown "access denied"
```
Workstation Found json dump sample :
```
[
{
"Name" : "Workstation01",
"Manufacturer" : "Hewlett-Packard",
"Model" : "HP ProDesk 400 G2 MT (TPM DP)",
"Serial Number" : "TSJN9JAEF89N",
"CPU" : "Intel(R) Core(TM) i8-1357L CPU @ 3.00GHz",
"HHDCapacity" : "238",
"HHDFreeSpace" : "24.41 % Free (26.85GB)",
"RAM" : "7.92GB",
"OS" : "Microsoft Windows 10 Enterprise, Service Pack: 0 , 64-bit Operating System",
"UserLoggedIn" : "MyDomain\Username",
"OSInstallDate" : "03/18/2019 13:26:45",
"LastReboot" : "04/17/2019 12:02:18",
"Network Adapter" : "Realtek PCIe GBE Family Controller",
"IPv4/IPv6" : "192.168.1.250 jf80::zx99:sfer:m1b0:zzzz",
"DHCP" : "True",
"MacAddress" : "ZZ:ZZ:00:00:ZZ:ZZ",
"HardwareType" : "Desktop",
"Company" : "Company1",
"MonitorCount" :"2",
"Monitors" : [{"Monitor-Model" : "DELL P2213", "Monitor-SN" : "TSJN9KH#AEF89N"},{"Monitor-Model" : "B22W-6 LED pG", "Monitor-SN" : "TSJN9JER%AEF89N"}]
}
]
```
Workstation Notfound json dump sample :
```
[
{
"Name" : "Workstation2",
"Reason" : "NoWMI"
}
]
```
Workstation Notify json dump sample :
```
[
{
"Name" : "IZN-W7",
"Reason" : "AccessDenied"
}
]
```

## SystemInfoScheduled.ps1

SystemInfoScheduled.ps1 is a variation of the above script with the appropriate modifications to be used as a scheduled task.
> Mostly removed the user interaction.

## Sources

SystemInfo.ps1 is a by-product of the following sources :

- https://4sysops.com/archives/building-a-computer-reporting-script-with-powershell/ (general)
- http://www.admin-magazine.com/Archive/2015/26/Retrieving-Windows-performance-data-in-PowerShell (general)
- https://books.google.gr/books?id=IsVy5iQxx_UC&pg=PA646&lpg=PA646&dq=keep+track+of+queried+computers+powershell&source=bl&ots=3LU5F3C7kO&sig=xDQJBiWu-P49NxK-5xkJ88PqhD0&hl=en&sa=X&ved=0ahUKEwj-l5G6kJHVAhUGKcAKHXsSCS44ChDoAQgdMAE#v=onepage&q&f=false" (optimization)
- https://blogs.technet.microsoft.com/heyscriptingguy/2013/10/03/use-powershell-to-discover-multi-monitor-information/ (monitors)
- https://social.technet.microsoft.com/Forums/ie/en-US/4ec50b2a-e797-4b0b-ad92-b0b6b6fbf788/how-to-get-serial-numbers-of-monitors-from-remote-computers?forum=ITCG (monitors)
- https://msdn.microsoft.com/en-us/library/windows/desktop/aa394542(v=vs.85).aspx (monitors)
- https://github.com/MaxAnderson95/Get-Monitor-Information/blob/master/Get-Monitor.ps1 (monitors final)
- http://mikefrobbins.com/2012/04/20/determine-hardware-type-2012-powershell-scripting-games-beginner-event-8/ (laptop or not)
- https://blogs.technet.microsoft.com/heyscriptingguy/2011/09/21/two-simple-powershell-methods-to-remove-the-last-letter-of-a-string/ (last character removal from string)
- https://docs.microsoft.com/en-us/windows/desktop/cimwin32prov/win32-diskdrive (Hard Disk size)
