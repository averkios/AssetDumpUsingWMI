#Catch passed variables from powershell
param([String[]] $Computers, [switch]$NotFound)

#make all errors terminating to catch access denied error from wmi
$erroractionPreference

#clear screen
Clear-Host

#To mute all errors uncomment the following line
#$ErrorActionPreference= 'silentlycontinue'

#Set the folder paths relatively to the script position
$FolderRoot = ($PSScriptRoot + "\SystemInfo")
$OutputPathRoot = ($PSScriptRoot + "\SystemInfo\Found\")
$NotFoundPathRoot = ($PSScriptRoot + "\SystemInfo\NotFound\")
$NoAccessPathRoot = ($PSScriptRoot + "\SystemInfo\Notify\")

#Create directory structure
$PathArray = $FolderRoot, $OutputPathRoot, $NotFoundPathRoot, $NoAccessPathRoot
foreach ($path in $pathArray)
{
    if (! (Test-Path $path))
    {
        New-Item -ItemType Directory -Force -path $path
    }
}

#Clear the error log
if (Test-Path ($PSScriptRoot + '\SystemInfo\Error.log'))
{
    Clear-Content ($PSScriptRoot + '\SystemInfo\Error.log')
}

#Suppress Clear-Variable cmdlet errors of unset variables
Set-Variable -name OutputPath
Set-Variable -name NotFoundPath
Set-Variable -name NoAccessPath

function GetInfo
{
    foreach ($Computer in $ArrComputers) 
    {
        #Clear error log
        $Error.Clear()
        
        #Clear path variables
        Clear-Variable -name OutputPath
        Clear-Variable -name NotFoundPath
        Clear-Variable -name NoAccessPath
    
        #Set found path and file name
        $OutputPath = ($PSScriptRoot + "\SystemInfo\Found\" + $Computer + ".txt")
        #Set not found path and file name
        $NotFoundPath = ($PSScriptRoot + "\SystemInfo\NotFound\" + $Computer + ".txt")
        #Set no access path and file name
        $NoAccessPath = ($PSScriptRoot + "\SystemInfo\Notify\" + $Computer + ".txt")

        #jump to next computer on first wmi fail and write to NotFound or no access folder
        try {
            $computerSystem = get-wmiobject Win32_ComputerSystem -Computer $Computer -erroraction 'stop'
        }
        catch [Exception]
        {
            #Extract info from catching the exception
            $ExceptionErrorInfo = $_.Exception
            $ExceptionName = $ExceptionErrorInfo.GetType().Name

            if ($ExceptionName -eq "COMException") 
            {
               #write errors to the error log
               "Timestamp : " + "$(Get-Date)" + "`t" + "|`tComputerName : " + $Computer + "`t|`t`tError : " + $Error | Add-Content ( $PSScriptRoot + '\SystemInfo\Error.log')
        
               #Test if computer record exists in NotFound folder. if it exists clear contents else create a new.
               if (! (Test-Path $NotFoundPath) )
               {
                    New-Item $NotFoundPath -type file
                    "[" | Add-Content $NotFoundPath
                    "{" | Add-Content $NotFoundPath
                    '"Name" : ' + '"' + $Computer + '",' | Add-Content $NotFoundPath
                    '"Reason" : "NoWMI"' | Add-Content $NotFoundPath
                    "}" | Add-Content $NotFoundPath
                    "]" | Add-Content $NotFoundPath
                    continue
                } 
                else
                {
                    Clear-Content $NotFoundPath
                    "[" | Add-Content $NotFoundPath
                    "{" | Add-Content $NotFoundPath
                    '"Name" : ' + '"' + $Computer + '",' | Add-Content $NotFoundPath
                    '"Reason" : "NoWMI"' | Add-Content $NotFoundPath
                    "}" | Add-Content $NotFoundPath
                    "]" | Add-Content $NotFoundPath
                    continue
                } 
            }
            elseif ($ExceptionName -eq "UnauthorizedAccessException")
            {
               Write-Host "`nAccess Denied on computer : $Computer`n"
               #write errors to the error log
               "Timestamp : " + "$(Get-Date)" + "`t" + "|`tComputerName : " + $Computer + "`t|`t`tError : " + $ExceptionName | Add-Content ( $PSScriptRoot + '\SystemInfo\Error.log')
        
               #Test if computer record exists in NotFound folder. if it exists clear contents else create a new.
               if (! (Test-Path $NoAccessPath) )
               {
                    New-Item $NoAccessPath -type file
                    "[" | Add-Content $NoAccessPath
                    "{" | Add-Content $NoAccessPath
                    '"Name" : ' + '"' + $Computer + '",' | Add-Content $NoAccessPath
                    '"Reason" : "AccessDenied"' | Add-Content $NoAccessPath
                    "}" | Add-Content $NoAccessPath
                    "]" | Add-Content $NoAccessPath

                    #Remove the item from NotFound folder, user intervention needed
                    if(Test-Path $NotFoundPath)
                    {
                        Remove-Item $NotFoundPath
                    }
                    continue
                } 
                else
                {
                    Clear-Content $NoAccessPath
                    "[" | Add-Content $NoAccessPath
                    "{" | Add-Content $NoAccessPath
                    '"Name" : ' + '"' + $Computer + '",' | Add-Content $NoAccessPath
                    '"Reason" : "AccessDenied"' | Add-Content $NoAccessPath
                    "}" | Add-Content $NoAccessPath
                    "]" | Add-Content $NoAccessPath
                    continue
                }
            }
            else {
                write-host "Something unexpected happened :`n"
                Write-Host $ExceptionErrorInfo
                break
            }

            
        }

        #Remove computer record from NotFound folder if existent
        if(Test-Path $NotFoundPath)
        {
            Remove-Item $NotFoundPath
        }

        #clear previous Output Json file or create new
        if (! (Test-Path $OutputPath) )
        {
            #Output Json file creation start
            New-Item $OutputPath -type file
            "[" | Add-Content $OutputPath
        }
        else {
            #Output Json file creation start
            Clear-Content $OutputPath
            "[" | Add-Content $OutputPath   
        }

        #continue to dump wmi into values
        $computerBIOS = Get-WmiObject Win32_Bios -ComputerName $Computer
        $computerOS = Get-wmiobject Win32_OperatingSystem -Computer $Computer
        $computerCPU = Get-wmiobject Win32_Processor -Computer $Computer
        $computerHDD = Get-WmiObject Win32_LogicalDisk -ComputerName $Computer -Filter drivetype=3
        $Network = Get-WmiObject win32_networkadapterconfiguration -ComputerName $Computer -Filter 'ipenabled = "true"' | select Description, DHCPEnabled, IPAddress, Macaddress
        $Gfx = Get-WmiObject win32_videocontroller -ComputerName $Computer | select Description, CurrentHorizontalResolution, CurrentVerticalResolution

        #calculate some values
        
        #DiskSize of drive 0, usually drive C
        $HDDCapacity = Get-WMIObject win32_diskdrive  -ComputerName $Computer | Where-Object DeviceID -like "\\.\PHYSICALDRIVE0" | Select Size -ExpandProperty Size
        #Make it human readable and round it
        $HDDCapacity = [math]::Round(($HDDCapacity/1GB))

        #Partition Size
        $HDDPartitionCapacity = "{0:N2}" -f [double](($computerHDD | Measure-Object Size -Sum).Sum/1GB) + "GB"
        
        $HDDFreeSpace = "{0:P2}" -f ((($computerhdd | Measure-Object freespace -sum).sum/1GB) / [double](($computerHDD | Measure-Object Size -Sum).Sum/1GB)) + " Free (" + "{0:N2}" -f (($computerhdd | Measure-Object freespace -sum).sum/1GB) + "GB)"
        $RAM = "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"
        $OS = $computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion + " , " + $computerOS.OSArchitecture + " Operating System"
  
        #json info write start    
        "{" | Add-Content $OutputPath
        '"Name" : ' + '"' + $computerSystem.Name + '",' | Add-Content $OutputPath
        '"Manufacturer" : ' + '"' + $computerSystem.Manufacturer  + '",' | Add-Content $OutputPath
        '"Model" : ' + '"' + $computerSystem.Model  + '",' | Add-Content $OutputPath
        '"Serial Number" : ' + '"' + $computerBIOS.SerialNumber  + '",' | Add-Content $OutputPath
        '"CPU" : ' + '"' + $computerCPU.Name  + '",' | Add-Content $OutputPath
        '"HHDCapacity" : ' + '"' + $HDDCapacity  + '",' | Add-Content $OutputPath
        '"HHDFreeSpace" : ' + '"' + $HDDFreeSpace  + '",' | Add-Content $OutputPath
        '"RAM" : ' + '"' + $RAM  + '",' | Add-Content $OutputPath
        '"OS" : ' + '"' + $OS  + '",' | Add-Content $OutputPath
        '"UserLoggedIn" : ' + '"' + $computerSystem.UserName  + '",' | Add-Content $OutputPath
        '"OSInstallDate" : ' + '"' + $computerOS.ConvertToDateTime($computerOS.InstallDate)  + '",' | Add-Content $OutputPath
        '"LastReboot" : ' + '"' + $computerOS.ConvertToDateTime($computerOS.LastBootUpTime, "yyyy/MM/dd")  + '",' | Add-Content $OutputPath
        '"Network Adapter" : ' + '"' + $Network.Description  + '",' | Add-Content $OutputPath
        '"IPv4/IPv6" : ' + '"' + $Network.IPAddress  + '",' | Add-Content $OutputPath
        '"DHCP" : ' + '"' + $Network.DHCPEnabled  + '",' | Add-Content $OutputPath
        '"MacAddress" : ' + '"' + $Network.MacAddress  + '",' | Add-Content $OutputPath
    
        #DumpHardwareType
        $hardwaretype = $computerSystem.PCSystemType
    
        if ($hardwaretype -ne 2)
        {
            '"HardwareType" : "Desktop",' | Add-Content $OutputPath
        }
        else {
            '"HardwareType" : "Laptop",' | Add-Content $OutputPath
        }

        #Get OU Distinguished Name
        $DistinguishedNameVar = Get-ADComputer -Identity $Computer | Select-Object -ExpandProperty DistinguishedName
    
        #Get the company
        switch ($DistinguishedNameVar)
        {
            {$_ -like '*Company1*'} {'"Company" : "Company1",' | Add-Content $OutputPath}
            {$_ -like '*Company2*'} {'"Company" : "Company2",' | Add-Content $OutputPath}
            {$_ -like '*Company3*'} {'"Company" : "Company3",' | Add-Content $OutputPath}
            default {'"Company" : "Wrong OU",'| Add-Content $OutputPath}
        }
 
        #Monitor Detailed Info start
        #List of Manufacture Codes that could be pulled from WMI and their respective full names. Used for translating later down.
        $ManufacturerHash = @{ 
            "AAC" =	"AcerView";
            "ACR" = "Acer";
            "AOC" = "AOC";
            "AIC" = "AG Neovo";
            "APP" = "Apple Computer";
            "AST" = "AST Research";
            "AUO" = "Asus";
            "BNQ" = "BenQ";
            "CMO" = "Acer";
            "CPL" = "Compal";
            "CPQ" = "Compaq";
            "CPT" = "Chunghwa Pciture Tubes, Ltd.";
            "CTX" = "CTX";
            "DEC" = "DEC";
            "DEL" = "Dell";
            "DPC" = "Delta";
            "DWE" = "Daewoo";
            "EIZ" = "EIZO";
            "ELS" = "ELSA";
            "ENC" = "EIZO";
            "EPI" = "Envision";
            "FCM" = "Funai";
            "FUJ" = "Fujitsu";
            "FUS" = "Fujitsu-Siemens";
            "GSM" = "LG Electronics";
            "GWY" = "Gateway 2000";
            "HEI" = "Hyundai";
            "HIT" = "Hyundai";
            "HSL" = "Hansol";
            "HTC" = "Hitachi/Nissei";
            "HWP" = "HP";
            "IBM" = "IBM";
            "ICL" = "Fujitsu ICL";
            "IVM" = "Iiyama";
            "KDS" = "Korea Data Systems";
            "LEN" = "Lenovo";
            "LGD" = "Asus";
            "LPL" = "Fujitsu";
            "MAX" = "Belinea"; 
            "MEI" = "Panasonic";
            "MEL" = "Mitsubishi Electronics";
            "MS_" = "Panasonic";
            "NAN" = "Nanao";
            "NEC" = "NEC";
            "NOK" = "Nokia Data";
            "NVD" = "Fujitsu";
            "OPT" = "Optoma";
            "PHL" = "Philips";
            "REL" = "Relisys";
            "SAN" = "Samsung";
            "SAM" = "Samsung";
            "SBI" = "Smarttech";
            "SGI" = "SGI";
            "SNY" = "Sony";
            "SRC" = "Shamrock";
            "SUN" = "Sun Microsystems";
            "SEC" = "Hewlett-Packard";
            "TAT" = "Tatung";
            "TOS" = "Toshiba";
            "TSB" = "Toshiba";
            "VSC" = "ViewSonic";
            "ZCM" = "Zenith";
            "UNK" = "Unknown";
            "_YV" = "Fujitsu";
          }
      
        #Grabs the Monitor objects from WMI
        $Monitors = Get-WmiObject -Namespace "root\WMI" -Class "WMIMonitorID" -ComputerName $Computer -ErrorAction SilentlyContinue
    
        #Monitor Count
        $m = $Monitors | measure
        '"MonitorCount" :' + '"' + ($m.Count) + '",' | Add-Content $OutputPath
        #If there are no monitors don't create a string
        if ( ($m.Count) -ne 0)
        {
            #Takes each monitor object found and runs the following code:
            #Monitor String construction start
            $MonitorsString = '"Monitors" : ['
            #Start ForEach Monitor
            ForEach ($Monitor in $Monitors) {
              #Grabs respective data and converts it from ASCII encoding and removes any trailing ASCII null values
              If ([System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName) -ne $null) {
                $Mon_Model = ([System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName)).Replace("$([char]0x0000)","")
              } else {
                $Mon_Model = $null
              }
              $Mon_Serial_Number = ([System.Text.Encoding]::ASCII.GetString($Monitor.SerialNumberID)).Replace("$([char]0x0000)","")
              $Mon_Attached_Computer = ($Monitor.PSComputerName).Replace("$([char]0x0000)","")
              $Mon_Manufacturer = ([System.Text.Encoding]::ASCII.GetString($Monitor.ManufacturerName)).Replace("$([char]0x0000)","")
      
              #Filters out "non monitors". Place any of your own filters here. These two are all-in-one computers with built in displays. I don't need the info from these.
              If ($Mon_Model -like "*800 AIO*" -or $Mon_Model -like "*8300 AiO*") {Break}
      
              #Sets a friendly name based on the hash table above. If no entry found sets it to the original 3 character code
              $Mon_Manufacturer_Friendly = $ManufacturerHash.$Mon_Manufacturer
              If ($Mon_Manufacturer_Friendly -eq $null) {
                $Mon_Manufacturer_Friendly = $Mon_Manufacturer
              }
          
              $MonitorsString = $MonitorsString + '{"Monitor-Model" : "' + $Mon_Model + '", ' + '"Monitor-SN" : ' + '"' + $Mon_Serial_Number + '"},'
              #Monitor Detailed Info end
            }
            #Delete last "," close monitor string and write Monitor info to file
            $MonitorsString = $MonitorsString -replace ".$"
            $MonitorsString = $MonitorsString + "]" | Add-Content $OutputPath
        }
        else
        {
            '"Monitors" : ""' | Add-Content $OutputPath
        }
        #json file end
        "}" | Add-Content $OutputPath
        "]" | Add-Content $OutputPath

        #Clear the variables to kill a bug that used previous values if there was no objects found from wmi
        Clear-Variable -name computerHDD
    }
}

#Set variables and clear them
Set-Variable -name ArrComputers
Clear-Variable -name ArrComputers

#Check parameters passed from CLI and ask confirmation to run for the whole domain
if ($Computers -ne $null)
{
    $ArrComputers = $Computers
    Write-Host "Running for these computers :" $ArrComputers
    GetInfo $ArrComputers
}
elseif ($NotFound)
{
    Write-Host "Running for not found computers from previous execution of this script.`nScript will not stop running until all computers are found.`n There is a 2 hour sleep between the executions."
    #Break if there is no previous execution
    if(! (Test-Path($NotFoundPathRoot)) )
    {
        Write-Host "Wrong directory of execution`nThere are no records of previous execution of this script.`nPlease check the execution path of this script."
        break
    }
    $ArrComputers = Get-ChildItem $NotFoundPathRoot | Where-Object {$_.Extension -eq ".txt"} | % {$_.BaseName}
    GetInfo $ArrComputers
    #Write-Host "Sleeping.. Please terminate this script only when it is sleeping !`nSleep lasts for two hours."
    #Start-Sleep -s 7200
    break
    
}
else
{
    #Clear results from previous runs
    Get-ChildItem -Path $FolderRoot -Include *.* -File -Recurse | foreach { $_.Delete()}
    #Get all AD user computers
    $ArrComputers = Get-ADComputer -Filter * | where-object { $_.DistinguishedName -like "*OU=UserPCs*" -OR $_.DistinguishedName -like "*OU=Workstations*" -AND $_.DistinguishedName -notlike "*OU=Toasters*" } | select-object -expandproperty name
    GetInfo $ArrComputers
    break
}