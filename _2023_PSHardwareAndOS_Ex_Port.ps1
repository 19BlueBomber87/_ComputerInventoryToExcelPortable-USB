    #10.31.2023_MK_<@:D_99577
    #V.3 adds user profile list 
    $usbDriveLetter = Get-CimInstance -Class Win32_DiskDrive -Filter 'InterfaceType = "USB"' -KeyOnly |`
    Get-CimAssociatedInstance -Association Win32_DiskDriveToDiskPartition -KeyOnly |`
    Get-CimAssociatedInstance -Association Win32_LogicalDiskToPartition |`
    Where-Object -Property VolumeName -like *Yahoo*
    $usb = $usbDriveLetter.Name
#non-USB#
# Save-Module importexcel -Path C:\Users\admin.mkruse\Documents -Verbose *>&1
# Import-Module -Name C:\Users\admin.mkruse\Documents\importexcel -Verbose *>&1

    Import-Module -Name "$usb\ImportExcel" -Verbose *>&1
  ####Start###
    $Date = Get-Date -Format "MM-dd-yyyy HH:mm"
    $Bios = Get-WmiObject win32_bios -Computername $ENV:COMPUTERNAME
    $Hardware = Get-WmiObject Win32_computerSystem -Computername $ENV:COMPUTERNAME
    $Sysbuild = Get-WmiObject Win32_WmiSetting -Computername $ENV:COMPUTERNAME
    $OS = Get-WmiObject Win32_OperatingSystem -Computername $ENV:COMPUTERNAME
    $OSBuild = (Get-Item "HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue('ReleaseID')
    $Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $ENV:COMPUTERNAME | Where-Object {$_.IPEnabled}
    $driveSpace =  Get-WmiObject Win32_LogicalDisk -ComputerName $ENV:COMPUTERNAME -Filter "DeviceID='C:'" | Select-Object @{Name="HDSize";Expression={[math]::truncate($_.size/1GB)}}
    $cpu = Get-WmiObject Win32_Processor  -computername $ENV:COMPUTERNAME
    $memory = [math]::round($Hardware.TotalPhysicalMemory/1GB)
    $start_date=(Get-ComputerInfo).OsInstallDate
    $WMI_ChassisProps = @('ChassisTypes','Manufacturer','Model','SerialNumber')
    $dock = Read-Host "Enter dock Model " #I cant figure out how to use powershel to get the dock model number.  Its pretty easy to get it from the unit itself.  
    $dock_serial =  Read-Host "Enter dock serial number "
    $Domain = Read-Host "Enter Domain"
    $monitor_1_model = ((Get-WmiObject WmiMonitorID -Namespace root\wmi)[0]).InstanceName
    
    $monitor_2_model = ((Get-WmiObject WmiMonitorID -Namespace root\wmi)[1]).InstanceName
    $monitor_3_year = ((Get-WmiObject WmiMonitorID -Namespace root\wmi)[2]).YearOfManufacture
    $monitor_3_model = ((Get-WmiObject WmiMonitorID -Namespace root\wmi)[2]).InstanceName
    $IPAddress  = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.DefaultIPGateway -ne $null }).IPAddress | Select-Object -First 1
    $MACAddress  = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.DefaultIPGateway -ne $null } | Select-Object -ExpandProperty MACAddress
    $systemBios = $Bios.serialnumber
    $userPool = Get-ChildItem C:\Users | Sort-Object -Property LastWriteTime -Descending
    
    
    
    
    
    
 ###########EXPORT LOCATION#############
        # try{
    #     $location = ((Get-Volume -FileSystemLabel multiboot -ErrorAction Stop).DriveLetter + ":")
    #     }
    #     catch{
    #     $location = (get-location).path
    #     }
##########FILENAME        

#USB
$file_name = "$usb\"+"_PC_Inventory" + ".xlsx"
# #noUSB
#     $file_name = "_PC_Inventory" + ".xlsx"
    
    
    
    
######FUNCTIONS###############
    #1 GPU(s)

function GetGPUInfo {
    $GPUs = Get-WmiObject -Class Win32_VideoController
    foreach ($GPU in $GPUs) {
      $GPU | Select-Object -ExpandProperty Description
    }
  }
  ###################
  #2 Monitor
  function GetMonitorInfo {
    # Thanks to https://github.com/MaxAnderson95/Get-Monitor-Information
    $Monitors = Get-WmiObject -Namespace "root\WMI" -Class "WMIMonitorID"
    foreach ($Monitor in $Monitors) {
      ([System.Text.Encoding]::ASCII.GetString($Monitor.ManufacturerName)).Replace("$([char]0x0000)","")
      ([System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName)).Replace("$([char]0x0000)","")
      ([System.Text.Encoding]::ASCII.GetString($Monitor.SerialNumberID)).Replace("$([char]0x0000)","")
    }
  }
  

  $Monitor_1 = GetMonitorInfo | Select-Object -Index 1
  $Monitor_1_SN = GetMonitorInfo | Select-Object -Index 2
  $monitor_1_year = ((Get-WmiObject WmiMonitorID -Namespace root\wmi)[0]).YearOfManufacture
  $Monitor_2 = GetMonitorInfo | Select-Object -Index 4
  $Monitor_2_SN = GetMonitorInfo | Select-Object -Index 5
  $monitor_2_year = ((Get-WmiObject WmiMonitorID -Namespace root\wmi)[1]).YearOfManufacture
  $Monitor_3 = GetMonitorInfo | Select-Object -Index 7
  $Monitor_3_SN = GetMonitorInfo | Select-Object -Index 8
  $monitor_3_year = ((Get-WmiObject WmiMonitorID -Namespace root\wmi)[2]).YearOfManufacture
  
  ###########
  $GPU0 = GetGPUInfo | Select-Object -Index 0
  $GPU1 = GetGPUInfo | Select-Object -Index 1
  $Chassis = Get-CimInstance -ClassName Win32_SystemEnclosure -Namespace 'root\CIMV2' -Property ChassisTypes | Select-Object -ExpandProperty ChassisTypes
  
  ##LOGIC TO FIND DEVICE TYPE#############
  # https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-systemenclosure
  if ($Chassis -eq "1") {
    $Chassis = "Other"
  }
  if ($Chassis -eq "2") {
    $Chassis = "Unknown"
  }
  if ($Chassis -eq "3") {
    $Chassis = "Desktop"
  }
  if ($Chassis -eq "4") {
    $Chassis = "Low Profile Desktop"
  }
  if ($Chassis -eq "5") {
    $Chassis = "Pizza Box"
  }
  if ($Chassis -eq "6") {
    $Chassis = "Mini Tower"
  }
  if ($Chassis -eq "7") {
    $Chassis = "Tower"
  }
  if ($Chassis -eq "8") {
    $Chassis = "Portable"
  }
  if ($Chassis -eq "9") {
    $Chassis = "Laptop"
  }
  if ($Chassis -eq "10") {
    $Chassis = "Notebook"
  }
  if ($Chassis -eq "11") {
    $Chassis = "Hand Held"
  }
  if ($Chassis -eq "12") {
    $Chassis = "Docking Station"
  }
  if ($Chassis -eq "13") {
    $Chassis = "All in One"
  }
  if ($Chassis -eq "14") {
    $Chassis = "Sub Notebook"
  }
  if ($Chassis -eq "15") {
    $Chassis = "Space-Saving"
  }
  if ($Chassis -eq "16") {
    $Chassis = "Lunch Box"
  }
  if ($Chassis -eq "17") {
    $Chassis = "Main System Chassis"
  }
  if ($Chassis -eq "18") {
    $Chassis = "Expansion Chassis"
  }
  if ($Chassis -eq "19") {
    $Chassis = "SubChassis"
  }
  if ($Chassis -eq "20") {
    $Chassis = "Bus Expansion Chassis"
  }
  if ($Chassis -eq "21") {
    $Chassis = "Peripheral Chassis"
  }
  if ($Chassis -eq "22") {
    $Chassis = "Storage Chassis"
  }
  if ($Chassis -eq "23") {
    $Chassis = "Rack Mount Chassis"
  }
  if ($Chassis -eq "24") {
    $Chassis = "Sealed-Case PC"
  }
  $FloorPlan = ""
  Work_Space_View = ""
  #build inventory object
    $PC_object  = New-Object -Type PSObject
    $PC_object  | Add-Member -MemberType NoteProperty -Name CollectionDate -Value $Date
    $PC_object  | Add-Member -MemberType NoteProperty -Name IP_Address -Value $IPAddress
    $PC_object  | Add-Member -MemberType NoteProperty -Name Last_User -Value $userPool[0].name
    $PC_object  | Add-Member -MemberType NoteProperty -Name DeviceName -Value $ENV:COMPUTERNAME
    $PC_object  | Add-Member -MemberType NoteProperty -Name MAC_Address -Value $MACAddress
    $PC_object  | Add-Member -MemberType NoteProperty -Name Type -Value $Chassis
    $PC_object  | Add-Member -MemberType NoteProperty -Name Model -Value $Hardware.Model
    $PC_object  | Add-Member -MemberType NoteProperty -Name FloorPlan -Value $FloorPlan
    $PC_object  | Add-Member -MemberType NoteProperty -Name Work_Space_View -Value $Work_Space_View
    $PC_object  | Add-Member -MemberType NoteProperty -Name Serial_Number -Value $systemBios
    $PC_object  | Add-Member -MemberType NoteProperty -Name Processor_Type -Value $cpu.Name
    $PC_object  | Add-Member -MemberType NoteProperty -Name HDSize_GB -Value $driveSpace.HDSize
    $PC_object  | Add-Member -MemberType NoteProperty -Name Total_Memory_GB -Value $memory
    $PC_object  | Add-Member -MemberType NoteProperty -Name Operating_System -Value $OS.Caption 
    $PC_object  | Add-Member -MemberType NoteProperty -Name OS_Build -Value $OSBuild
    $PC_object  | Add-Member -MemberType NoteProperty -Name OS_Install_Date -Value $start_date
    $PC_object  | Add-Member -MemberType NoteProperty -Name Docking_Station_Model -Value $dock #update this after a good test. 
    $PC_object  | Add-Member -MemberType NoteProperty -Name Docking_Station_serial -Value $dock_serial #update this after a good test. 
    $PC_object  | Add-Member -MemberType NoteProperty -Name Monitor_1_Year -Value $monitor_1_year
    $PC_object  | Add-Member -MemberType NoteProperty -Name Monitor_1 -Value $monitor_1
    $PC_object  | Add-Member -MemberType NoteProperty -Name Monitor_1_SN -Value $Monitor_1_sn
    $PC_object  | Add-Member -MemberType NoteProperty -Name Monitor_2_Year -Value $monitor_2_year
    $PC_object  | Add-Member -MemberType NoteProperty -Name Monitor_2 -Value $monitor_2
    $PC_object  | Add-Member -MemberType NoteProperty -Name Monitor_2_SN -Value $Monitor_2_SN
    $PC_object  | Add-Member -MemberType NoteProperty -Name Monitor_3_year -Value $monitor_3_year
    $PC_object  | Add-Member -MemberType NoteProperty -Name Monitor_3_model -Value $monitor_3
    $PC_object  | Add-Member -MemberType NoteProperty -Name Monitor_3_SN -Value $Monitor_3_SN
    (Get-ChildItem c:\users) | ForEach-Object {$UserList += ("[" + $_.name + "]")}
    $PC_object  | Add-Member -MemberType NoteProperty -Name UserList -Value $userlist #Version 3 Update


######EXPORT OBJECT TO EXCEL
    try{
    $PC_object  | Export-Excel $file_name -Append -TableName $Domain -AutoSize -WorksheetName $Domain 
    }
    catch{
    Write-Verbose -Message "Error :(" -Verbose *>&1

    }$pc

     $PC_object  | fl *
<#notes
(gwmi win32_physicalmemory)[0] | fl banklabel, @{Name="GB";Expression={$_.Capacity/1GB}}
(gwmi win32_physicalmemory)[1] | fl banklabel, @{Name="GB";Expression={$_.Capacity/1GB}}
gwmi win32_computersystem| fl @{Name="GB";Expression={$_.totalphysicalmemory /1000MB}}
$PC_object | Out-GridView
#>
$dock


#https://www.hofferle.com/retrieve-monitor-serial-numbers-with-powershell/

#Combine
#$yahoo = Get-ChildItem | ? -Property name -like *.csv
#$yahoo | ForEach-Object -Process  {Import-Csv $psitem.fullname | Export-Excel .\Master\yahoo.xlsx -Append}
