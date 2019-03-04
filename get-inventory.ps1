<#
.SYNOPSIS
  Name: Get-Inventory.ps1
  The purpose of this script is to create a simple inventory.

.DESCRIPTION
  This is a simple script to retrieve all computer objects in Active Directory and then connect
  to each one and gather basic hardware information using WMI. The information includes Manufacturer,
  Model, CPU, RAM, Disks and Operating System.

.EXAMPLE
  Get-Inventory.ps1
  Find the output under C:\Scripts_Output
#>

$Inventory = New-Object System.Collections.ArrayList

$AllComputers = Get-ADComputer -Filter * -Properties Name
$AllComputersNames = $AllComputers.Name

Foreach ($ComputerName in $AllComputersNames) {

   $Connection = Test-Connection $ComputerName -Count 1 -Quiet

   $ComputerInfo = New-Object System.Object

   $ComputerOS = Get-ADComputer $ComputerName -Properties OperatingSystem,OperatingSystemServicePack

   $ComputerInfoOperatingSystem = $ComputerOS.OperatingSystem
   $ComputerInfoOperatingSystemServicePack = $ComputerOS.OperatingSystemServicePack

   $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Name" -Value "$ComputerName" -Force
   $ComputerInfo | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $ComputerInfoOperatingSystem
   $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ServicePack" -Value $ComputerInfoOperatingSystemServicePack

   if ($Connection -eq "True"){

      $ComputerHW = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName | select Manufacturer,Model,NumberOfProcessors,@{Expression={$_.TotalPhysicalMemory / 1GB};Label="TotalPhysicalMemoryGB"}
      $ComputerCPU = Get-WmiObject win32_processor -ComputerName $ComputerName | select DeviceID,Name,Manufacturer,NumberOfCores,NumberOfLogicalProcessors
      $ComputerDisks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $ComputerName | select DeviceID,VolumeName,@{Expression={$_.Size / 1GB};Label="SizeGB"}

      $ComputerInfoManufacturer = $ComputerHW.Manufacturer
      $ComputerInfoModel = $ComputerHW.Model
      $ComputerInfoNumberOfProcessors = $ComputerHW.NumberOfProcessors
      $ComputerInfoProcessorID = $ComputerCPU.DeviceID
      $ComputerInfoProcessorManufacturer = $ComputerCPU.Manufacturer
      $ComputerInfoProcessorName = $ComputerCPU.Name
      $ComputerInfoNumberOfCores = $ComputerCPU.NumberOfCores
      $ComputerInfoNumberOfLogicalProcessors = $ComputerCPU.NumberOfLogicalProcessors
      $ComputerInfoRAM = $ComputerHW.TotalPhysicalMemoryGB
      $ComputerInfoDiskDrive = $ComputerDisks.DeviceID
      $ComputerInfoDriveName = $ComputerDisks.VolumeName
      $ComputerInfoSize = $ComputerDisks.SizeGB

      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Manufacturer" -Value "$ComputerInfoManufacturer" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Model" -Value "$ComputerInfoModel" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "NumberOfProcessors" -Value "$ComputerInfoNumberOfProcessors" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ProcessorID" -Value "$ComputerInfoProcessorID" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ProcessorManufacturer" -Value "$ComputerInfoProcessorManufacturer" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "ProcessorName" -Value "$ComputerInfoProcessorName" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "NumberOfCores" -Value "$ComputerInfoNumberOfCores" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "NumberOfLogicalProcessors" -Value "$ComputerInfoNumberOfLogicalProcessors" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "RAM" -Value "$ComputerInfoRAM" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "DiskDrive" -Value "$ComputerInfoDiskDrive" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "DriveName" -Value "$ComputerInfoDriveName" -Force
      $ComputerInfo | Add-Member -MemberType NoteProperty -Name "Size" -Value "$ComputerInfoSize"-Force
   }

$Inventory.Add($ComputerInfo) | Out-Null

$ComputerHW = ""
$ComputerCPU = ""
$ComputerDisks = ""
}

$Inventory | Export-Csv C:\Scripts_Output\Inventory.csv