Function Get-ITReport{
    <#
    .SYNOPSIS
    Pulls a report for various IT related things.

    .DESCRIPTION
    This module collects various information, specified by the argument, into CSV reports. Some of the information that is collected by this script include: QFE, Shutdown Log, Uptime, and Installed Services.

    Created By James A. Arnett

    UpTime: 
        This will generate a report for the last time the server was rebooted. 
    QFE: 
        This will generate a report of all patches and hotfixes installed based on a given number of days.  
    ShutdownLog: 
        This will generate a report of users that have initiated a shutdown on a server based on a given number of days.
    Service:
        This will generate a report of services that installed on the server.  
    Server:
        This parameter must be used in conjunction with one or more of the above listed parameters.
        This is used to specify a specific server. For usage syntax see Examples section.
        
    REQUIREMENTS
    - Powershell window must be ran as an Elevated User with access to administrative rights.    
        
    .EXAMPLE

    This will run all the above listed reports on the given list of servers.
    C:\PS> Get-ITReport    

    This will generate a report of all patches and hotfixes installed based on a given number of days.
    C:\PS> Get-ITReport -QFE        

    This will generate a report of users that have initiated a shutdown on a server based on a given number of days.
    C:\PS> Get-ITReport -ShutdownLog

    This will generate a report for the last time the server was rebooted.
    C:\PS> Get-ITReport -UpTime

    This will generate a report of services that installed on the server.
    C:\PS> Get-ITReport -Service

    This parameter allows you to specify a specific server instead of the module prompting you for a list. This can be used with any of the parameters above.
    C:\PS> Get-ITReport -Server <Server Name>

    .LINK

    http://bloggintechie.blogspot.com/
    https://chromebookparadise.wordpress.com/
    #>
    [cmdletbinding()]
    Param(
        [switch]$Hardware,
        [switch]$UpTime,
        [switch]$QFE,
        [switch]$ShutdownLog,
        [switch]$Service,
        [string]$Server,
        [switch]$QFEDaily
    )
    Function GetServerList{   
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "Text Files (*.txt)| *.txt"
        $OpenFileDialog.Title = "Please Specify Serverlist File:"
        $Show = $OpenFileDialog.ShowDialog()
        If ($Show -eq "OK"){
            Return $OpenFileDialog.filename
            }Else{
                Write-Error "Operation cancelled by user." -ErrorAction Stop
                }

        if(!$server){
            $ServerListPath = GetServerList
            $serverList = Get-Content $ServerListPath
            $ReportPath = GetReportPath
            $file_now = Get-Date -format MM-dd-yy.hhmmtt
            $now = Get-Date
            }
            else{
                $serverList = $Server
                $ReportPath = GetReportPath
                $file_now = Get-Date -format MM-dd-yy.hhmmtt
                $now = Get-Date
                }
        
    }
    Function GetReportPath{
        param([string]$Description="Select Folder for Report",[string]$RootFolder="Desktop")
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog()
        If ($Show -eq "OK"){
            Return $objForm.SelectedPath
            }Else{
                Write-Error "Operation cancelled by user." -ErrorAction Stop
                }
    }
    if($Hardware){
        Get-HardwareReport
        }
    if($UpTime){
        Get-UpTimeReport
        }
    if($QFE){
        Get-QFEReport
            }
    if($ShutdownLog){
        Get-ShutdownLogReport
        }
    if($Service){
        Get-ServiceReport
        }
    if($QFEDaily){
        Get-DailyQFEReport
        }
    if((!$Service)-and(!$ShutdownLog)-and(!$QFE)-and(!$UpTime)-and(!$Hardware)){
        Get-UpTimeReport
        Get-QFEReport
        Get-ShutdownLogReport
        Get-ServiceReport
        Get-HardwareReport
        }
    
}
Function Get-HardwareReport{
    write-host `n`rStarting Hardware Report Generation.... -ForegroundColor Yellow `n`r
    Try{
        $HadwareReportsDir = "Hardware-Reports"
        if(!(Test-Path -Path $ReportPath\$HadwareReportsDir )){
            New-Item -path $ReportPath -name $HadwareReportsDir -ItemType directory -ErrorAction Stop | out-null
        }
        $TRD = "$file_now"
        if(!(Test-Path -Path $ReportPath\$HadwareReportsDir\$TRD)){
            New-Item -path $ReportPath\$HadwareReportsDir -name $TRD -ItemType directory -ErrorAction Stop | out-null
        }
        $infoColl = @()
        Foreach ($s in $ServerList){
	        $CPUInfo = Get-WmiObject Win32_Processor -ComputerName $s #Get CPU Information
	        $OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $s #Get OS Information
	        #Get Memory Information. The data will be shown in a table as MB, rounded to the nearest second decimal.
	        $OSTotalVirtualMemory = [math]::round($OSInfo.TotalVirtualMemorySize / 1MB, 2)
	        $OSTotalVisibleMemory = [math]::round(($OSInfo.TotalVisibleMemorySize / 1MB), 2)
	        $PhysicalMemory = Get-WmiObject CIM_PhysicalMemory -ComputerName $s | Measure-Object -Property capacity -Sum | % { [Math]::Round(($_.sum / 1GB), 2) }
            $LogicalDisks = gwmi win32_logicaldisk -ComputerName $s
            $i = 1
            $infoObject = New-Object PSObject
            Add-Member -inputObject $infoObject -memberType NoteProperty -name "ServerName" -value $s
	        Foreach ($CPU in $CPUInfo){		    
		        Add-Member -inputObject $infoObject -memberType NoteProperty -name "Processor" -value $CPU.Name
		        Add-Member -inputObject $infoObject -memberType NoteProperty -name "Model" -value $CPU.Description
		        Add-Member -inputObject $infoObject -memberType NoteProperty -name "Manufacturer" -value $CPU.Manufacturer
		        Add-Member -inputObject $infoObject -memberType NoteProperty -name "PhysicalCores" -value $CPU.NumberOfCores
		        Add-Member -inputObject $infoObject -memberType NoteProperty -name "CPU_L2CacheSize" -value $CPU.L2CacheSize
		        Add-Member -inputObject $infoObject -memberType NoteProperty -name "CPU_L3CacheSize" -value $CPU.L3CacheSize
		        Add-Member -inputObject $infoObject -memberType NoteProperty -name "Sockets" -value $CPU.SocketDesignation
		        Add-Member -inputObject $infoObject -memberType NoteProperty -name "LogicalCores" -value $CPU.NumberOfLogicalProcessors
            }
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "OS_Name" -value $OSInfo.Caption
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "OS_Version" -value $OSInfo.Version
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "TotalPhysical_Memory_GB" -value $PhysicalMemory
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "TotalVirtual_Memory_MB" -value $OSTotalVirtualMemory
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "TotalVisable_Memory_MB" -value $OSTotalVisibleMemory
            foreach($LogicalDisk in $LogicalDisks){
                $LDFreeSpace = [math]::round($LogicalDisk.FreeSpace / 1GB)
                $LDSize = [math]::round($LogicalDisk.Size / 1GB)
                Add-Member -inputObject $infoObject -memberType NoteProperty -name "Disk $i Letter" -Value $LogicalDisk.DeviceID
                Add-Member -inputObject $infoObject -memberType NoteProperty -name "Disk $i Freespace (GB)" -Value $LDFreeSpace
                Add-Member -inputObject $infoObject -memberType NoteProperty -name "Disk $i Size (GB)" -Value $LDSize
                Add-Member -inputObject $infoObject -memberType NoteProperty -name "Disk $i Provider Name"$LogicalDisk.ProviderName
                $infoColl += $infoObject
                $i++
                }
		    $infoObject
		    $infoColl += $infoObject
	        }
    }Catch{
        Write-Error An error has occured attempting to pull the hardware report. Error: $ReportError
    }
    Write-Host `n`rHardware Report generation completed... -ForegroundColor Yellow `n`r
    Try{
        $infoColl | Export-Csv "$ReportPath\$HadwareReportsDir\$TRD.csv" -NoTypeInformation -ErrorAction Stop -ErrorVariable CSVERROR
        Write-Host "Report has been created at $ReportPath\$HadwareReportsDir\$TRD.csv"
        }Catch{
            Write-Error "An error has occured creating the report. Error: $CSVERROR"
        }
}