Function Get-ITReport{
    <#
    .SYNOPSIS
    Pulls a report for various IT related things.

    .DESCRIPTION
    This module collects various information, specified by the argument, into CSV reports. Some of the information that is collected by this script include: QFE, Shutdown Log, Uptime, and Installed Services.

    Created By James Arnett

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
    Function GetReportData{   
        Function GetServerListPath {
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
        $ReportData = New-Object PSObject
        $ServerListPath = GetServerListPath
        $ReportPath = GetReportPath
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "ServerListPath" -Value $ServerListPath
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "ReportPath" -Value $ReportPath
        Return $ReportData
    }
    Try{
        $ReportData = GetReportData -ErrorAction Stop
        $ServerListPath = $ReportData.ServerListPath
        $ReportPath = $ReportData.ReportPath
        }Catch{
            Write-Error "Error attempting to gather report data!"
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
    <#
    .SYNOPSIS
    Pulls a report for hardware attached to the server/ computer specified.

    .DESCRIPTION
    This module collects information about the server/ computer hardware into a CSV report. Some of the information that is collected by this module include: CPU, Memory, and Hard Drive.

    Created By James Arnett.
        
    REQUIREMENTS
    - Powershell window must be ran as an Elevated User with access to administrative rights.    
        
    .EXAMPLE

    This will run all the above listed reports on the given list of servers.
    C:\PS> Get-HardwareReport

    This parameter allows you to specify a specific server instead of the module prompting you for a list. This can be used with any of the parameters above.
    C:\PS> Get-HardwareReport -Server <Server Name>

    .LINK

    http://bloggintechie.blogspot.com/
    https://chromebookparadise.wordpress.com/
    #>
    [cmdletbinding()]
    Param(
        [string]$Server,
        [string]$ServerListPath,
        [string]$ReportPath
    )
    Function GetReportData{   
        Function GetServerListPath {
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
        $ReportData = New-Object PSObject
        if(!$server -and !$ServerListPath -and !$ReportPath){
            $ServerListPath = GetServerListPath
            $serverList = Get-Content $ServerListPath
            $ReportPath = GetReportPath
            $file_now = Get-Date -format MM-dd-yy.hhmmtt
            $now = Get-Date
            }elseif($ServerListPath -and $ReportPath){
                $serverList = Get-Content $ServerListPath
                $file_now = Get-Date -format MM-dd-yy.hhmmtt
                $now = Get-Date
                }else{
                    $serverList = $Server
                    $ReportPath = GetReportPath
                    $file_now = Get-Date -format MM-dd-yy.hhmmtt
                    $now = Get-Date
                    }
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "serverList" -Value $serverList
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "ReportPath" -Value $ReportPath
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "file_now" -Value $file_now
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "now" -Value $now
        Return $ReportData
    }
    Try{
        $ReportData = GetReportData -ErrorAction Stop
        $serverList = $ReportData.serverList
        $ReportPath = $ReportData.ReportPath
        [string]$file_now = $ReportData.file_now
        $now = $ReportData.now
        }Catch{
            Write-Error "Error attempting to gather report data!"
            }
    write-host `n`rStarting Hardware Report Generation.... -ForegroundColor Yellow `n`r
    Try{
        $HadwareReportsDir = "Hardware-Reports"
        if(!(Test-Path -Path $ReportPath\$HadwareReportsDir )){
            New-Item -path $ReportPath -name $HadwareReportsDir -ItemType directory -ErrorAction Stop | out-null
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
        Write-Error "An error has occured attempting to pull the hardware report. Error: $ReportError"
    }
    Write-Host `n`rHardware Report generation completed... -ForegroundColor Yellow `n`r
    Try{
        $infoColl | Export-Csv "$ReportPath\$HadwareReportsDir\$file_now.csv" -NoTypeInformation -ErrorAction Stop -ErrorVariable ReportErr
        Write-Host "Report has been created at $ReportPath\$HadwareReportsDir\$file_now.csv"
        }Catch{
            Write-Error "An error has occured creating the report. Error: $ReportErr"
        }
}
Function Get-ShutdownLogReport{
    <#
    .SYNOPSIS
    Pulls a report for Shutdown logs for the server/ computer specified.

    .DESCRIPTION
    This module collects information about the server/ computer Shutdown logs into a CSV report. This module will prompt you for the amount of events that you would like to go back in the logs. 
    *** Disclaimer ***
    The more events you go back the longer this will take to run.

    Created By James Arnett.
        
    REQUIREMENTS
    - Powershell window must be ran as an Elevated User with access to administrative rights.    
        
    .EXAMPLE

    This will run all the above listed reports on the given list of servers.
    C:\PS> Get-ShutdownLogReport

    This parameter allows you to specify a specific server instead of the module prompting you for a list. This can be used with any of the parameters above.
    C:\PS> Get-ShutdownLogReport -Server <Server Name>

    .LINK

    http://bloggintechie.blogspot.com/
    https://chromebookparadise.wordpress.com/
    #>
    [cmdletbinding()]
    Param(
        [string]$Server,
        [string]$ServerListPath,
        [string]$ReportPath
    )
    Function GetReportData{   
        Function GetServerListPath {
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
        $ReportData = New-Object PSObject
        if(!$server -and !$ServerListPath -and !$ReportPath){
            $ServerListPath = GetServerListPath
            $serverList = Get-Content $ServerListPath
            $ReportPath = GetReportPath
            $file_now = Get-Date -format MM-dd-yy.hhmmtt
            $now = Get-Date
            }elseif($ServerListPath -and $ReportPath){
                $serverList = Get-Content $ServerListPath
                $file_now = Get-Date -format MM-dd-yy.hhmmtt
                $now = Get-Date
                }else{
                    $serverList = $Server
                    $ReportPath = GetReportPath
                    $file_now = Get-Date -format MM-dd-yy.hhmmtt
                    $now = Get-Date
                    }
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "serverList" -Value $serverList
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "ReportPath" -Value $ReportPath
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "file_now" -Value $file_now
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "now" -Value $now
        Return $ReportData
    }
    Try{
        $ReportData = GetReportData -ErrorAction Stop
        $serverList = $ReportData.serverList
        $ReportPath = $ReportData.ReportPath
        [string]$file_now = $ReportData.file_now
        $now = $ReportData.now
        }Catch{
            Write-Error "Error attempting to gather report data!"
            }
    write-host `n`rStarting Shutdown Log Report Generation....`n`r -ForegroundColor Yellow
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $EventAmount = [Microsoft.VisualBasic.Interaction]::InputBox("From Newest to Oldest, how many Shutdown events would you like to go back?", "Shutdown Events", "10")
    $EventAmount =  $EventAmount.Trim();
    $ShutdownLogDir = "ShutdownLog-Reports"
    if(!(Test-Path -Path $ReportPath\$ShutdownLogDirr )){
        New-Item -path $ReportPath -name $ShutdownLogDir -ItemType directory -ErrorAction Stop | out-null
    }
    $infoColl = @()
    Foreach($Server in $ServerList){
        Try{
            $ScriptBlock = [scriptblock]::Create("Get-EventLog -ComputerName $Server -LogName System -Newest $EventAmount -InstanceId 2147484722 -ErrorAction SilentlyContinue | select MachineName,Message,TimeWritten,UserName")
            $EventlogJob = Start-Job -ScriptBlock $ScriptBlock -ErrorAction Stop
            Wait-Job $EventlogJob -Timeout 60
            Stop-Job $EventlogJob
            $Eventlog = Receive-Job $EventlogJob -AutoRemoveJob -ErrorVariable EventlogJobErr
            if($EventlogJobErr){
                $infoObject = New-Object PSObject
                Add-Member -inputObject $infoObject -memberType NoteProperty -name "ServerName" -value $Server
                Add-Member -inputObject $infoObject -memberType NoteProperty -name "Error" -value $EventlogJobErr
                Write-Error "An error has occured for $Server Error: $EventlogJobErr" -ErrorAction Stop
                $infoObject
                $infoColl += $infoObject
                }else{
	                Foreach ($Event in $Eventlog){
                        $infoObject = New-Object PSObject
                        Add-Member -inputObject $infoObject -memberType NoteProperty -name "ServerName" -value $Server
                        Add-Member -inputObject $infoObject -memberType NoteProperty -name "MachineName" -value $Event.MachineName
				        Add-Member -inputObject $infoObject -memberType NoteProperty -name "Message" -value $Event.Message
				        Add-Member -inputObject $infoObject -memberType NoteProperty -name "TimeWritten" -value $Event.TimeWritten
				        Add-Member -inputObject $infoObject -memberType NoteProperty -name "UserName" -value $Event.UserName
				        Add-Member -inputObject $infoObject -memberType NoteProperty -name "PSComputerName" -value $Event.PSComputerName
				        Add-Member -inputObject $infoObject -memberType NoteProperty -name "RunspaceId" -value $Event.RunspaceId
				        Add-Member -inputObject $infoObject -memberType NoteProperty -name "PSShowComputerName" -value $Event.PSShowComputerName
                        $infoObject
                        $infoColl += $infoObject
                    }
                }
            Write-Host Shutdown Log was successfully pulled for $Server`n`r
            }Catch{
                Write-Error "An error has occurred attempting to pull logs for $Server. Please try again! Error: $EventlogJobErr"
                }
    }
    Try{
        $infoColl | Export-Csv "$ReportPath\$ShutdownLogDir\$file_now.csv" -NoTypeInformation -ErrorAction Stop -ErrorVariable ReportErr
        Write-Host `n`r "Report can be found here: $ReportPath\$ShutdownLogDir\$file_now.csv" -ForegroundColor Green `n`r
        }Catch{
            Write-Error "An error has occured generating the report. Error: $ReportErr"
            }
    Write-Host `n`rShutdown Log Report generation has completed... -ForegroundColor Yellow `n`r
}
Function Get-UpTimeReport{
    <#
    .SYNOPSIS
    Pulls a report for Uptime for the server/ computer specified.

    .DESCRIPTION
    This module collects information about the server/ computer Uptime into a CSV report.

    Created By James Arnett.
        
    REQUIREMENTS
    - Powershell window must be ran as an Elevated User with access to administrative rights.    
        
    .EXAMPLE

    This will run all the above listed reports on the given list of servers.
    C:\PS> Get-UpTimeReport

    This parameter allows you to specify a specific server instead of the module prompting you for a list. This can be used with any of the parameters above.
    C:\PS> Get-UpTimeReport -Server <Server Name>

    .LINK

    http://bloggintechie.blogspot.com/
    https://chromebookparadise.wordpress.com/
    #>
    [cmdletbinding()]
    Param(
        [string]$Server,
        [string]$ServerListPath,
        [string]$ReportPath
    )
    Function GetReportData{   
        Function GetServerListPath {
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
        $ReportData = New-Object PSObject
        if(!$server -and !$ServerListPath -and !$ReportPath){
            $ServerListPath = GetServerListPath
            $serverList = Get-Content $ServerListPath
            $ReportPath = GetReportPath
            $file_now = Get-Date -format MM-dd-yy.hhmmtt
            $now = Get-Date
            }elseif($ServerListPath -and $ReportPath){
                $serverList = Get-Content $ServerListPath
                $file_now = Get-Date -format MM-dd-yy.hhmmtt
                $now = Get-Date
                }else{
                    $serverList = $Server
                    $ReportPath = GetReportPath
                    $file_now = Get-Date -format MM-dd-yy.hhmmtt
                    $now = Get-Date
                    }
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "serverList" -Value $serverList
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "ReportPath" -Value $ReportPath
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "file_now" -Value $file_now
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "now" -Value $now
        Return $ReportData
    }
    Try{
        $ReportData = GetReportData -ErrorAction Stop
        $serverList = $ReportData.serverList
        $ReportPath = $ReportData.ReportPath
        [string]$file_now = $ReportData.file_now
        $now = $ReportData.now
        }Catch{
            Write-Error "Error attempting to gather report data!"
            }
    write-host `n`rStarting UpTime Report Generation.... -ForegroundColor Yellow `n`r
    $ReportsDir = "Uptime-Reports"
    if(!(Test-Path -Path $ReportPath\$ReportsDir )){
        New-Item -path $ReportPath -name $ReportsDir -ItemType directory -ErrorAction Stop | out-null
        }
    $infoColl = @()
    ForEach($server in $serverlist){
        Try{
            $infoObject = New-Object PSObject
            $Uptime = Get-CimInstance Win32_OperatingSystem -ComputerName $server -ErrorAction Stop -ErrorVariable UptimeError
            Add-Member -inputObject $infoObject -memberType NoteProperty -name "ServerName" -Value $Server
            Add-Member -inputObject $infoObject -memberType NoteProperty -name "LastBoot" -Value $Uptime.LastBootUpTime
            $infoObject
            $infoColl += $infoObject
            }Catch{
                Write-Error "An error has occurred attempting to gather uptime data for $Server. Error: $UptimeError"
                }
        }
    Try{
        $infoColl | Export-Csv "$ReportPath\$ReportsDir\$file_now.csv" -NoTypeInformation -ErrorAction Stop -ErrorVariable ReportErr
        Write-Host `n`r "Report can be found here: $ReportPath\$ReportsDir\$file_now.csv" -ForegroundColor Green `n`r
        }Catch{
            Write-Error "An error has occured generating the report. Error: $ReportErr"
            }
    Write-Host `n`rUpTime Report generation completed... -ForegroundColor Yellow `n`r
}
Function Get-ServiceReport{
    <#
    .SYNOPSIS
    Pulls a report for Installed Services for the server/ computer specified.

    .DESCRIPTION
    This module collects information about the server/ computer Installed Services into a CSV report.

    Created By James Arnett.
        
    REQUIREMENTS
    - Powershell window must be ran as an Elevated User with access to administrative rights.    
        
    .EXAMPLE

    This will run all the above listed reports on the given list of servers.
    C:\PS> GGet-ServiceReport

    This parameter allows you to specify a specific server instead of the module prompting you for a list. This can be used with any of the parameters above.
    C:\PS> Get-ServiceReport -Server <Server Name>

    .LINK

    http://bloggintechie.blogspot.com/
    https://chromebookparadise.wordpress.com/
    #>
    [cmdletbinding()]
    Param(
        [string]$Server,
        [string]$ServerListPath,
        [string]$ReportPath
    )
    Function GetReportData{   
        Function GetServerListPath {
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
        $ReportData = New-Object PSObject
        if(!$server -and !$ServerListPath -and !$ReportPath){
            $ServerListPath = GetServerListPath
            $serverList = Get-Content $ServerListPath
            $ReportPath = GetReportPath
            $file_now = Get-Date -format MM-dd-yy.hhmmtt
            $now = Get-Date
            }elseif($ServerListPath -and $ReportPath){
                $serverList = Get-Content $ServerListPath
                $file_now = Get-Date -format MM-dd-yy.hhmmtt
                $now = Get-Date
                }else{
                    $serverList = $Server
                    $ReportPath = GetReportPath
                    $file_now = Get-Date -format MM-dd-yy.hhmmtt
                    $now = Get-Date
                    }
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "serverList" -Value $serverList
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "ReportPath" -Value $ReportPath
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "file_now" -Value $file_now
        Add-Member -inputObject $ReportData -memberType NoteProperty -name "now" -Value $now
        Return $ReportData
    }
    Try{
        $ReportData = GetReportData -ErrorAction Stop
        $serverList = $ReportData.serverList
        $ReportPath = $ReportData.ReportPath
        [string]$file_now = $ReportData.file_now
        $now = $ReportData.now
        }Catch{
            Write-Error "Error attempting to gather report data!"
            }
    write-host `n`rStarting Service Report Generation.... -ForegroundColor Yellow `n`r
    $ReportsDir = "Service-Reports"
    if(!(Test-Path -Path $ReportPath\$ReportsDir )){
        New-Item -path $ReportPath -name $ReportsDir -ItemType directory -ErrorAction Stop | out-null
        }
    $infoColl = @()
    foreach ($server in $serverlist){
        Try{
            $Services = gwmi win32_service -ComputerName $server -ErrorAction Stop -ErrorVariable ServicesError
            $infoColl += $Services
            }Catch{
                Write-Error "An error has occurred attempting to gather Services data for $Server. Error: $ServicesError"
                }
        }
    Try{
        $infoColl | Export-Csv "$ReportPath\$ReportsDir\$file_now.csv" -NoTypeInformation -ErrorAction Stop -ErrorVariable ReportErr
        Write-Host `n`r "Report can be found here: $ReportPath\$ReportsDir\$file_now.csv" -ForegroundColor Green `n`r
        }Catch{
            Write-Error "An error has occured generating the report. Error: $ReportErr"
            }
    write-host `n`rService Report Generation Completed... -ForegroundColor Yellow `n`r
}