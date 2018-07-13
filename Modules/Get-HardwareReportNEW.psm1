Function Get-HardwareReport{
    write-host `n`rStarting Hardware Report Generation.... -ForegroundColor Yellow `n`r
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
    $ServerListPath = GetServerList
    $serverList = Get-Content $ServerListPath
    $ReportPath = GetReportPath
    $file_now = Get-Date -format MM-dd-yy.hhmmtt
    $now = Get-Date
    $HRD = "Hardware-Reports"
    New-Item -path $ReportPath -name $HRD -ItemType directory -ErrorAction SilentlyContinue | out-null
    $TRD = "$file_now"
    New-Item -path $ReportPath\$HRD -name $TRD -ItemType directory -ErrorAction SilentlyContinue | out-null
    $infoColl = @()
    Foreach ($s in $servers){
	    $CPUInfo = Get-WmiObject Win32_Processor -ComputerName $s #Get CPU Information
	    $OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $s #Get OS Information
	    #Get Memory Information. The data will be shown in a table as MB, rounded to the nearest second decimal.
	    $OSTotalVirtualMemory = [math]::round($OSInfo.TotalVirtualMemorySize / 1MB, 2)
	    $OSTotalVisibleMemory = [math]::round(($OSInfo.TotalVisibleMemorySize / 1MB), 2)
	    $PhysicalMemory = Get-WmiObject CIM_PhysicalMemory -ComputerName $s | Measure-Object -Property capacity -Sum | % { [Math]::Round(($_.sum / 1GB), 2) }
        $LogicalDisks = gwmi win32_logicaldisk -ComputerName $s
        $LDFreeSpace = [math]::round($LogicalDisks.FreeSpace / 1GB);"GB"
        $LDSize = [math]::round($LogicalDisks.Size / 1GB);"GB"
        $i = 1
	    Foreach ($CPU in $CPUInfo)
	    {
		    $infoObject = New-Object PSObject
		    #The following add data to the infoObjects.	
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "ServerName" -value $CPU.SystemName
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "Processor" -value $CPU.Name
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "Model" -value $CPU.Description
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "Manufacturer" -value $CPU.Manufacturer
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "PhysicalCores" -value $CPU.NumberOfCores
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "CPU_L2CacheSize" -value $CPU.L2CacheSize
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "CPU_L3CacheSize" -value $CPU.L3CacheSize
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "Sockets" -value $CPU.SocketDesignation
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "LogicalCores" -value $CPU.NumberOfLogicalProcessors
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "OS_Name" -value $OSInfo.Caption
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "OS_Version" -value $OSInfo.Version
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "TotalPhysical_Memory_GB" -value $PhysicalMemory
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "TotalVirtual_Memory_MB" -value $OSTotalVirtualMemory
		    Add-Member -inputObject $infoObject -memberType NoteProperty -name "TotalVisable_Memory_MB" -value $OSTotalVisibleMemory
            Add-Member -inputObject $infoObject -memberType NoteProperty -name "LD Letter" -Value $LogicalDisks.DeviceID
            Add-Member -inputObject $infoObject -memberType NoteProperty -name "LD Freespace" -Value $LDFreeSpace
            Add-Member -inputObject $infoObject -memberType NoteProperty -name "LD Size" -Value $LDSize
            Add-Member -inputObject $infoObject -memberType NoteProperty -name "LD Provider Name"$LogicalDisks.ProviderName
		    $infoObject #Output to the screen for a visual feedback.
		    $infoColl += $infoObject
            $i++
	    }
    }
    Write-Host `n`rHardware Report generation completed... -ForegroundColor Yellow `n`r
}