Function Get-ITReport{
    <#
    .SYNOPSIS
    Pulls Reports for Various things.

    .DESCRIPTION
    This module collects various information, specified by the argument, into CSV reports.

    v1.1 Changes:
    Added ability to specify one server.

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
    This is the basic example of the syntax
    C:\PS> Get-ITReport
    This will run all the above listed reports on the given list of servers.

    C:\PS> Get-ITReport -QFE
        This will generate a report of all patches and hotfixes installed based on a given number of days.

    C:\PS> Get-ITReport -ShutdownLog
        This will generate a report of users that have initiated a shutdown on a server based on a given number of days.

    C:\PS> Get-ITReport -UpTime
        This will generate a report for the last time the server was rebooted.

    C:\PS> Get-ITReport -Service
        This will generate a report of services that installed on the server.

    C:\PS> Get-ITReport -Server <Server Name>
        This parameter allows you to specify a specific server instead of the module prompting you for a list. This can be used with any of the parameters above.


    .LINK
    http://bloggintechie.blogspot.com/
    https://chromebookparadise.wordpress.com/
    james.arnett@gmail.com
    #>
    [cmdletbinding()]Param([switch]$Hardware,[switch]$UpTime,[switch]$QFE,[switch]$ShutdownLog,[switch]$Service,[string]$Server,[switch]$QFEDaily)
    Function Get-ServerList{   
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "All files (*.*)| *.*"
        $OpenFileDialog.Title = "Please Specify Serverlist File:"
        $Show = $OpenFileDialog.ShowDialog()
        If ($Show -eq "OK"){
            Return $OpenFileDialog.filename
            }Else{
                Write-Error "Operation cancelled by user." -ErrorAction Stop
                }
        
    }
    Function Get-ReportPath{
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
    if(!$server){
        $ServerListPath = Get-ServerList
        $serverList = Get-Content $ServerListPath
        $ReportPath = Get-ReportPath
        $file_now = Get-Date -format MM-dd-yy.hhmmtt
        $now = Get-Date
        }
        else{
            $serverList = $Server
            $ReportPath = Get-ReportPath
            $file_now = Get-Date -format MM-dd-yy.hhmmtt
            $now = Get-Date
            }
    Function Get-ITUpTimeReport{
        write-host `n`rStarting UpTime Report Generation.... -ForegroundColor Yellow `n`r
        $ErrorActionPreference = "SilentlyContinue"
        $URD = "Uptime-Reports"
        New-Item -path $ReportPath -name $URD -ItemType directory -ErrorAction SilentlyContinue | out-null
        $TRD = "$file_now"
        New-Item -path $ReportPath\$URD -name $TRD -ItemType directory -ErrorAction SilentlyContinue | out-null
        ForEach($server in $serverlist){
            $TC = Test-Connection $Server -ErrorAction Stop -Quiet -Count 2
            if($TC){
                $UR = gwmi Win32_OperatingSystem -ComputerName $server -ErrorAction SilentlyContinue -ErrorVariable E1V | select @{LABEL='Server Name';E={$_.csname}},@{L=' ';E={''}},@{LABEL='LastBoot';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}},@{L='Error';E={''}}
                if((!$UR)-and(!$E1V)){
                    $EE = '"Server Name"," ","LastBoot","Error"'
                    $EE | select @{LABEL='Server Name';E={$server}},@{L=' ';E={''}},@{LABEL='LastBoot';EXPRESSION={''}},@{L='Error';E={'No WMI boot records found'}} | Export-Csv "$ReportPath\$URD\$file_now\Uptime Report.csv" -Append -NoTypeInformation
                    write-host No WMI boot records found on $server.... -ForegroundColor DarkRed -BackgroundColor Green `n`r
                    }elseif((!$UR)-and($E1V)){
                        Write-Host An error has occured when attempting to find boot records on $server -ForegroundColor Red -BackgroundColor White `n`r
                        $ER = '"Server Name"," ","LastBoot","Error"'
                        $ER | select @{LABEL='Server Name';E={$Server}},@{L=' ';E={''}},@{LABEL='LastBoot';EXPRESSION={''}},@{LABEL='Error';E={$E1V}} | Export-Csv "$ReportPath\$URD\$file_now\Uptime Report.csv" -Append -NoTypeInformation                    
                        }else{
                            $UR | Export-Csv "$ReportPath\$URD\$file_now\Uptime Report.csv" -Append -NoTypeInformation
                            write-host Uptime Report gathered for $server.... `n`r
                            }
                }else{
                    $EE = '"Server Name"," ","LastBoot","Error"'
                    $EE | select @{LABEL='Server Name';E={$Server}},@{L=' ';E={''}},@{LABEL='LastBoot';EXPRESSION={''}},@{LABEL='Error';E={'Failed to Connect'}} | Export-Csv "$ReportPath\$URD\$file_now\Uptime Report.csv" -Append -NoTypeInformation
                    write-host Failed to connect to $server.... -ForegroundColor Red -BackgroundColor Yellow `n`r
                    }  
        }
        Write-Host `n`rUPTime Report generation completed... -ForegroundColor Yellow `n`r
        Pause
    }
    Function Get-ITQFEReport{
        write-host `n`rStarting Quick Fix Engineering Report Generation.... -ForegroundColor Yellow `n`r
        [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $qfe_days = [Microsoft.VisualBasic.Interaction]::InputBox("How many days would you like the Quick Fix Engineering Report for?", "QFE Report", "2")
        $qfe_days = $qfe_days.trim();
        $ErrorActionPreference = "SilentlyContinue"
        $QRD = "QFE-Reports"
        New-Item -path $ReportPath -name $QRD -ItemType directory -ErrorAction SilentlyContinue | out-null
        $TRD = "$file_now"
        New-Item -path $ReportPath\$QRD -name $TRD -ItemType directory -ErrorAction SilentlyContinue | out-null
        foreach ($Server in $Serverlist){
            $TC = Test-Connection $Server -ErrorAction Stop -Quiet -Count 2
            if($TC){
	            $QFE = gwmi Win32_Quickfixengineering -ComputerName $Server -ErrorAction SilentlyContinue -ErrorVariable E2V | get-wmiobject -class win32_quickfixengineering | ?{$_.HotFixid -eq 'kb4012212' -or $_.HotFixid -eq 'kb4012214' -or $_.HotFixid -eq 'kb4012213' -or $_.HotFixid -eq 'kb4012598' -or $_.HotFixid -eq 'kb4018466'} | select @{N="Server Name";E={$server}},Description,HotFixID,InstalledON,Error
                if((!$QFE)-and(!$E2V)){
                    $QFE = '"Server Name","Description","HotFixID","InstalledOn","Error"'
                    $QFE | select @{N="Server Name";E={$Server}},@{N="Description";E={"No updates found for specified number of days($qfe_days)"}},HotFixID,InstalledON,Error | Export-Csv "$ReportPath\$QRD\$file_now\QFE Reports.csv" -Append -NoTypeInformation
                    Write-Host "No updates found for $Server for specified number of days($qfe_days)" -ForegroundColor DarkRed -BackgroundColor Green `n`r
                    }elseif((!$QFE)-and($E2V)){
                        Write-Host An error has occured when attempting to find QFE records on $server for $qfe_days days.... -ForegroundColor Red -BackgroundColor White `n`r
                        $ER = '"Server Name","Description","HotFixID","InstalledOn","Error"'
                        $ER | select @{N="Server Name";E={$server}},@{L='Description';E={'An error has occured'}},HotFixID,InstalledON,@{LABEL='Error';E={$E2V}} | Export-Csv "$ReportPath\$QRD\$file_now\QFE Reports.csv" -Append -NoTypeInformation                    
                        }else{
                           $QFE | Export-Csv "$ReportPath\$QRD\$file_now\QFE Reports.csv" -Append -NoTypeInformation
                           Write-Host QFE Report successfully generated for $Server `n`r 
                            }                
            }else{  
	            $EE = '"Server Name"'
                $EE | select @{LABEL='Server Name';E={$Server}},@{L='Description';E={'Failed to Connect'}},HotFixID,InstalledON,Error | Export-Csv "$ReportPath\$QRD\$file_now\QFE Reports.csv" -Append -NoTypeInformation
                write-host Failed to connect to $server.... -ForegroundColor Red -BackgroundColor Yellow `n`r
            }
        }
        write-host `n`rQFE Report generation completed... -ForegroundColor Yellow `n`r
        Pause
    }
    Function Get-ITDailyQFEReport{
        write-host `n`rStarting Daily Quick Fix Engineering Report Generation.... -ForegroundColor Yellow `n`r
        [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $qfe_days = 60
        $ErrorActionPreference = "SilentlyContinue"
        $QRD = "QFE-Daily-Reports"
        New-Item -path $ReportPath -name $QRD -ItemType directory -ErrorAction SilentlyContinue | out-null
        $TRD = "$file_now"
        New-Item -path $ReportPath\$QRD -name $TRD -ItemType directory -ErrorAction SilentlyContinue | out-null
        foreach ($Server in $Serverlist){
            $TC = Test-Connection $Server -ErrorAction Stop -Quiet -Count 2
            if($TC){
	            $QFE = Get-HotFix -ComputerName $Server -ErrorAction SilentlyContinue -ErrorVariable E2V | ? {$_.InstalledOn -ge $now.AddDays(-$qfe_days)} | select @{N="Server Name";E={$server}},Description,HotFixID,InstalledON,Error
                if((!$QFE)-and(!$E2V)){
                    $QFE = '"Server Name","Error"'
                    $QFE | select @{L="Server Name";E={$Server}},@{L='Error';E={''}} | Export-Csv "$ReportPath\$QRD\$file_now\QFE Daily Reports.csv" -Append -NoTypeInformation
                    Write-Host "No updates found for $Server" -ForegroundColor DarkRed -BackgroundColor Green `n`r
                    }elseif((!$QFE)-and($E2V)){
                        Write-Host An error has occured when attempting to find QFE records on $server -ForegroundColor Red -BackgroundColor White `n`r
                        $ER = '"Server Name","Error"'
                        $ER | select @{N="Server Name";E={$server}},@{LABEL='Error';E={$E2V}} | Export-Csv "$ReportPath\$QRD\$file_now\QFE Daily Reports.csv" -Append -NoTypeInformation                    
                        }else{
                           Write-Host Patches found for QFE Report on $Server moving on `n`r -BackgroundColor Magenta
                            }                
            }else{  
	            $EE = '"Server Name","Error"'
                $EE | select @{LABEL='Server Name';E={$Server}},@{L='Error';E={'Failed to Connect'}} | Export-Csv "$ReportPath\$QRD\$file_now\QFE Daily Reports.csv" -Append -NoTypeInformation
                write-host Failed to connect to $server.... -ForegroundColor Red -BackgroundColor Yellow `n`r
            }
        }
        write-host `n`rQFE Daily Report generation completed... -ForegroundColor Yellow `n`r
        Pause
    }
    Function Get-ITShutdownLogReport{
        write-host `n`rStarting Shutdown Log Report Generation....`n`r -ForegroundColor Yellow
        [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $NOE = [Microsoft.VisualBasic.Interaction]::InputBox("From Newest to Oldest, how many Shutdown events would you like to go back?", "Shutdown Events", "10")
        $NOE =  $NOE.Trim();
        $ErrorActionPreference = "SilentlyContinue"
        $SLD = "ShutdownLog-Reports"
        New-Item -path $ReportPath -name $SLD -ItemType directory -ErrorAction SilentlyContinue | out-null
        $TRD = "$file_now"
        New-Item -path $ReportPath\$SLD -name $TRD -ItemType directory -ErrorAction SilentlyContinue | out-null
        $EN = '"MachineName","Message","TimeWritten","UserName","PSComputerName","RunspaceId","PSShowComputerName","Error"'
        $EN | select @{L='MachineName';E={''}},@{L='Message';E={''}},@{L='TimeWritten';E={''}},@{L='UserName';E={''}},@{L='PSComputerName';E={''}},@{L='RunspaceId';E={''}},@{L='PSShowComputerName';E={''}},@{L='Error';E={''}} | Export-Csv "$ReportPath\$SLD\$file_now\ShutdownLog Report.csv" -Append -NoTypeInformation
        Foreach($Server in $ServerList){
            $TC = Test-Connection $Server -ErrorAction Stop -Quiet -Count 2
            if($TC){
                $SB = [scriptblock]::Create("Get-EventLog -ComputerName $Server -LogName System -Newest $NOE -InstanceId 2147484722 -ErrorAction SilentlyContinue | select MachineName,Message,TimeWritten,UserName")
                $EL = Start-Job -ScriptBlock $SB
                Wait-Job $EL -Timeout 60
                Stop-Job $EL
                $EL = Receive-Job $EL -ErrorVariable E3V
                if((!$EL)-and(!$E3v)){
                    $EN = '"MachineName","Message","TimeWritten","UserName","PSComputerName","RunspaceId","PSShowComputerName","Error"'
                    $EN | select @{L='MachineName';E={$server}},@{L='Message';E={''}},@{L='TimeWritten';E={''}},@{L='UserName';E={''}},@{L='PSComputerName';E={''}},@{L='RunspaceId';E={''}},@{L='PSShowComputerName';E={''}},@{L='Error';E={'No Logs Found'}} | Export-Csv "$ReportPath\$SLD\$file_now\ShutdownLog Report.csv" -Append -NoTypeInformation
                    Write-Host "No logs found for $server for specified number of logs($NOE)" -ForegroundColor Red -BackgroundColor white `n`r
                    }elseif((!$EL)-and($E3V)){
                        Write-Host An error has occured when attempting to find Shutdown Log records on $server -ForegroundColor Red -BackgroundColor White `n`r
                        $ER = '"MachineName","Message","TimeWritten","UserName","PSComputerName","RunspaceId","PSShowComputerName","Error"'
                        $ER | select @{L='MachineName';E={$Server}},@{L='Message';E={''}},@{L='TimeWritten';E={''}},@{L='UserName';E={''}},@{L='PSComputerName';E={''}},@{L='RunspaceId';E={''}},@{L='PSShowComputerName';E={''}},@{L='Error';E={$E3V}} | Export-Csv "$ReportPath\$SLD\$file_now\ShutdownLog Report.csv" -Append -NoTypeInformation                    
                        }else{
                            $EL | select MachineName,Message,TimeWritten,UserName,PSComputerName,RunspaceId,PSShowComputerName,@{L='Error';E={''}} | Export-Csv "$ReportPath\$SLD\$file_now\ShutdownLog Report.csv" -Append -NoTypeInformation
                            Write-Host Shutdown Log was successfully pulled for $Server...`n`r
                            }
            }Else{
                $EE = '"MachineName","Message","TimeWritten","UserName","PSComputerName","RunspaceId","PSShowComputerName","Error"'
                $EE | select @{L='MachineName';E={$Server}},@{L='Message';E={''}},@{L='TimeWritten';E={''}},@{L='UserName';E={''}},@{L='PSComputerName';E={''}},@{L='RunspaceId';E={''}},@{L='PSShowComputerName';E={''}},@{L='Error';E={'Failed to Connect'}} | Export-Csv "$ReportPath\$SLD\$file_now\ShutdownLog Report.csv" -Append -NoTypeInformation
                write-host Failed to connect to $server.... -ForegroundColor Red -BackgroundColor Yellow `n`r
            }
        }
        Write-Host `n`rShutdown Log Report generation has completed... -ForegroundColor Yellow `n`r
        Pause
    }
    Function Get-ITServiceReport{
        write-host `n`rStarting Service Report Generation.... -ForegroundColor Yellow `n`r
        $ErrorActionPreference = "SilentlyContinue"
        $SRD = "Service-Reports"
        New-Item -path $ReportPath -name $SRD -ItemType directory -ErrorAction SilentlyContinue | out-null
        $TRD = "$file_now"
        New-Item -path $ReportPath\$SRD -name $TRD -ItemType directory -ErrorAction SilentlyContinue | out-null
        foreach ($server in $serverlist){
            $TC = Test-Connection $Server -ErrorAction Stop -Quiet -Count 2
            if($TC){
                $srvfqdn = [System.Net.Dns]::GetHostByName("$server")
                $srvfqdn = $srvfqdn.HostName
                $SR = gwmi win32_service -ComputerName $srvfqdn -ErrorAction SilentlyContinue -ErrorVariable E4V
                if((!$SR)-and(!$E4V)){
                    $EE = '"Server Name"'
                    $EE | select @{LABEL='Server Name';E={$Server}} | Export-Csv "$ReportPath\$SRD\$file_now\Failed Service Report.csv" -Append -NoTypeInformation
                    write-host Service Report failed for $server -ForegroundColor Red -BackgroundColor White `n`r
                    }elseif((!$EL)-and($E4V)){
                        Write-Host An error has occured when attempting to gather Services on $server -ForegroundColor Red -BackgroundColor White `n`r
                        $ER = '"Server Name","Error"'
                        $ER | select @{LABEL='Server Name';E={$Server}},@{LABEL='Error';E={$E4V}} | Export-Csv "$ReportPath\$SRD\$file_now\ShutdownLog Report Error.csv" -Append -NoTypeInformation
                        }else{
                            $SR | export-csv "$ReportPath\$SRD\$file_now\$Server Service Report.csv" -NoTypeInformation
                            write-host Service Report for $server has been Generated`n`r
                            }
                }
                elseif($TC -eq $false){
                    $EE = '"Server Name"'
                    $EE | select @{LABEL='Server Name';E={$Server}} | Export-Csv "$ReportPath\$SRD\$file_now\Failed to Connect.csv" -Append -NoTypeInformation
                    write-host Failed to connect to $server.... -ForegroundColor Red -BackgroundColor Yellow `n`r
                }
        }
        write-host `n`rService Report Generation Completed... -ForegroundColor Yellow `n`r
        Pause
    }
    Function Get-ITHardwareReport{
        write-host `n`rStarting Hardware Report Generation.... -ForegroundColor Yellow `n`r
        $ErrorActionPreference = "SilentlyContinue"
        $HRD = "Hardware-Reports"
        New-Item -path $ReportPath -name $HRD -ItemType directory -ErrorAction SilentlyContinue | out-null
        $TRD = "$file_now"
        New-Item -path $ReportPath\$HRD -name $TRD -ItemType directory -ErrorAction SilentlyContinue | out-null
        $EN = '"Server Name","CPU","Mem","HDD 1 Letter","HDD 1 Freespace","HDD 1 Size","HDD 1 Provider Name","HDD 2 Letter","HDD 2 Freespace","HDD 2 Size","HDD 2 Provider Name","HDD 3 Letter","HDD 3 Freespace","HDD 3 Size","HDD 3 Provider Name","HDD 4 Letter","HDD 4 Freespace","HDD 4 Size","HDD 4 Provider Name","HDD 5 Letter","HDD 5 Freespace","HDD 5 Size","HDD 5 Provider Name","HDD 6 Letter","HDD 6 Freespace","HDD 6 Size","HDD 6 Provider Name","HDD 7 Letter","HDD 7 Freespace","HDD 7 Size","HDD 7 Provider Name","HDD 8 Letter","HDD 8 Freespace","HDD 8 Size","HDD 8 Provider Name","HDD 9 Letter","HDD 9 Freespace","HDD 9 Size","HDD 9 Provider Name","HDD 10 Letter","HDD 10 Freespace","HDD 10 Size","HDD 10 Provider Name","Error"'
        $EN | select @{L='Server Name';E={$Server}},
                @{L='CPU';E={''}},
                @{L='Mem';E={''}},
                @{L='HDD 1 Letter';E={''}},
                @{L='HDD 1 Freespace';E={''}},
                @{L='HDD 1 Size';E={''}},
                @{L='HDD 1 Provider Name';E={''}},
                @{L='HDD 2 Letter';E={''}},
                @{L='HDD 2 Freespace';E={''}},
                @{L='HDD 2 Size';E={''}},
                @{L='HDD 2 Provider Name';E={''}},
                @{L='HDD 3 Letter';E={''}},
                @{L='HDD 3 Freespace';E={''}},
                @{L='HDD 3 Size';E={''}},
                @{L='HDD 3 Provider Name';E={''}},
                @{L='HDD 4 Letter';E={''}},
                @{L='HDD 4 Freespace';E={''}},
                @{L='HDD 4 Size';E={''}},
                @{L='HDD 4 Provider Name';E={''}},
                @{L='HDD 5 Letter';E={''}},
                @{L='HDD 5 Freespace';E={''}},
                @{L='HDD 5 Size';E={''}},
                @{L='HDD 5 Provider Name';E={''}},
                @{L='HDD 6 Letter';E={''}},
                @{L='HDD 6 Freespace';E={''}},
                @{L='HDD 6 Size';E={''}},
                @{L='HDD 6 Provider Name';E={''}},
                @{L='HDD 7 Letter';E={''}},
                @{L='HDD 7 Freespace';E={''}},
                @{L='HDD 7 Size';E={''}},
                @{L='HDD 7 Provider Name';E={''}},
                @{L='HDD 8 Letter';E={''}},
                @{L='HDD 8 Freespace';E={''}},
                @{L='HDD 8 Size';E={''}},
                @{L='HDD 8 Provider Name';E={''}},
                @{L='HDD 9 Letter';E={''}},
                @{L='HDD 9 Freespace';E={''}},
                @{L='HDD 9 Size';E={''}},
                @{L='HDD 9 Provider Name';E={''}},
                @{L='HDD 10 Letter';E={''}},
                @{L='HDD 10 Freespace';E={''}},
                @{L='HDD 10 Size';E={''}},
                @{L='HDD 10 Provider Name';E={''}},
                @{L='Error';E={''}} | Export-Csv "$ReportPath\$HRD\$TRD\Hardware Report.csv" -Append -NoTypeInformation
        ForEach($server in $serverlist){
            $TC = Test-Connection $Server -ErrorAction Stop -Quiet -Count 2
            if($TC){
                $PR = gwmi win32_Processor -ComputerName $server -ErrorAction SilentlyContinue -ErrorVariable E1V
                $RR = gwmi win32_physicalmemory -ComputerName $server -ErrorAction SilentlyContinue -ErrorVariable E2V | ForEach-Object {$_.Capacity / 1GB}
                $HR = gwmi win32_logicaldisk -ComputerName $server -ErrorAction SilentlyContinue -ErrorVariable E3V
                $A = 0
                $B = ($HR.DeviceID.Count - 1)  
                do{      
                    Set-Variable -Value $HR.DeviceId[$A] -Name HDID$A
                    Set-Variable -Value $HR.ProviderName[$A] -Name HDPN$A
                    Set-Variable -Value $HR.Size[$A] -Name HDSZ$A
                    Set-Variable -Value $HR.FreeSpace[$A] -Name HDFS$A
                    $A++
                    }while($A -le $B)
                if((!$PR)-and(!$RR)-and(!$HR)){
                    Write-Host An error has occured when attempting to collect hardware information on $server -ForegroundColor Red -BackgroundColor White `n`r
                    $ER = '"Server Name","CPU","Mem","HDD 1 Letter","HDD 1 Freespace","HDD 1 Size","HDD 1 Provider Name","HDD 2 Letter","HDD 2 Freespace","HDD 2 Size","HDD 2 Provider Name","HDD 3 Letter","HDD 3 Freespace","HDD 3 Size","HDD 3 Provider Name","HDD 4 Letter","HDD 4 Freespace","HDD 4 Size","HDD 4 Provider Name","HDD 5 Letter","HDD 5 Freespace","HDD 5 Size","HDD 5 Provider Name","HDD 6 Letter","HDD 6 Freespace","HDD 6 Size","HDD 6 Provider Name","HDD 7 Letter","HDD 7 Freespace","HDD 7 Size","HDD 7 Provider Name","HDD 8 Letter","HDD 8 Freespace","HDD 8 Size","HDD 8 Provider Name","HDD 9 Letter","HDD 9 Freespace","HDD 9 Size","HDD 9 Provider Name","HDD 10 Letter","HDD 10 Freespace","HDD 10 Size","HDD 10 Provider Name","Error"'
                    $ER | select @{L='Server Name';E={$Server}},
                            @{L='CPU';E={''}},
                            @{L='Mem';E={''}},
                            @{L='HDD 1 Letter';E={''}},
                            @{L='HDD 1 Freespace';E={''}},
                            @{L='HDD 1 Size';E={''}},
                            @{L='HDD 1 Provider Name';E={''}},
                            @{L='HDD 2 Letter';E={''}},
                            @{L='HDD 2 Freespace';E={''}},
                            @{L='HDD 2 Size';E={''}},
                            @{L='HDD 2 Provider Name';E={''}},
                            @{L='HDD 3 Letter';E={''}},
                            @{L='HDD 3 Freespace';E={''}},
                            @{L='HDD 3 Size';E={''}},
                            @{L='HDD 3 Provider Name';E={''}},
                            @{L='HDD 4 Letter';E={''}},
                            @{L='HDD 4 Freespace';E={''}},
                            @{L='HDD 4 Size';E={''}},
                            @{L='HDD 4 Provider Name';E={''}},
                            @{L='HDD 5 Letter';E={''}},
                            @{L='HDD 5 Freespace';E={''}},
                            @{L='HDD 5 Size';E={''}},
                            @{L='HDD 5 Provider Name';E={''}},
                            @{L='HDD 6 Letter';E={''}},
                            @{L='HDD 6 Freespace';E={''}},
                            @{L='HDD 6 Size';E={''}},
                            @{L='HDD 6 Provider Name';E={''}},
                            @{L='HDD 7 Letter';E={''}},
                            @{L='HDD 7 Freespace';E={''}},
                            @{L='HDD 7 Size';E={''}},
                            @{L='HDD 7 Provider Name';E={''}},
                            @{L='HDD 8 Letter';E={''}},
                            @{L='HDD 8 Freespace';E={''}},
                            @{L='HDD 8 Size';E={''}},
                            @{L='HDD 8 Provider Name';E={''}},
                            @{L='HDD 9 Letter';E={''}},
                            @{L='HDD 9 Freespace';E={''}},
                            @{L='HDD 9 Size';E={''}},
                            @{L='HDD 9 Provider Name';E={''}},
                            @{L='HDD 10 Letter';E={''}},
                            @{L='HDD 10 Freespace';E={''}},
                            @{L='HDD 10 Size';E={''}},
                            @{L='HDD 10 Provider Name';E={''}},
                            @{L='Error';E={"$E1V $E2V $E3V"}} | Export-Csv "$ReportPath\$HRD\$TRD\Hardware Report.csv" -Append -NoTypeInformation
                    }else{
                        $ER = '"Server Name","CPU","Mem","HDD 1 Letter","HDD 1 Freespace","HDD 1 Size","HDD 1 Provider Name","HDD 2 Letter","HDD 2 Freespace","HDD 2 Size","HDD 2 Provider Name","HDD 3 Letter","HDD 3 Freespace","HDD 3 Size","HDD 3 Provider Name","HDD 4 Letter","HDD 4 Freespace","HDD 4 Size","HDD 4 Provider Name","HDD 5 Letter","HDD 5 Freespace","HDD 5 Size","HDD 5 Provider Name","HDD 6 Letter","HDD 6 Freespace","HDD 6 Size","HDD 6 Provider Name","HDD 7 Letter","HDD 7 Freespace","HDD 7 Size","HDD 7 Provider Name","HDD 8 Letter","HDD 8 Freespace","HDD 8 Size","HDD 8 Provider Name","HDD 9 Letter","HDD 9 Freespace","HDD 9 Size","HDD 9 Provider Name","HDD 10 Letter","HDD 10 Freespace","HDD 10 Size","HDD 10 Provider Name","Error"'
                        $ER | select @{L='Server Name';E={$Server}},
                            @{L='CPU';E={($PR.Name).count; "CPUs"}},
                            @{L='Mem';E={($RR | Measure-Object -Sum).Sum; "GB"}},
                            @{L='HDD 1 Letter';E={$HDID0}},
                            @{L='HDD 1 Freespace';E={[math]::round($HDFS0 / 1GB);"GB"}},
                            @{L='HDD 1 Size';E={[math]::round($HDSZ0 / 1GB);"GB"}},
                            @{L='HDD 1 Provider Name';E={$HDPN0}},
                            @{L='HDD 2 Letter';E={$HDID1}},
                            @{L='HDD 2 Freespace';E={[math]::round($HDFS1 / 1GB);"GB"}},
                            @{L='HDD 2 Size';E={[math]::round($HDSZ1 / 1GB);"GB"}},
                            @{L='HDD 2 Provider Name';E={$HDPN1}},
                            @{L='HDD 3 Letter';E={$HDID2}},
                            @{L='HDD 3 Freespace';E={[math]::round($HDFS2 / 1GB);"GB"}},
                            @{L='HDD 3 Size';E={[math]::round($HDSZ2 / 1GB);"GB"}},
                            @{L='HDD 3 Provider Name';E={$HDPN2}},
                            @{L='HDD 4 Letter';E={$HDID3}},
                            @{L='HDD 4 Freespace';E={[math]::round($HDFS3 / 1GB);"GB"}},
                            @{L='HDD 4 Size';E={[math]::round($HDSZ3 / 1GB);"GB"}},
                            @{L='HDD 4 Provider Name';E={$HDPN3}},
                            @{L='HDD 5 Letter';E={$HDID4}},
                            @{L='HDD 5 Freespace';E={[math]::round($HDFS4 / 1GB);"GB"}},
                            @{L='HDD 5 Size';E={[math]::round($HDSZ4 / 1GB);"GB"}},
                            @{L='HDD 5 Provider Name';E={$HDPN4}},
                            @{L='HDD 6 Letter';E={$HDID5}},
                            @{L='HDD 6 Freespace';E={[math]::round($HDFS5 / 1GB);"GB"}},
                            @{L='HDD 6 Size';E={[math]::round($HDSZ5 / 1GB);"GB"}},
                            @{L='HDD 6 Provider Name';E={$HDPN5}},
                            @{L='HDD 7 Letter';E={$HDID6}},
                            @{L='HDD 7 Freespace';E={[math]::round($HDFS6 / 1GB);"GB"}},
                            @{L='HDD 7 Size';E={[math]::round($HDSZ6 / 1GB);"GB"}},
                            @{L='HDD 7 Provider Name';E={$HDPN6}},
                            @{L='HDD 8 Letter';E={$HDID7}},
                            @{L='HDD 8 Freespace';E={[math]::round($HDFS7 / 1GB);"GB"}},
                            @{L='HDD 8 Size';E={[math]::round($HDSZ7 / 1GB);"GB"}},
                            @{L='HDD 8 Provider Name';E={$HDPN7}},
                            @{L='HDD 9 Letter';E={$HDID8}},
                            @{L='HDD 9 Freespace';E={[math]::round($HDFS8 / 1GB);"GB"}},
                            @{L='HDD 9 Size';E={[math]::round($HDSZ8 / 1GB);"GB"}},
                            @{L='HDD 9 Provider Name';E={$HDPN8}},
                            @{L='HDD 10 Letter';E={$HDID9}},
                            @{L='HDD 10 Freespace';E={[math]::round($HDFS9 / 1GB);"GB"}},
                            @{L='HDD 10 Size';E={[math]::round($HDSZ9 / 1GB);"GB"}},
                            @{L='HDD 10 Provider Name';E={$HDPN9}},
                            @{L='Error';E={"$E1V $E2V $E3V"}} | Export-Csv "$ReportPath\$HRD\$TRD\Hardware Report.csv" -Append -NoTypeInformation

                        write-host Hardware Report gathered for $server.... `n`r
                        }
                }else{
                    $EE = '"Server Name","CPU","Mem","HDD 1 Letter","HDD 1 Freespace","HDD 1 Size","HDD 1 Provider Name","HDD 2 Letter","HDD 2 Freespace","HDD 2 Size","HDD 2 Provider Name","HDD 3 Letter","HDD 3 Freespace","HDD 3 Size","HDD 3 Provider Name","HDD 4 Letter","HDD 4 Freespace","HDD 4 Size","HDD 4 Provider Name","HDD 5 Letter","HDD 5 Freespace","HDD 5 Size","HDD 5 Provider Name","HDD 6 Letter","HDD 6 Freespace","HDD 6 Size","HDD 6 Provider Name","HDD 7 Letter","HDD 7 Freespace","HDD 7 Size","HDD 7 Provider Name","HDD 8 Letter","HDD 8 Freespace","HDD 8 Size","HDD 8 Provider Name","HDD 9 Letter","HDD 9 Freespace","HDD 9 Size","HDD 9 Provider Name","HDD 10 Letter","HDD 10 Freespace","HDD 10 Size","HDD 10 Provider Name","Error"'
                    $EE | select @{L='Server Name';E={$Server}},
                            @{L='CPU';E={''}},
                            @{L='Mem';E={''}},
                            @{L='HDD 1 Letter';E={''}},
                            @{L='HDD 1 Freespace';E={''}},
                            @{L='HDD 1 Size';E={''}},
                            @{L='HDD 1 Provider Name';E={''}},
                            @{L='HDD 2 Letter';E={''}},
                            @{L='HDD 2 Freespace';E={''}},
                            @{L='HDD 2 Size';E={''}},
                            @{L='HDD 2 Provider Name';E={''}},
                            @{L='HDD 3 Letter';E={''}},
                            @{L='HDD 3 Freespace';E={''}},
                            @{L='HDD 3 Size';E={''}},
                            @{L='HDD 3 Provider Name';E={''}},
                            @{L='HDD 4 Letter';E={''}},
                            @{L='HDD 4 Freespace';E={''}},
                            @{L='HDD 4 Size';E={''}},
                            @{L='HDD 4 Provider Name';E={''}},
                            @{L='HDD 5 Letter';E={''}},
                            @{L='HDD 5 Freespace';E={''}},
                            @{L='HDD 5 Size';E={''}},
                            @{L='HDD 5 Provider Name';E={''}},
                            @{L='HDD 6 Letter';E={''}},
                            @{L='HDD 6 Freespace';E={''}},
                            @{L='HDD 6 Size';E={''}},
                            @{L='HDD 6 Provider Name';E={''}},
                            @{L='HDD 7 Letter';E={''}},
                            @{L='HDD 7 Freespace';E={''}},
                            @{L='HDD 7 Size';E={''}},
                            @{L='HDD 7 Provider Name';E={''}},
                            @{L='HDD 8 Letter';E={''}},
                            @{L='HDD 8 Freespace';E={''}},
                            @{L='HDD 8 Size';E={''}},
                            @{L='HDD 8 Provider Name';E={''}},
                            @{L='HDD 9 Letter';E={''}},
                            @{L='HDD 9 Freespace';E={''}},
                            @{L='HDD 9 Size';E={''}},
                            @{L='HDD 9 Provider Name';E={''}},
                            @{L='HDD 10 Letter';E={''}},
                            @{L='HDD 10 Freespace';E={''}},
                            @{L='HDD 10 Size';E={''}},
                            @{L='HDD 10 Provider Name';E={''}},
                            @{L='Error';E={"Failed to Connect"}} | Export-Csv "$ReportPath\$HRD\$TRD\Hardware Report.csv" -Append -NoTypeInformation
                    write-host Failed to connect to $server.... -ForegroundColor Red -BackgroundColor Yellow `n`r
                    }  
        }
        Write-Host `n`rHardware Report generation completed... -ForegroundColor Yellow `n`r
    }
    if($Hardware){
        Get-ITHardwareReport
        }
    if($UpTime){
        Get-ITUpTimeReport
        }
    if($QFE){
        Get-ITQFEReport
            }
    if($ShutdownLog){
        Get-ITShutdownLogReport
        }
    if($Service){
        Get-ITServiceReport
        }
    if($QFEDaily){
        Get-ITDailyQFEReport
        }
    if((!$Service)-and(!$ShutdownLog)-and(!$QFE)-and(!$UpTime)-and(!$Hardware)){
        Get-ITUpTimeReport
        Get-ITQFEReport
        Get-ITShutdownLogReport
        Get-ITServiceReport
        Get-ITHardwareReport
        }
    
}

