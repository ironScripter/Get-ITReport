Function Get-ShutdownLogReport{
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
    }