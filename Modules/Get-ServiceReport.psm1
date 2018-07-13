Function Get-ServiceReport{
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
    }