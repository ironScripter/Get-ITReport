Function Get-UpTimeReport{
        write-host `n`rStarting UpTime Report Generation.... -ForegroundColor Yellow `n`r
        $ErrorActionPreference = "SilentlyContinue"
        $UptimeReportsDir = "Uptime-Reports"
        if(!(Test-Path -Path $ReportPath\$UptimeReportsDir )){
            New-Item -path $ReportPath -name $UptimeReportsDir -ItemType directory -ErrorAction Stop | out-null
        }
        $TRD = "$file_now"
        if(!(Test-Path -Path $ReportPath\$UptimeReportsDir\$TRD)){
            New-Item -path $ReportPath\$UptimeReportsDir -name $TRD -ItemType directory -ErrorAction Stop | out-null
        }
        ForEach($server in $serverlist){
            $TC = Test-Connection $Server -ErrorAction Stop -Quiet -Count 2
            if($TC){
                $UR = gwmi Win32_OperatingSystem -ComputerName $server -ErrorAction SilentlyContinue -ErrorVariable E1V | select @{LABEL='Server Name';E={$_.csname}},@{L=' ';E={''}},@{LABEL='LastBoot';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}},@{L='Error';E={''}}
                if((!$UR)-and(!$E1V)){
                    $EE = '"Server Name"," ","LastBoot","Error"'
                    $EE | select @{LABEL='Server Name';E={$server}},@{L=' ';E={''}},@{LABEL='LastBoot';EXPRESSION={''}},@{L='Error';E={'No WMI boot records found'}} | Export-Csv "$ReportPath\$UptimeReportsDir\$file_now\Uptime Report.csv" -Append -NoTypeInformation
                    write-host No WMI boot records found on $server.... -ForegroundColor DarkRed -BackgroundColor Green `n`r
                    }elseif((!$UR)-and($E1V)){
                        Write-Host An error has occured when attempting to find boot records on $server -ForegroundColor Red -BackgroundColor White `n`r
                        $ER = '"Server Name"," ","LastBoot","Error"'
                        $ER | select @{LABEL='Server Name';E={$Server}},@{L=' ';E={''}},@{LABEL='LastBoot';EXPRESSION={''}},@{LABEL='Error';E={$E1V}} | Export-Csv "$ReportPath\$UptimeReportsDir\$file_now\Uptime Report.csv" -Append -NoTypeInformation                    
                        }else{
                            $UR | Export-Csv "$ReportPath\$UptimeReportsDir\$file_now\Uptime Report.csv" -Append -NoTypeInformation
                            write-host Uptime Report gathered for $server.... `n`r
                            }
                }else{
                    $EE = '"Server Name"," ","LastBoot","Error"'
                    $EE | select @{LABEL='Server Name';E={$Server}},@{L=' ';E={''}},@{LABEL='LastBoot';EXPRESSION={''}},@{LABEL='Error';E={'Failed to Connect'}} | Export-Csv "$ReportPath\$UptimeReportsDir\$file_now\Uptime Report.csv" -Append -NoTypeInformation
                    write-host Failed to connect to $server.... -ForegroundColor Red -BackgroundColor Yellow `n`r
                    }  
        }
        Write-Host `n`rUPTime Report generation completed... -ForegroundColor Yellow `n`r
        Pause
    }