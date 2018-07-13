Function Get-HardwareReport{
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