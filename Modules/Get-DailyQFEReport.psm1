Function Get-DailyQFEReport{
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