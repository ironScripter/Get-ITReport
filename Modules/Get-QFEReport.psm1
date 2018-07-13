Function Get-QFEReport{
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
}