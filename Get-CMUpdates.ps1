$Ous = Get-Content ($PSScriptRoot + "\OUs.txt")
$objOus = @()

ForEach($Ou in $Ous)
{
    $Tokens = $Ou.Split(";")
    $Ou = $Tokens[0];
    $Contact = $Tokens[1];
    Write-Host "Checking " $Ou -ForegroundColor Yellow
    Write-Host "Contact is " $Contact -ForegroundColor Yellow

    $objOu = New-Object System.Object
    $objOu | Add-Member -MemberType NoteProperty -Name "Name" $Ou
    $objOu | Add-Member -MemberType NoteProperty -Name "Contact" $Contact
    $objOu | Add-Member -MemberType NoteProperty -Name "Computers" @()

    $Computers = Get-ADComputer -Filter * -SearchBase $Ou
    ForEach($Computer in $Computers){
        # $Computer | fl
        Write-Host "Querying "$Computer.Name -ForegroundColor Yellow

        $objComputer = New-Object System.Object
        $objComputer | Add-Member -MemberType NoteProperty -Name "Name" $Computer.Name
        $objComputer | Add-Member -MemberType NoteProperty -Name "Updates" @()
        $objComputer | Add-Member -MemberType NoteProperty -Name "Success" 1
        if($Computer.Name -ne "UKFALMSC01"){
            Try {
                $Updates = Get-WmiObject -ComputerName $Computer.DNSHostName -Query "SELECT * FROM CCM_SoftwareUpdate" -Namespace "ROOT\ccm\ClientSDK"
                if($?){
                    $objComputer.Updates = $Updates
                } else {
                    throw $error[0].Exception
                }
            } Catch {
                $objComputer.Success = 0
                Write-Host "Caught exception!" -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
        }
        
        $objOu.Computers += $objComputer
    }
    $objOus += $objOu
}

ForEach($Ou in $objOus){
    if($Ou.Length -gt 0){
        $UpdateTable = ""
        Write-Host $Ou.Name
        $Tokens = $OuName = $Ou.Name.Split(",")
        $OuName = $Tokens[0].Substring(($Tokens[0].IndexOf('=')+1))
        Write-Host "Ou Name: " $OuName

        ForEach($Computer in $Ou.Computers){
            Write-Host "    " $Computer.Name
            if($Computer.Success){
                $Count = 0
                if($Computer.Updates.Count -gt 0){
                    ForEach($Update in $Computer.Updates){
                        Write-Host ("        " + $Update.ArticleID + "," + $Update.Name)
                        if($Count -eq 0){
                            $UpdateTable += "<tr><td style='text-align: center'>" + $Computer.Name + "</td><td style='text-align: center'>KB" + $Update.ArticleID + "</td></tr>"
                        } else {
                            $UpdateTable += "<tr><td style='text-align: center'>&nbsp;</td><td style='text-align: center'>KB" + $Update.ArticleID + "</td></tr>"
                        }
                        $Count++
                    }
                } else {
                    $UpdateTable += "<tr><td style='text-align: center;'>" + $Computer.Name + "</td><td><p style='color: #17c600; text-align: center;'>No patches pending!</p></td></tr>"
                }
            } else {
                Write-Host "    There was an error querying this computer!" -ForegroundColor Red
                $UpdateTable += "<tr><td style='text-align: center'>" + $Computer.Name + "</td><td><p style='color: #be2625; text-align: center;'>There was an error connecting to this computer, please check.</p></td></tr>"
            }
        }
        Write-Host ""
        $SmtpServer = "nldataex01.ad.fugro.com"
        $SmtpFrom = "IT OneFugro Service Desk<onefugroservicedesk@fugro.com>"
        if($OuName -eq "SCCM"){
            $SmtpCc = ""
        } else {
            $SmtpCc = "gistsccmadmins@fugro.com"
        }
        $MessageSubject = "ACTION REQUIRED: Non-compliant " + $OuName + " Computers"

        [string] $body = Get-Content ($PSScriptRoot + "\UpdateCompliance.html")
        $body = $body -replace "{UPDATE_TABLE}", $UpdateTable

        $Attachments = $(($PSScriptRoot + "\\OneFugro-Notification.png"))
        #Send-Email -SmtpFrom $SmtpFrom -SmtpTo $Ou.Contact -MessageSubject $MessageSubject -Body $Body -Attachments $Attachments
        #Send-Email -SmtpFrom $SmtpFrom -SmtpTo $SmtpCc -MessageSubject $MessageSubject -Body $Body -Attachments $Attachments

        Send-MailMessage -SmtpServer $SmtpServer -From $SmtpFrom -To $Ou.Contact -Subject $MessageSubject -BodyAsHtml $body -Attachments $Attachments
        if($SmtpCc.Length -gt 0){
            Send-MailMessage -SmtpServer $SmtpServer -From $SmtpFrom -To $SmtpCc -Subject $MessageSubject -BodyAsHtml $body -Attachments $Attachments
        }
    }
}
