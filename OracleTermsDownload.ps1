#############################  User Variables
$Spassword = ConvertTo-SecureString -String "uNgZJDS7d3wLrrs9" -AsPlainText -Force
$Slogin = "PMRY\tolt.headcount"
$workingdir = "C:\Users\johns\Dropbox\_Tolt\_activeprojects\Headcount\"
$filename = "US and CA Terms $(get-date -format `"MM.dd.yy`").xlsx"
$url = "https://portal.pomeroy.com/SiteDirectory/hr/headcount_docs/Termination%20Reports/US%20and%20CA%20Terms%20$(get-date -format `"MM.dd.yy`").xlsx"
$url2 = "https://portal.pomeroy.com/SiteDirectory/hr/headcount_docs/Termination%20Reports/US%20and%20CA%20Terms%20$(get-date -format `"M.d.yy`").xlsx"
$email = "serverteam@toltsolutions.com"
#############################

$outfile = $workingdir + $filename
$pmrycred =  New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $Slogin, $Spassword
Invoke-WebRequest $url -OutFile $outfile -Credential $pmrycred
if (!(Test-Path $outfile)){
	Invoke-WebRequest $url2 -OutFile $outfile -Credential $pmrycred
}
if (Test-Path $outfile) {
	Send-MailMessage -To $email -From "Terminations@toltsoltions.com"	-SmtpServer "gvsmtp.d06.us" -Subject "New Terminations Report Downloaded $(get-date -format `"MM.dd.yy`")" -Body "A new terminations report has been downloaded on $(get-date -format `"MM.dd.yy`")"
	$Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $wb = $Excel.Workbooks.Open($outfile)
    foreach ($ws in $wb.Worksheets)
    {
        $ws.SaveAs($workingdir + "terms" + ".csv", 6)
    }
    $Excel.Quit()
	c:\users\johns\Dropbox\Scripts\powershell\OracleTerms.ps1
}