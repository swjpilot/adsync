#############################  User Variables
$Spassword = ConvertTo-SecureString -String "passw0rd" -AsPlainText -Force
$Slogin = "DOMAIN\User"
$workingdir = ".\Headcount\"
$filename = "US and CA Headct $(get-date -format `"MM.dd.yy`").xlsx"
$url = "https://portal.company.com/SiteDirectory/hr/headcount_docs/Headcount%20Reports/US%20and%20CA%20Headct%20$(get-date -format `"MM.dd.yy`").xlsx"
$url2 = "https://portal.company.com/SiteDirectory/hr/headcount_docs/Headcount%20Reports/US%20and%20CA%20Headct%20$(get-date -format `"M.d.yy`").xlsx"
$email = "serverteam@company.com"
#############################
. ".\functions.ps1"
$outfile = $workingdir + $filename
$pmrycred =  New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $Slogin, $Spassword
Invoke-WebRequest $url -OutFile $outfile -Credential $pmrycred
if (!(Test-Path $outfile)){
	Invoke-WebRequest $url2 -OutFile $outfile -Credential $pmrycred
}

if (Test-Path $outfile) {
	Send-MailMessage -To $email -From "headcount@company.com"	-SmtpServer "smtp.company.com" -Subject "New Headcount Report Downloaded $(get-date -format `"MM.dd.yy`")" -Body "A new headcount report has been downloaded on $(get-date -format `"MM.dd.yy`")"
	$Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $wb = $Excel.Workbooks.Open($outfile)
    foreach ($ws in $wb.Worksheets)
    {
        $ws.SaveAs($workingdir + "tolt" + ".csv", 6)
    }
    $Excel.Quit()
	.\OracleCompare.ps1
	.\OracleImport.ps1
}