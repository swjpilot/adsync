[CmdletBinding()]
Param(
	[Parameter(Mandatory=$False)]
	[switch]$Connected = $false
)
###################### Settings Begin
$Slogin = "service@domain.com"
$adpcontact = "user@adp.com"
$contacts = "serverteam@domain.com"
$Spassword = ConvertTo-SecureString -String "passw0rd" -AsPlainText -Force
$importfile = ".\Headcount\terms.csv"
$exclusionsfile = ".\Headcount\exclusions.csv"
$archivepath = ".\Headcount\History\"
$logfile = ".\Headcount\log\Terms_$(get-date -format `"yyyyMMdd_hhmmtt`").log"
$ageout = (Get-Date).AddDays(-90)
$adserver = "ad.domain.com"
$samanageuser = "user@company.com"
$samanagepass = "passw0rd"
###################### Settings End
Import-module ActiveDirectory
Import-Module MSOnline
$PSDefaultParameterValues = @{"*-AD*:Server"="$adserver"}
Function LogWrite
{
   Param ([string]$logstring)

   Add-content $logfile -value $logstring
   Write-Host $logstring
}
if ($Connected -eq $false){
	$O365Cred = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $Slogin, $Spassword
	$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
	Import-PSSession $O365Session -AllowClobber
	Connect-MsolService –Credential $O365Cred
}

LogWrite "******************Begin File Check******************"
if (Test-Path $importfile){
	LogWrite "File Found preparing to process"
	$importdata = Import-CSV $importfile
}
else {
	LogWrite "File not found Emailing ADP"
#	Send-MailMessage -To $adpcontact  -Cc $contacts -Subject "Tolt Solutions Nightly ADP Processing for $(get-date -format `"MM-dd-yyyy`")" -Body "We did not receive the nightly ADP download can you please check the scheduled task on your end and verify that it ran." -SmtpServer smtp.domain.us -From "admin@domain.com"
#	LogWrite "Email Sent to ADP ADP"
	exit
}
LogWrite "******************End File Check******************"

LogWrite "******************Begin File Rotation******************"
Copy-Item -Path $importfile -Destination ($archivepath + "ToltTerminations_$(get-date -format `"MM-dd-yyyy`").csv")
Get-ChildItem -Path $archivepath -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $ageout } | Remove-Item -Force
LogWrite "******************End File Rotation******************"

$terms = Import-Csv $importfile
$exclusions = Import-Csv $exclusionsfile
$terms | ForEach-Object {
	if ($_."Company" -eq "125") {
		LogWrite "******************Begin User Deactivation******************"
		$eid = $_."Employee Number"
		$exclusions | ForEach-Object {
			if ($_."excludedEID" -eq $eid){
				$exclude = $true
			}
		}
		if ($exclude -eq $false){
			$user = Get-ADUser -Filter {EmployeeID -like $eid} -Properties *
		}
		else {
			$user = $null
		}
		if ($user) {
			Logwrite "Found User $eid"
			$SAMAccountName = $user.samaccountname
			$fn = $user.givenname
			$ln = $user.sn
			$UPN = $user.UserPrincipalName
			
			Disable-ADAccount $SAMAccountName
			Get-ADUser -Filter {(surname -like $ln) -and (givenname -like $fn)} -Server 10.253.0.50 | Disable-ADAccount -Server 10.253.0.50
			$user | Set-aduser -Replace @{msExchHideFromAddressLists="TRUE"}
			LogWrite "AD Account disabled for $fn $ln"
			Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "kyrus:ENTERPRISEPACK"
			LogWrite "365 E3 License Removed for $fn $ln"
			$Slogin = $samanageuser
			$Spassword = ConvertTo-SecureString -String $samanagepass -AsPlainText -Force
			$credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $Slogin, $Spassword
			Invoke-WebRequest -Uri "https://app.samanage.com/api.xml" -Credential $credential -Headers @{'Accept'='application/vnd.samanage.v1+xml'} -SessionVariable loggedin
			[xml]$samanageuser = Invoke-RestMethod -Uri "https://toltsolutions.samanage.com/users.xml?email=$UPN" -Method Get -WebSession $loggedin
			$samanagenumber = $samanageuser.users.user.id
			$updateXML = "<user>"
			$updateXML += "<disabled>true</disabled>"
			$updateXML += "</user>"
			$url = "https://api.samanage.com/users/" + $samanagenumber + ".xml"
			Invoke-RestMethod -Uri $url -Method PUT -Body $updateXML -WebSession $loggedin -ContentType "text/xml"
			LogWrite "Samanage Account Disabled for $fn $ln"
			Invoke-Sqlcmd -Query "Update dbo.person set UserEnabled=0 where name like '$SAMAccountName'" -ServerInstance gvsql2.d06.us -Database Chiyu
			LogWrite "Badge Disabled for $fn $ln"	
			LogWrite "******************End User Deactivation******************"
		}
	}

}