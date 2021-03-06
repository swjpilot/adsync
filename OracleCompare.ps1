
###################### Settings Begin
$importfile = "tolt.csv"
$archivepath = ".\History\"
$adpcontact = "user@adp.com"
$toltcontacts = "serverteam@company.com"
$departments = Import-CSV "Departments.csv"
$logfile = ".\log\Compare_$(get-date -format `"yyyyMMdd_hhmmtt`").log"
$comparefile = ".\Compare.csv"
$exceptionsfile = ".\CompareExceptions.csv"
$exceptionsout = @()
$compareout = @()
$ageout = (Get-Date).AddDays(-90)
$adserver = "ad.domain.com"
$samanageuser = "user@company.com"
$samanagepass = "passw0rd"
$samanageassignee = "assignee@company.com"
###################### Settings End
. ".\functions.ps1"
Import-module ActiveDirectory
$PSDefaultParameterValues = @{"*-AD*:Server"="$adserver"}

LogWrite "******************Begin File Check******************"
if (Test-Path $importfile){
	LogWrite "File Found preparing to process"
	$importdata = Import-CSV $importfile
}
else {
	LogWrite "File not found Emailing ADP"
#	Send-MailMessage -To $adpcontact  -Cc $toltcontacts -Subject "Tolt Solutions Nightly ADP Processing for $(get-date -format `"MM-dd-yyyy`")" -Body "We did not receive the nightly ADP download can you please check the scheduled task on your end and verify that it ran." -SmtpServer gvsmtp.d06.us -From "Scott.John@toltsolutions.com"
#	LogWrite "Email Sent to ADP ADP"
	exit
}
LogWrite "******************End File Check******************"

LogWrite "******************Begin File Rotation******************"
Copy-Item -Path $importfile -Destination ($archivepath + "TEMP_FILE_DO_NOT_OPEN.csv")
Copy-Item -Path $importfile -Destination ($archivepath + "ToltHeadCount_$(get-date -format `"MM-dd-yyyy`").csv")
Get-ChildItem -Path $archivepath -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $ageout } | Remove-Item -Force
LogWrite "******************End File Rotation******************"

LogWrite "******************Begin ADP Processing $(get-date -format `"yyyy-MM-dd_hh:mm tt`")******************"
$importdata | ForEach-Object {
	if ($_."Company" -eq "125") {
		$num = $_."Employee Number" -as [int]
		$fname = $_."First Name"
		$lname = $_."Last Name"
		$mgrnum = $_."Supervisor" -as [int]
		$mgrfname = $_."Supervisor First Name"
		$mgrlname = $_."Supervisor Last Name"
		$deptnum = $_."Business Unit" -as [int]
		$dept = $_."Organization Name"
		$title = $_."Job"
		$office = $_."Location Code"
		$state = $_."State"
		$city = $_."City"
		$change = $false

		LogWrite "Processing ************* $num $fname $lname"
		#       Try { $exists = Get-ADUser -Filter {(givenName -like $fname) -and (surname -like $lname)} }
		Try { $exists = Get-ADUser -Filter {(employeeID -eq $num)} -Properties * }
		Catch { }
		If ($exists) {
			LogWrite "Found Employee in Directory"
			if ($exists.department -NotLike "$deptnum-$dept") {
				LogWrite "Change to Department"
				$change = $true
			}
			Try { $manager = Get-ADUser -Filter {(employeeID -eq $mgrnum)} -Properties * }
		    Catch { 
				LogWrite "No Manager Found $mgrnum"
			}
			if ($exists.manager -ne $manager) {
				LogWrite "Change to Manager"
				$manager = Get-ADUser $exists.manager -Properties *
				$change = $true
			}
			if ($exists.title -NotLike $title) {
				LogWrite "Change to Title"
				$change = $true
			}
			if ($exists.office -NotLike $office) {
				LogWrite "Change to Office"
				$change = $true
			}
			if ($exists.state -NotLike $state) {
				LogWrite "Change to State"
				$change = $true
			}
			if ($exists.city -NotLike $city) {
				LogWrite "Change to City"
				$change = $true
			}
		}
		else {
		    $lineitem = "" | Select "Oracle_EID","First_Name","Last_Name"
		    $lineitem.Oracle_EID = $num
		    $lineitem.First_Name = $fname
		    $lineitem.Last_Name = $lname
		    $exceptionsout += $lineitem
		    $lineitem = $null
		}
		If ($change) {
			LogWrite "Change Detected ++++++++++++++"
			$lineitem = "" | Select "Type","Oracle_EID","First_Name","Last_Name","Manager_EID","Manager_First_Name","Manager_Last_Name","Title","Department","Office","City","State"
		    $lineitem.Type = "Oracle"
			$lineitem.Oracle_EID = $num
		    $lineitem.First_Name = $fname
		    $lineitem.Last_Name = $lname
			$lineitem.Manager_EID = $mgrnum
			$lineitem.Manager_First_Name = $mgrfname 
			$lineitem.Manager_Last_Name = $mgrlname
			$lineitem.Title = $title
			$lineitem.Department = "$deptnum-$dept"
			$lineitem.Office = $office
			$lineitem.City = $city
			$lineitem.State = $state
			LogWrite $lineitem
			$compareout += $lineitem
			$lineitem = $null
			$lineitem = "" | Select "Type","Oracle_EID","First_Name","Last_Name","Manager_EID","Manager_First_Name","Manager_Last_Name","Title","Department","Office","City","State"
		    $lineitem.Type = "AD"
			$lineitem.Oracle_EID = $exists.employeeID
		    $lineitem.First_Name = $exists.givenName
		    $lineitem.Last_Name = $exists.surname
			$lineitem.Manager_EID = $manager.employeeID
			$lineitem.Manager_First_Name = $manager.givenName
			$lineitem.Manager_Last_Name = $manager.surName
			$lineitem.Title = $exists.title
			$lineitem.Department = $exists.department
			$lineitem.Office = $exists.office
			$lineitem.City = $exists.city
			$lineitem.State = $exists.st
			LogWrite $lineitem
			$compareout += $lineitem
			$lineitem = $null
		}

	$exists = $null
	$manager = $null
	$dept = $null
	LogWrite "Processing Complete #########################"
	}
}


$exceptionsout | export-csv $exceptionsfile
$compareout | export-csv $comparefile
Send-MailMessage -To $toltcontacts -Subject "Nightly Oracle Processing for $(get-date -format `"MM-dd-yyyy`")" -Body "The Nightly Oracle File Compare is attached for review `r`n Please run the following script when you are done reviewing the proposed changes to Active Directory \\d06.us\NETLOGON\PS\OracleImport.ps1 " -SmtpServer gvsmtp.d06.us -From "OracleImport@toltsolutions.com" -Attachments $comparefile
LogWrite "Email Sent"
$Slogin = $samanageuser
$Spassword = ConvertTo-SecureString -String $samanagepass -AsPlainText -Force
$credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $Slogin, $Spassword
Invoke-WebRequest -Uri "https://app.samanage.com/api.xml" -Credential $credential -Headers @{'Accept'='application/vnd.samanage.v1+xml'} -SessionVariable loggedin
$XML = "<incident>"
$XML += "<name>Review and Apply Oracle Headcount File for $(get-date -format `"MM-dd-yyyy`")</name>"
$XML += "<priority>Medium</priority>"
$XML += "<requester><email>$samanageuser</email></requester>"
$XML += "<description>Review and Apply Oracle Headcount File for $(get-date -format `"MM-dd-yyyy`") File is located here \\d06.us\gv\ftp\adp\Compare.csv</description>"
$XML += "<assignee><email>$samanageassignee</email></assignee>"
$XML += "</incident>"
$url = "https://api.samanage.com/incidents.xml"
$XML
#Invoke-RestMethod -Uri $url -Method POST -Body $XML -WebSession $loggedin -ContentType "text/xml"
LogWrite "Ticket Opened"
LogWrite "******************End ADP Processing $(get-date -format `"yyyy-MM-dd_hh:mm tt`")******************"
exit