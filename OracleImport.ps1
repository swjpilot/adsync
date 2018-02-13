
######################### User Variables
$importfile = ".\Headcount\tolt.csv"
$exceptionsfile = ".\Headcount\ImportExceptions.csv"
$exceptionsmgrfile = ".\Headcount\ExceptionsMgr.csv"
$logfile = ".\log\OracleImport_$(get-date -format `"yyyyMMdd_hhmmtt`").log"
##########################
. ".\functions.ps1"
Import-module ActiveDirectory
LogWrite "******************Begin Oracle Processing $(get-date -format `"yyyy-MM-dd_hh:mmtt`")******************"
$exceptionsout = @()
$exceptionsmgr = @()
Import-CSV $importfile | ForEach-Object {
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
		$tolteid = $_."Ve Id"
		LogWrite "Processing ************* $num $fname $lname"

	#       Try { $exists = Get-ADUser -Filter {(givenName -like $fname) -and (surname -like $lname)} }
	    Try { $exists = Get-ADUser -Filter {(employeeID -eq $num)} }
	    Catch { }
	    If ($exists) {
	        LogWrite "Found in Directory"
	#           $exists | Set-ADUser -EmployeeID $num
	#           Try { $manager = Get-ADUser -Filter {(givenName -like $mgrfname) -and (surname -like $mgrlname)} }
	        Try { $manager = Get-ADUser -Filter {(employeeID -eq $mgrnum)} }
	        Catch {
				LogWrite "Manager Not Found"
	            $entry = "" | Select "Oracle_EID","First_Name","Last_Name","Manager_EID","Manager_First_Name","Manager_Last_Name"
	            $entry.Oracle_EID = $num
	            $entry.First_Name = $fname
	            $entry.Last_Name = $lname
				$entry.Manager_EID = $mgrnum
	            $entry.Manager_First_Name = $mgrfname
	            $entry.Manager_Last_Name = $mgrlname
	            $exceptionsmgr += $entry
	            $entry = $null
	        }
	        If ($manager) {
	            LogWrite "Found Manager"
	            $exists | Set-ADUser -Manager $manager.SamAccountName
	        }
	        $exists | Set-ADUser -Department "$deptnum-$dept"
			$exists | Set-ADUser -Clear departmentNumber
			$exists | Set-ADUser -Add @{departmentNumber="$deptnum"}
	        $exists | Set-ADUser -Title $title
	        $exists | Set-ADUser -Office $office
	        $exists | Set-ADUser -State $state
			$exists | Set-ADUser -City $city
			$exists | Set-ADUser -EmployeeNumber $toltEID

	    }
	    else {
			LogWrite "Not Found in Directory"
	        $lineitem = "" | Select "File_Number","First_Name","Last_Name"
	        $lineitem.File_Number = $num
	        $lineitem.First_Name = $fname
	        $lineitem.Last_Name = $lname
	        $exceptionsout += $lineitem
	        $lineitem = $null
	    }
	    If ($deptnum -eq 62140){
	        LogWrite "Adding User to toltsalesteam distro"
			Add-ADGroupMember -Identity toltsalesteam -Members $exists
	    }
	    If ($deptnum -eq 62142){
	        LogWrite "Adding User to toltsalesteam distro"
			Add-ADGroupMember -Identity toltsalesteam -Members $exists
	    }
	    If ($deptnum -eq 62143){
	        LogWrite "Adding User to toltsalesteam distro"
			Add-ADGroupMember -Identity toltsalesteam -Members $exists
	    }
		If ($deptnum -eq 62145){
	        LogWrite "Adding User to toltsalesteam distro"
			Add-ADGroupMember -Identity toltsalesteam -Members $exists
		}
	    If ($deptnum -ne 62143 -and $deptnum -ne 62142 -and $deptnum -ne 62140 -and $deptnum -ne 62145 ){
	        LogWrite "Removing User from toltsalesteam distro"
			Remove-ADGroupMember -Identity toltsalesteam -Members $exists -Confirm:$false
	    }

	    $exists = $null
	    $manager = $null
	    $dept = $null
		LogWrite "Processing Complete ##########################"
	}
}
$exceptionsout | export-csv $exceptionsfile
$exceptionsmgr | export-csv $exceptionsmgrfile
LogWrite "******************End Oracle Processing $(get-date -format `"yyyy-MM-dd_hh:mmtt`")******************"
exit