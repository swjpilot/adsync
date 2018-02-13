# adsync


This project is a collection of scripts that are designed to be run by 
windows task manager on windows to download a csv file from sharepoint
and process it to update user objects in active directory.

The Run order is as follows:

1.  OracleHeadcountDownload.ps1
	This downloads the file and converts it from a xls filr to a csv 
	file for processing by the following powershell script.
2.  OracleCompare.ps1
	This performs a comparison between the downloaded spreadsheet and
	active directory and then outputs a report in csv format of the 
	changes to be made.
3.  OracleImport.ps1
	This actually performs the updates to active directory.
4.  OracleTermsDownload.ps1
	This downloads the termination report from sharepoint and converts
	it from an xls file to a csv file for processing by the following 
	powershell script.
5.  OracleTerms.ps1
	This processes the Terminations disabling the accounts for the 
	terminated employees.
	