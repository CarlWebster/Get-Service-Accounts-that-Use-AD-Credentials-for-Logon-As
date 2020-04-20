# GetServiceAccounts
Get Windows Service Accounts that use domain credentials
	Find Services using a domain account on specified computers in Microsoft Active 
	Directory.
	
	Process each computer looking for Services using a domain account for Log On As.
	
	Builds a list of computer names, Service names, service display names, and service start 
	names.
	
	Creates two text files, by default, in the folder where the script is run.
	
	Optionally, can specify the output folder.
	
	The script has been tested with PowerShell versions 3, 4, 5, and 5.1.
	The script has been tested with Microsoft Windows Server 2008 R2 (with PowerShell V3), 
	2012, 2012 R2, 2016, 2019 and Windows 10.
