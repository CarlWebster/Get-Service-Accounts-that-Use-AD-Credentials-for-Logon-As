<#
.SYNOPSIS
	Find Services using a domain account on specified computers in Microsoft Active 
	Directory.
.DESCRIPTION
	Find Services using a domain account on specified computers in Microsoft Active 
	Directory.
	
	Process each computer looking for Services using a domain account for Log On As.
	
	Builds a list of computer names, Service names, service display names, and service start 
	names.
	
	Creates two text files and one CSV file, by default, in the folder where the script 
	is run.
	
	Optionally, can specify the output folder.
	
	The script has been tested with PowerShell versions 3, 4, 5, and 5.1.
	The script has been tested with Microsoft Windows Server 2008 R2 (with PowerShell V3), 
	2012, 2012 R2, 2016, 2019, and Windows 10.
.PARAMETER Name
	Specifies the Name of the target computer.
	
	Accepts input from the pipeline.
.PARAMETER UseDcom
	The script, by default, uses TCP Port 139 for testing if a computer is online and 
	available.
	Use the UseDcom switch to have the script use TCP Port 5985 to use WinRM 2.0.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output reports. 
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	The default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	The default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.EXAMPLE
	Get-ADComputer -Filter * | Select Name | Sort Name | .\Get-ServiceAccounts.ps1

.EXAMPLE
	$Names = Get-ADComputer -Filter * | Select Name | Sort Name
	.\Get-ServiceAccounts.ps1 -Name $Names
	
.EXAMPLE
	"MgmtComputer" | .\Get-ServiceAccounts.ps1
	
	Finds service accounts on the one specified computer.
.EXAMPLE
	Get-ADComputer -Filter * | Select Name | Sort Name | .\Get-ServiceAccounts.ps1 
	-UseDcom
	
	Use TCP port 5985 instead of TCP port 139.
.EXAMPLE
	$Names = Get-ADComputer -Filter * | Select Name | Sort Name
	.\Get-ServiceAccounts.ps1 -Name $Names -UseDcom

	Use TCP port 5985 instead of TCP port 139.
.EXAMPLE
	Get-AdComputer -filter {OperatingSystem -like "*window*"} | Sort Name | 
	.\Get-ServiceAccounts.ps1 -Folder \\FileServer\ShareName
	
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	Get-AdComputer -filter {OperatingSystem -like "*window*"} -SearchBase 
	"ou=SQLServers,dc=domain,dc=tld" -SearchScope Subtree -properties Name -EA 0 | Sort 
	Name | .\Get-ServiceAccounts.ps1
.EXAMPLE
	Get-AdComputer -filter {OperatingSystem -like "*window*"} -properties Name -EA 
	0 | Sort Name | .\Get-ServiceAccounts.ps1
	
	Processes only computers with "window" in the OperatingSystem property
.EXAMPLE
	Get-AdComputer -filter {OperatingSystem -like "*window*" -and OperatingSystem 
	-like "*server*"} -properties Name -EA 0 | Sort Name | .\Get-ServiceAccounts.ps1
	
	Processes only computers with "window" and "server" in the OperatingSystem property.
	This catches operating systems like Windows 2000 Server and Windows Server 2003.
.EXAMPLE
	Get-Content "C:\webster\computernames.txt" | .\Get-ServiceAccounts.ps1
	
	computernames.txt is a plain text file that contains a list of computer names.
	
	For example:
	
	LABCA
	LABDC1
	LABDC2
	LABFS
	LABIGEL
	LABMGMTPC
	LABSQL1

.EXAMPLE
	Get-ADComputer -Filter * | .\Get-ServiceAccounts.ps1 -SmtpServer 
	mail.domain.tld -From XDAdmin@domain.tld -To ITGroup@domain.tld	

	The script uses the email server mail.domain.tld, sending from XDAdmin@domain.tld and 
	sending to ITGroup@domain.tld.

	The script uses the default SMTP port 25 and does not use SSL.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.
.EXAMPLE
	Get-ADComputer -Filter * | .\Get-ServiceAccounts.ps1 -SmtpServer 
	mailrelay.domain.tld -From Anonymous@domain.tld -To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script uses the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld and sending to ITGroup@domain.tld.

	To send an unauthenticated email using an email relay server requires the From email 
	account use the name Anonymous.

	The script uses the default SMTP port 25 and does not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send an email using a Gmail or g-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script generates an anonymous, secure password for the anonymous@domain.tld 
	account.
.EXAMPLE
	Get-ADComputer -Filter * | .\Get-ServiceAccounts.ps1 -SmtpServer 
	labaddomain-com.mail.protection.outlook.com -UseSSL -From 
	SomeEmailAddress@labaddomain.com -To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script uses the email server labaddomain-com.mail.protection.outlook.com, sending 
	from SomeEmailAddress@labaddomain.com and sending to ITGroupDL@labaddomain.com.

	The script uses the default SMTP port 25 and SSL.
.EXAMPLE
	Get-ADComputer -Filter * | .\Get-ServiceAccounts.ps1 -SmtpServer 
	smtp.office365.com -SmtpPort 587 -UseSSL -From Webster@CarlWebster.com -To 
	ITGroup@CarlWebster.com	

	The script uses the email server smtp.office365.com on port 587 using SSL, sending from 
	webster@carlwebster.com and sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.
.EXAMPLE
	Get-ADComputer -Filter * | .\Get-ServiceAccounts.ps1 -SmtpServer smtp.gmail.com 
	-SmtpPort 587 -UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send an email using a Gmail or g-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	*** NOTE ***
	
	The script uses the email server smtp.gmail.com on port 587 using SSL, sending from 
	webster@gmail.com and sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.
.INPUTS
	Accepts pipeline input with the property Name or a list of computer names.
.OUTPUTS
	No objects are output from this script.  This script creates two text files and a CSV
	file.
.NOTES
	NAME: Get-ServiceAccounts.ps1
	VERSION: 1.20
	AUTHOR: Carl Webster and Michael B. Smith
	LASTEDIT: January 19, 2021
#>


#region script change log	
#Created by Carl Webster and Michael B. Smith
#webster@carlwebster.com
#@carlwebster on Twitter
#https://www.CarlWebster.com
#
#michael@smithcons.com
#@essentialexch on Twitter
#https://www.essential.exchange/blog/
#
#Created on October 31, 2019
#Version 1.00 released to the community on 19-Dec-2019
#
#Version 1.20 19-Jan-2021
#	Added creating a CSV file
#		Changed to creating an object instead of text
#	Added switch UseDcom to change from using TCP port 139 for testing connections to port 5985 (WinRM 2.0)
#		Updated messages to reflect that change
#	Added the domain name to all output files and the script title
#	Change to using Invoke-Command to gather the service information on remote computers when -UseDcom is used
#		If testing the computer running the script, don't use Invoke-Command
#	Reordered the parameters in an order recommended by Guy Leech
#	Updated help text
#	Updated ReadMe file
#
#Version 1.10 29-Apr-2020
#	Add email capability to match other scripts
#		Update Help Text with examples
#	Add ScriptInfo Parameter
#		Add code to show Script Options and write out Script Info file
#	Change location of the -Dev, -Log, and -ScriptInfo output files from the script folder to the -Folder location (Thanks to Guy Leech for the "suggestion")
#	Cleanup screen output
#	Enable Verbose output
#	If the tested computer is online and no service with domain creds was found, write that to the output file
#	Make sure that filename3 (ComputersWithDomainServiceAccountsErrors.txt) is new for each run
#	Reformatted the terminating Write-Error messages to make them more visible and readable in the console
#endregion


[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "") ]

Param(
	[parameter(
		Mandatory                       = $True,
		ValueFromPipeline               = $True,
		ValueFromPipelineByPropertyName = $True,
		Position                        = 0)] 
	[string[]]$Name,
	
	[parameter(Mandatory=$False)] 
	[Switch]$UseDcom=$False,
	
	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$SmtpServer="",

	[parameter(Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(Mandatory=$False)] 
	[string]$From="",

	[parameter(Mandatory=$False)] 
	[string]$To=""

	)

Begin
{
    Set-StrictMode -Version Latest
	$PSDefaultParameterValues = @{"*:Verbose"=$True}
	
	If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
	{
		Write-Error "
		`n`n
		`t`t
		You specified an SmtpServer but did not include a From or To email address.
		`n`n
		`t`t
		Script cannot continue.
		`n`n"
		Exit
	}
	If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
	{
		Write-Error "
		`n`n
		`t`t
		You specified an SmtpServer and a To email address but did not include a From email address.
		`n`n
		`t`t
		Script cannot continue.
		`n`n"
		Exit
	}
	If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
	{
		Write-Error "
		`n`n
		`t`t
		You specified an SmtpServer and a From email address but did not include a To email address.
		`n`n
		`t`t
		Script cannot continue.
		`n`n"
		Exit
	}
	If(![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
	{
		Write-Error "
		`n`n
		`t`t
		You specified From and To email addresses but did not include the SmtpServer.
		`n`n
		`t`t
		Script cannot continue.
		`n`n"
		Exit
	}
	If(![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
	{
		Write-Error "
		`n`n
		`t`t
		You specified a From email address but did not include the SmtpServer.
		`n`n
		`t`t
		Script cannot continue.
		`n`n"
		Exit
	}
	If(![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
	{
		Write-Error "
		`n`n
		`t`t
		You specified a To email address but did not include the SmtpServer.
		`n`n
		`t`t
		Script cannot continue.
		`n`n"
		Exit
	}
    Write-Verbose "$(Get-Date): Setting up script"

    If($Folder -ne "")
    {
	    Write-Verbose "$(Get-Date): Testing folder path"
	    #does it exist
	    If(Test-Path $Folder -EA 0)
	    {
		    #it exists, now check to see if it is a folder and not a file
		    If(Test-Path $Folder -pathType Container -EA 0)
		    {
			    #it exists and it is a folder
			    Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		    }
		    Else
		    {
			    #it exists but it is a file not a folder
			    Write-Error "
				`n`n
				`t`t
				Folder $Folder is a file, not a folder.
				`n`n
				`t`t
				Script cannot continue.
				`n`n
				"
			    Exit
		    }
	    }
	    Else
	    {
		    #does not exist
		    Write-Error "
			`n`n
			`t`t
			Folder $Folder does not exist.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
		    Exit
	    }
    }

    If($Folder -eq "")
    {
	    $Script:pwdpath = $pwd.Path
    }
    Else
    {
	    $Script:pwdpath = $Folder
    }

	If($Script:pwdpath.EndsWith("\"))
	{
		#remove the trailing \
		$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
	}

	If($Log) 
	{
		#start transcript logging
		$Script:LogPath = "$($Script:pwdpath)\ComputersWithDomainServiceAccountsScriptTranscript_ForDomain_$($Domain)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		
		try 
		{
			Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
			Write-Verbose "$(Get-Date): Transcript/log started at $Script:LogPath"
			$Script:StartLog = $true
		} 
		catch 
		{
			Write-Verbose "$(Get-Date): Transcript/log failed at $Script:LogPath"
			$Script:StartLog = $false
		}
	}

	If($Dev)
	{
		$Error.Clear()
		$Script:DevErrorFile = "$($Script:pwdpath)\ComputersWithDomainServiceAccountsScriptErrors_ForDomain_$($Domain)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	$Domain = $Env:UserDNSDomain
    [string]$Script:FileName1 = "$($Script:pwdpath)\ComputersWithDomainServiceAccounts_ForDomain_$($Domain).txt"
    [string]$Script:FileName2 = "$($Script:pwdpath)\ComputersWithDomainServiceAccounts_ForDomain_$($Domain).csv"
    [string]$Script:filename3 = "$($Script:pwdpath)\ComputersWithDomainServiceAccountsErrors_ForDomain_$($Domain).txt"
	[string]$Script:Title = "Computers with Domain Service Accounts For Domain $($Domain)"
	[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

    $startTime = Get-Date

	#make sure the error file is new
	Out-File -FilePath $Script:filename3 -InputObject "" -EA 0 4>$Null

	$TCPPort = 139
	If($UseDcom)
	{
		$TCPPort = 5985
	}

	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Dev                : $($Dev)"
	Write-Verbose "$(Get-Date): Filename1          : $($Script:filename1)"
	Write-Verbose "$(Get-Date): Filename2          : $($Script:filename2)"
	Write-Verbose "$(Get-Date): Filename3          : $($Script:filename3)"
	Write-Verbose "$(Get-Date): Folder             : $($Script:pwdpath)"
	Write-Verbose "$(Get-Date): Log                : $($Log)"
	Write-Verbose "$(Get-Date): ScriptInfo         : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): TCP Port           : $($TCPPort)"
	Write-Verbose "$(Get-Date): Title              : $($Script:Title)"
	Write-Verbose "$(Get-Date): Smtp Port          : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server        : $($SmtpServer)"
	Write-Verbose "$(Get-Date): From               : $($From)"
	Write-Verbose "$(Get-Date): To                 : $($To)"
	Write-Verbose "$(Get-Date): Use SSL            : $($UseSSL)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected        : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version       : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture          : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture        : $($PSUICulture)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start       : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "

	#region email function
	Function SendEmail
	{
		Param([array]$Attachments)
		Write-Verbose "$(Get-Date): Prepare to email"

		$emailAttachment = $Attachments
		$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.

"@ 

		If($Dev)
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}

		$error.Clear()
		
		If($From -Like "anonymous@*")
		{
			#https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
			$anonUsername = "anonymous"
			$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
			$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

			If($UseSSL)
			{
				Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
				-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
				-UseSSL -credential $anonCredentials *>$Null 
			}
			Else
			{
				Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
				-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
				-credential $anonCredentials *>$Null 
			}
			
			If($?)
			{
				Write-Verbose "$(Get-Date): Email successfully sent using anonymous credentials"
			}
			ElseIf(!$?)
			{
				$e = $error[0]

				Write-Verbose "$(Get-Date): Email was not sent:"
				Write-Warning "$(Get-Date): Exception: $e.Exception" 
			}
		}
		Else
		{
			If($UseSSL)
			{
				Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
				Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
				-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
				-UseSSL *>$Null
			}
			Else
			{
				Write-Verbose "$(Get-Date): Trying to send email using current user's credentials without SSL"
				Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
				-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
			}

			If(!$?)
			{
				$e = $error[0]
				
				#error 5.7.57 is O365 and error 5.7.0 is gmail
				If($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7"))
				{
					#The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
					Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

					If($Dev)
					{
						Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
					}

					$error.Clear()

					$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

					If($UseSSL)
					{
						Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
						-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
						-UseSSL -credential $emailCredentials *>$Null 
					}
					Else
					{
						Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
						-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
						-credential $emailCredentials *>$Null 
					}

					If($?)
					{
						Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
					}
					ElseIf(!$?)
					{
						$e = $error[0]

						Write-Verbose "$(Get-Date): Email was not sent:"
						Write-Warning "$(Get-Date): Exception: $e.Exception" 
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): Email was not sent:"
					Write-Warning "$(Get-Date): Exception: $e.Exception" 
				}
			}
		}
	}
	#endregion

	Function ProcessComputer
	{
		Param
		(
			[String] $Name
		)

		$Computer = $Name.Trim()
		Write-Verbose "$(Get-Date): Testing computer $($Computer)"
		
		$TestResult = Test-NetConnection -ComputerName $Computer -Port $TCPPort -EA 0 3>$Null 4>$Null

		If(($TestResult.PingSucceeded -eq $true) -or ($TestResult.PingSucceeded -eq $False -and $TestResult.TcpTestSucceeded -eq $True))
		{
			If($TestResult.TcpTestSucceeded)
			{
				If($Env:ComputerName -eq $Computer)
				{
					#process the local computer. Can't use invoke-command on self
					$Results = Get-WmiObject -ComputerName $Computer Win32_Service -EA 0 | 
					Where-Object {
						$_.ServiceType -ne "Unknown" -And 
						$_.StartName -NotLike ".\*" -And 
						$_.StartName -NotLike "LocalSystem" -And 
						$_.StartName -NotLike "LocalService*" -And 
						$_.StartName -NotLike "NT AUTHORITY*" -And 
						$_.StartName -NotLike "NT SERVICE*"} | 
					Select-Object SystemName, Name, DisplayName, StartName
			
					If($? -and $Null -ne $Results)
					{
						Write-Verbose "$(Get-Date): `tFound a match"
						ForEach($Result in $Results)
						{
							$obj1 = [PSCustomObject] @{
								SystemName  = $Result.SystemName
								ServiceName = $Result.Name
								DisplayName = $Result.DisplayName
								StartName   = $Result.StartName
							}
							$null = $Script:AllMatches.Add($obj1)
						}
					}
					Else
					{
						Write-Verbose "$(Get-Date): `tNo services using domain credentials were found on computer $($Computer)"
						$obj1 = [PSCustomObject] @{
							SystemName  = $Computer
							ServiceName = "N/A"
							DisplayName = "No services using domain credentials were found"
							StartName   = "N/A"
						}
						$null = $Script:AllMatches.Add($obj1)
					}
				}
				Else
				{
					If( $UseDcom )
					{
						$Results = Invoke-Command -ComputerName $Computer -ScriptBlock {
							Get-WmiObject Win32_Service -EA 0 | 
								Where-Object {
									$_.ServiceType -ne "Unknown" -And 
									$_.StartName -NotLike ".\*" -And 
									$_.StartName -NotLike "LocalSystem" -And 
									$_.StartName -NotLike "LocalService*" -And 
									$_.StartName -NotLike "NT AUTHORITY*" -And 
									$_.StartName -NotLike "NT SERVICE*"} | 
								Select-Object SystemName, Name, DisplayName, StartName
						}
					}
					Else
					{
						$Results = Get-WmiObject Win32_Service -ComputerName $Computer -EA 0 | 
							Where-Object {
								$_.ServiceType -ne "Unknown" -And 
								$_.StartName -NotLike ".\*" -And 
								$_.StartName -NotLike "LocalSystem" -And 
								$_.StartName -NotLike "LocalService*" -And 
								$_.StartName -NotLike "NT AUTHORITY*" -And 
								$_.StartName -NotLike "NT SERVICE*"} | 
							Select-Object SystemName, Name, DisplayName, StartName
					}
			
					If($? -and $Null -ne $Results)
					{
						Write-Verbose "$(Get-Date): `tFound a match"
						ForEach($Result in $Results)
						{
							$obj1 = [PSCustomObject] @{
								SystemName  = $Result.SystemName
								ServiceName = $Result.Name
								DisplayName = $Result.DisplayName
								StartName   = $Result.StartName
							}
							$null = $Script:AllMatches.Add($obj1)
						}
					}
					Else
					{
						Write-Verbose "$(Get-Date): `tNo services using domain credentials were found on computer $($Computer)"
						$obj1 = [PSCustomObject] @{
							SystemName  = $Computer
							ServiceName = "N/A"
							DisplayName = "No services using domain credentials were found"
							StartName   = "N/A"
						}
						$null = $Script:AllMatches.Add($obj1)
					}
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date): `tComputer $($Computer) is online but the test for TCP Port $TCPPort failed"
				Out-File -FilePath $Script:filename3 -Append `
					-InputObject "Computer $($Computer) is online but the test for TCP Port $TCPPort failed" 4>$Null
			}
		}
		Else
		{
			If($TestResult.PingSucceeded -eq $False -and $Null -eq $TestResult.RemoteAddress)
			{
				Write-Verbose "$(Get-Date): `tComputer $($Computer) was not found in DNS $(Get-Date)"
				Out-File -FilePath $Script:filename3 -Append `
					-InputObject "Computer $($Computer) was not found in DNS $(Get-Date)" 4>$Null
			}
			Else
			{
				Write-Verbose "$(Get-Date): `tComputer $($Computer) is not online or is online but is not a Windows computer"
				Out-File -FilePath $Script:filename3 -Append `
					-InputObject "Computer $($Computer) was not online $(Get-Date) or is online but is not a Windows computer" 4>$Null
			}
			
		}
	}

    $Script:AllMatches = New-Object System.Collections.ArrayList
}

Process
{
	ForEach($Computer in $Name)
	{
		#$Computer value is @{Name=ADComputerName}
		$StartPos = $Computer.IndexOf("=")
		If( $StartPos -ge 0 )
		{
			$EndPos = $Computer.IndexOf("}")
			$ComputerName = $Computer.Substring($StartPos+1,$EndPos-$StartPos-1)
			ProcessComputer $ComputerName
		}
		Else
		{
			ProcessComputer $Computer
		}
	}
}

End
{
    $Script:AllMatches = $Script:AllMatches | Sort-Object SystemName,Name

    $Script:AllMatches | Out-String -width 200 | Out-File -FilePath $Script:FileName1 -EA 0 4>$Null
	$Script:AllMatches | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $Script:FileName2

	$emailAttachment = @()
    If(Test-Path "$($Script:FileName1)")
    {
	    Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
		#email output file if requested
		If(![System.String]::IsNullOrEmpty( $SmtpServer ))
		{
			$emailAttachment += $Script:FileName1
		}
	}
    If(Test-Path "$($Script:FileName2)")
    {
	    Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
		#email output file if requested
		If(![System.String]::IsNullOrEmpty( $SmtpServer ))
		{
			$emailAttachment += $Script:FileName2
		}
	}
    If(Test-Path "$($Script:filename3)")
    {
	    Write-Verbose "$(Get-Date): $($Script:filename3) is ready for use"
		#email output file if requested
		If(![System.String]::IsNullOrEmpty( $SmtpServer ))
		{
			$emailAttachment += $Script:filename3
		}
    }

	If(![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		SendEmail $emailAttachment
	}
	
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

    Write-Verbose "$(Get-Date): Script started: $($StartTime)"
    Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
    $runtime = $(Get-Date) - $StartTime
    $Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	    $runtime.Days, `
	    $runtime.Hours, `
	    $runtime.Minutes, `
	    $runtime.Seconds,
	    $runtime.Milliseconds)
    Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$Script:pwdpath\ComputersWithDomainServiceAccountsScriptInfo_ForDomain_$($Domain)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev                : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile       : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Filename1          : $($Script:filename1)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Filename2          : $($Script:filename2)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Filename3          : $($Script:filename3)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Folder             : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                : $($Log)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info        : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "TCP Port           : $($TCPPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title              : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port          : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server        : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From               : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                 : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL            : $($UseSSL)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected        : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version       : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture          : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture        : $($PSUICulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start       : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time       : $($Str)" 4>$Null
	}

	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date): Transcript/log stop failed"
			}
		}
	}

	$runtime = $Null
	$Str = $Null
}
