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
	
	Creates two text files, by default, in the folder where the script is run.
	
	Optionally, can specify the output folder.
	
	The script has been tested with PowerShell versions 3, 4, 5, and 5.1.
	The script has been tested with Microsoft Windows Server 2008 R2 (with PowerShell V3), 
	2012, 2012 R2, 2016, 2019 and Windows 10.
.PARAMETER Name
	Specifies the Name of the target computer.
	
	Accepts input from the pipeline.
.PARAMETER Folder
	Specifies the optional output folder to save the output reports. 
.EXAMPLE
	Get-ADComputer -Filter * | .\Get-ServiceAccounts.ps1

.EXAMPLE
	Get-AdComputer -filter {OperatingSystem -like "*window*"} | 
	.\Get-ServiceAccounts.ps1 -Folder \\FileServer\ShareName
	
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	Get-AdComputer -filter {OperatingSystem -like "*window*"} 
	-SearchBase "ou=SQLServers,dc=domain,dc=tld" 
	-SearchScope Subtree 
	-properties Name -EA 0 | 
	Sort Name | 
	.\Get-ServiceAccounts.ps1
.EXAMPLE
	Get-AdComputer -filter {OperatingSystem -like "*window*"} 
	-properties Name -EA 0 | Sort Name | .\Get-ServiceAccounts.ps1
	
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

.INPUTS
	Accepts pipeline input with the property Name or a list of computer names.
.OUTPUTS
	No objects are output from this script.  This script creates two texts files.
.NOTES
	NAME: Get-ServiceAccounts.ps1
	VERSION: 1.00
	AUTHOR: Carl Webster and Michael B. Smith
	LASTEDIT: December 19, 2019
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
	[string]$Folder=""
	
	)

Begin
{
    Set-StrictMode -Version Latest

	Function ProcessComputer
	{
		Param
		(
			[String] $Name
		)

		$Computer = $Name.Trim()
		Write-Host "Testing computer $($Computer)"

		$TestResult = Test-NetConnection -ComputerName $Computer -Port 139 -EA 0

		If(($TestResult.PingSucceeded -eq $true) -or ($TestResult.PingSucceeded -eq $False -and $TestResult.TcpTestSucceeded -eq $True))
		{
			If($TestResult.TcpTestSucceeded)
			{
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
					Write-Host "`tFound a match"
					$Script:AllMatches += $Results
				}
                Else
                {
					Write-Host "`tNo services using domain credentials were found"
                }
			}
			Else
			{
				Write-Host "`tComputer $($Computer) is online but the test for TCP Port 139 (File and Print Sharing) failed"
				Out-File -FilePath $Filename2 -Append `
					-InputObject "Computer $($Computer) is online but the test for TCP Port 139 (File and Print Sharing) failed"
			}
		}
		Else
		{
			If($TestResult.PingSucceeded -eq $False -and $Null -eq $TestResult.RemoteAddress)
			{
				Write-Host "`tComputer $($Computer) was not found in DNS $(Get-Date)"
				Out-File -FilePath $Filename2 -Append `
					-InputObject "Computer $($Computer) was not found in DNS $(Get-Date)"
			}
			Else
			{
				Write-Host "`tComputer $($Computer) is not online or is online but is not a Windows computer"
				Out-File -FilePath $Filename2 -Append `
					-InputObject "Computer $($Computer) was not online $(Get-Date) or is online but is not a Windows computer"
			}
			
		}
	}

    Write-Host "$(Get-Date): Setting up script"

    If($Folder -ne "")
    {
	    Write-Host "$(Get-Date): Testing folder path"
	    #does it exist
	    If(Test-Path $Folder -EA 0)
	    {
		    #it exists, now check to see if it is a folder and not a file
		    If(Test-Path $Folder -pathType Container -EA 0)
		    {
			    #it exists and it is a folder
			    Write-Host "$(Get-Date): Folder path $Folder exists and is a folder"
		    }
		    Else
		    {
			    #it exists but it is a file not a folder
			    Write-Error "Folder $Folder is a file, not a folder. Script cannot continue"
			    Exit
		    }
	    }
	    Else
	    {
		    #does not exist
		    Write-Error "Folder $Folder does not exist.  Script cannot continue"
		    Exit
	    }
    }

    If($Folder -eq "")
    {
	    $pwdpath = $pwd.Path
    }
    Else
    {
	    $pwdpath = $Folder
    }

    [string]$Script:FileName = Join-Path $pwdpath "ComputersWithDomainServiceAccounts.txt"
    [string]$Script:FileName2 = Join-Path $pwdpath "ComputersWithDomainServiceAccountsErrors.txt"

    $startTime = Get-Date

    $Script:AllMatches = @()
}

Process
{
    If($Name -is [array])
    {
        ForEach($Computer in $Name)
        {
			ProcessComputer $Computer
        }
    }
    Else
    {
		ProcessComputer $Name
    }
}

End
{
    $Script:AllMatches = $Script:AllMatches | Sort-Object SystemName,Name

    $Script:AllMatches | Out-String -width 200 | Out-File -FilePath $Script:FileName

    If(Test-Path "$($Script:FileName)")
    {
	    Write-Host "$(Get-Date): $($Script:FileName) is ready for use"
    }
    If(Test-Path "$($Script:FileName2)")
    {
	    Write-Host "$(Get-Date): $($Script:FileName2) is ready for use"
    }

    Write-Host "$(Get-Date): Script started: $($StartTime)"
    Write-Host "$(Get-Date): Script ended: $(Get-Date)"
    $runtime = $(Get-Date) - $StartTime
    $Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	    $runtime.Days, `
	    $runtime.Hours, `
	    $runtime.Minutes, `
	    $runtime.Seconds,
	    $runtime.Milliseconds)
    Write-Host "$(Get-Date): Elapsed time: $($Str)"
    $runtime = $Null

	Write-Host "                                                                                    " -BackgroundColor Black -ForegroundColor White
	Write-Host "               This FREE script was brought to you by Conversant Group              " -BackgroundColor Black -ForegroundColor White
	Write-Host "We design, build, and manage infrastructure for a secure, dependable user experience" -BackgroundColor Black -ForegroundColor White
	Write-Host "                       Visit our website conversantgroup.com                        " -BackgroundColor Black -ForegroundColor White
	Write-Host "                                                                                    " -BackgroundColor Black -ForegroundColor White
}
