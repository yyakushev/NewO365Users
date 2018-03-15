#
# MFAReport.ps1
#

<#
    .SYNOPSIS 
      Create a "MFA enabled Office 365 users" report 
	.PARAMETERS
    .EXAMPLE
		.\MFAReport.ps1 -ReportPath c:\reports -LogPath c:\Logs	
#>

Param(
  [Parameter(Mandatory=$false)]
   [string]$ReportPath = $PWD,
	
   [Parameter(Mandatory=$false)]
   [string]$LogPath = $PWD
)

$LogName = "MFAReportScript.log"
$ReportName = "MFA_enabled_users_$(get-date -Format "dd-MM-yyyy_HH-mm-ss")"
$script:ErrorActionPreference = "Stop"

function Write-ErrorEventLog ([string] $msg, [string] $LogPath) {
	$t = $host.ui.RawUI.ForegroundColor
	$host.ui.RawUI.ForegroundColor = "Red"
	Write-Output $msg
	$host.ui.RawUI.ForegroundColor = $t
	try {
		if  (Test-Path -Path $logpath -ErrorAction SilentlyContinue ) {
			"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$logPath\$LogName" -Append
		}
	} catch {
		"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$PWD\$LogName" -Append
	}
}

function Write-InformationEventLog ([string] $msg, [string] $LogPath) {
	$t = $host.ui.RawUI.ForegroundColor
	$host.ui.RawUI.ForegroundColor = "Yellow"
	Write-Output $msg
	$host.ui.RawUI.ForegroundColor = $t
	try {
		if  (Test-Path -Path $logpath -ErrorAction SilentlyContinue ) {
			"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$logPath\$LogName" -Append
		}
	} catch {
		"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$PWD\$LogName" -Append
	}
}


if (Get-Module -Name MSOnline -ListAvailable) {
	import-module MSOnline
	Write-InformationEventLog -msg "MSOnline module has been successfully loaded" -LogPath $LogPath
} else {
	Write-ErrorEventLog -msg "MSOnline module has not been found. Please install it first.`n`r More informantion you can find here https://docs.microsoft.com/en-us/powershell/azure/active-directory/overview?view=azureadps-1.0" -LogPath $LogPath
	exit
}

$credential = get-credential -Message "Please use account that have administrative credentials in Office 365"

try {
	Connect-MsolService  -Credential $credential
	$companyinfo = Get-MsolCompanyInformation 
	if ($companyinfo) {
		Write-InformationEventLog -msg "Connection to Office 365 has been esteblished" -LogPath $LogPath
		Write-InformationEventLog -msg "You have connected to the `"$($companyinfo.DisplayName)`" online services" -LogPath $LogPath
	}
} catch {
	Write-ErrorEventLog -msg "Connection to the Office 365 has not been established. `r`nPlease check user name and password" -LogPath $LogPath
	Write-ErrorEventLog -msg "Report could not be created." -LogPath $LogPath
	exit
}

try {
	if (Get-MsolUserRole -UserPrincipalName $credential.UserName | ? {$_.name -like "Company Administrator"}) {
		Write-InformationEventLog -msg "User $($credential.UserName) is a member of `"Company Administrator`" role" -LogPath $LogPath
	} else {
		Write-ErrorEventLog -msg "Please use user that is a member of `"Company Administrator`" role.`r`nReport could not be created." -LogPath $LogPath
		exit
	}
} catch {
	Write-ErrorEventLog -msg "Please use the user account that is a member of `"Company Administrator`" role.`r`nReport could not be created." -LogPath $LogPath
	exit
}

try {
	$MFAUsers = Get-MsolUser -All | sort userprincipalname
	Write-InformationEventLog -msg "The number of users configured for MFA is $($MFAUsers.count)" -LogPath $LogPath
} catch {
	Write-ErrorEventLog -msg "Not possible to get the users information. Please see an error: $($error[0].ToString())" -LogPath $LogPath
	exit
}

try {
	if 	(!(Test-Path -Path $ReportPath -ErrorAction SilentlyContinue )) {
		Write-InformationEventLog -msg "$ReportPath folder doesn't exist. The MFA report will be saved in the current direcotry." -LogPath $LogPath
		$ReportPath = $PWD
	}
	$Header = @"
		<style>
		 TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
		 TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
		 TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
		</style>
"@

	$ReportPS = $MFAUsers | select userprincipalname,`
		@{label="LoginEnabled";expression={$_.BlockCredential}}, `
		@{label="MFA state";expression={if ($_.StrongAuthenticationRequirements.state) {$_.StrongAuthenticationRequirements.state} else {"Disabled"}}},`
		@{label="DefaultMFAMethod";expression={$_.StrongAuthenticationMethods| % {if ($_.IsDefault) {"$($_.methodType)"}}}},`
		@{label="PhoneName";expression={$_.StrongAuthenticationPhoneAppDetails.DeviceName}},`
		@{label="PhoneAuthenticationType";expression={$_.StrongAuthenticationPhoneAppDetails.AuthenticationType}},`
		@{label="PhoneAppVersion";expression={$_.StrongAuthenticationPhoneAppDetails.PhoneAppVersion}} 
	$ReportPS | ConvertTo-Html -Head $Header | Out-File "$ReportPath\$ReportName.html"
	$ReportPS | Export-Csv "$ReportPath\$ReportName.csv"
	Write-InformationEventLog -msg "Please find the `"MFA enabled Office 365 users`" report here $($ReportPath)\$($ReportName).html" -LogPath $LogPath
	. "$ReportPath\$ReportName.html`r`nand CSV version here $ReportPath\$ReportName.csv"
} catch {
	Write-ErrorEventLog -msg "Report could not be builded. Please see an error: $($error[0].ToString())" -LogPath $LogPath
}