#
# RemoveO365GroupMember.ps1
#
[cmdletbinding(SupportsShouldProcess=$True)]

Param(
	[Parameter(Mandatory=$true)]
	[ValidateScript({Test-Path $_ -PathType Leaf})]
	[string] $CSVPath,

	[Parameter(Mandatory=$false)]
	[char] $delimiter = ';',

	[Parameter(Mandatory=$false)]
	[string[]]$Groups,

	[Parameter(Mandatory=$false)]
	[string]$LogPath = $PWD
)

$LogName = "RemoveO365GroupMember.log"

#Set the error action preference for the script
$script:ErrorActionPreference = "Stop"

function Write-ErrorEventLog ([string] $msg, [string] $LogPath) {
	$WhatIfPreferenceVariable = $WhatIfPreference
	$WhatIfPreference = $false
	$t = $host.ui.RawUI.ForegroundColor
	$host.ui.RawUI.ForegroundColor = "Red"
	Write-Output $msg
	$host.ui.RawUI.ForegroundColor = $t
	try {
		if  (Test-Path -Path $logpath -ErrorAction SilentlyContinue ) {
			"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$LogPath\$LogName" -Append
		}
	} catch {
		"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$LogPath\$LogName" -Append
	}
	$WhatIfPreference = $WhatIfPreferenceVariable
} #function write error 

function Write-InformationEventLog ([string] $msg, [string] $LogPath) {
	$WhatIfPreferenceVariable = $WhatIfPreference
	$WhatIfPreference = $false
	$t = $host.ui.RawUI.ForegroundColor
	$host.ui.RawUI.ForegroundColor = "Yellow"
	Write-Output $msg
	$host.ui.RawUI.ForegroundColor = $t
	try {
		if  (Test-Path -Path $logpath -ErrorAction SilentlyContinue ) {
			"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$LogPath\$LogName" -Append
		}
	} catch {
		"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$LogPath\$LogName" -Append 
	}
	$WhatIfPreference = $WhatIfPreferenceVariable
} #function write information


#Check if module MSOnline is installed
if (Get-Module -Name MSOnline -ListAvailable) {
	import-module MSOnline
	Write-InformationEventLog -msg "MSOnline module has been successfully loaded" -LogPath $LogPath 
} else {
	Write-ErrorEventLog -msg "MSOnline module has not been found. Please install it first.`n`r More informantion you can find here https://docs.microsoft.com/en-us/powershell/azure/active-directory/overview?view=azureadps-1.0" -LogPath $LogPath
	exit
}

$credential = get-credential -Message "Please use account that have administrative credentials in Office 365"

#Connect to Office 365
try {
	Connect-MsolService  -Credential $credential
	$companyinfo = Get-MsolCompanyInformation 
	if ($companyinfo) {
		Write-InformationEventLog -msg "Connection to Office 365 has been established" -LogPath $LogPath
		Write-InformationEventLog -msg "You have connected to the `"$($companyinfo.DisplayName)`" online services" -LogPath $LogPath
	}
} catch {
	Write-ErrorEventLog -msg "Connection to the Office 365 has not been established. `r`nPlease check user name and password" -LogPath $LogPath
	exit
}

#Import users from csv
$Users = Import-Csv $CSVPath -Delimiter $delimiter

$script:ErrorActionPreference = "SilentlyContinue"

#Get objectIDs for groups
foreach ($group in $Groups) {
	try {
		$GroupID = (Get-MsolGroup -All |? DisplayName -match "^$group$").objectid
		if ($GroupID) { 
			Write-InformationEventLog -msg "Objectid for group `"$group`" is $GroupID" -LogPath $LogPath 
		} else {
			Write-ErrorEventLog -msg "Group `"$group`" doesn't exist in Office 365" -LogPath $LogPath
		}
		$GroupIDs += $GroupID
	} catch {
		Write-ErrorEventLog -msg "Group `"$group`" cannot be processed. Please check the error: $($error[0].ToString())" -LogPath $LogPath
	}
}

foreach ($user in $Users) {
	try {
		$UserID = (Get-MsolUser -UserPrincipalName $user.UserPrincipalName).objectid
		foreach ($GroupID in $GroupIDs) {
			if ($WhatIfPreference) {
				"What if: Performing the operation `"Remove user $($user.UserPrincipalName) from the group `"$group`"`""
			} else {
				Remove-MsolGroupMember -GroupObjectId $GroupID -GroupMemberType User -GroupMemberObjectId $userID
				Write-InformationEventLog -msg "User $($user.UserPrincipalName) has been removed from the group `"$group`" is $GroupID" -LogPath $LogPath 
			}
		}
	} catch {
		Write-ErrorEventLog -msg "Group `"$group`" cannot be processed. Please check the error: $($error[0].ToString())" -LogPath $LogPath
	}
}