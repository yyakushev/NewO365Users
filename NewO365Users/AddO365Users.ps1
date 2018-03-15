#
# Script.ps1
#
[cmdletbinding(SupportsShouldProcess=$True)]

Param(
	[Parameter(Mandatory=$true)]
	[ValidateScript({Test-Path $_ -PathType Leaf})]
	[string] $CSVPath,

	[Parameter(Mandatory=$false)]
	[string[]]$Groups,

	[Parameter(Mandatory=$false)]
	[string]$License,

	[Parameter(Mandatory=$false)]
	[string]$LogPath = $PWD
)

$LogName = $MyInvocation.ScriptName+".log"

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
			"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$logPath\$LogName" -Append
		}
	} catch {
		"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$PWD\$LogName" -Append
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
			"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$logPath\$LogName" -Append
		}
	} catch {
		"$(get-date -Format "HH:mm:ss dd-MM-yyyy: ")$msg" | Out-File -FilePath "$PWD\$LogName" -Append 
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
$Users = Import-Csv $CSVPath 

$script:ErrorActionPreference = "SilentlyContinue"
#Create new users
foreach ($user in $Users) {
	try {
		if (-not (Get-MsolUser -UserPrincipalName $user.UserPrincipalName)) {
			if ($WhatIfPreference) {
				"What if: Performing the operation `"Create new user $($user.UserPrincipalName) in Office 365`""
			} else { 
				$NewO365Users += New-MsolUser -DisplayName $user.DisplayName `
						-UserPrincipalName $user.UserPrincipalName `
						-FirstName $user.DisplayName.split()[0] `
						-LastName $user.DisplayName.split()[1] `
						-Department $user.BusinessUnit `
						-City $user.City `
						-Country $user.Country 
				Write-InformationEventLog -msg "New user $($user.UserPrincipalName) has been created in Office 365" -LogPath $LogPath
			}
		} else {
			Write-InformationEventLog -msg "The user $($user.UserPrincipalName) is already exist in Office 365" -LogPath $LogPath
		}
	} catch {
		Write-ErrorEventLog -msg "User $($user.UserPrincipalName) cannot be created. Please check the error: $($error[0].ToString())" -LogPath $LogPath
	}
}
#Export new users to the CSV file
try {
	$NewO365Users | Export-Csv "CreatedO365Users-$(get-date -Format "dd-MM-yyyy_HH-mm-ss").csv" -Encoding UTF8
	Write-InformationEventLog -msg "Created users list has been exported to CSV" -LogPath $LogPath
} catch {
	Write-ErrorEventLog -msg "Export to CSV failed. Please check the error: $($error[0].ToString())" -LogPath $LogPath
}

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
				"What if: Performing the operation `"Add user $($user.UserPrincipalName) to the group `"$group`"`""
			} else {
				Add-MsolGroupMember -GroupObjectId $GroupID -GroupMemberType User -GroupMemberObjectId $userID
				Write-InformationEventLog -msg "User $($user.UserPrincipalName) has been added to group `"$group`" is $GroupID" -LogPath $LogPath 
			}
		}
	} catch {
		Write-ErrorEventLog -msg "Group `"$group`" cannot be processed. Please check the error: $($error[0].ToString())" -LogPath $LogPath
	}
}

if ($License) {
	foreach ($user in $Users) {
		try {
			if (-not (Get-MsolUser -UserPrincipalName $user.userprincipalname).UsageLocation) {
				if ($WhatIfPreference) {
					"What if: Performing the operation `"Set usage location to DE for the user $($user.UserPrincipalName)`""
				} else {
					Set-MsolUser -UserPrincipalName $user.userprincipalname -UsageLocation DE
					Write-InformationEventLog -msg "Set usage location to DE for the user $($user.UserPrincipalName)" -LogPath $LogPath 
				}
			}
			if ($user.AddLicense) {
				if ($WhatIfPreference) {
					"What if: Performing the operation `"Add license `"$License`" to the user $($user.UserPrincipalName)`""
				} else {
					Set-MsolUserLicense -UserPrincipalName $user.userprincipalname -AddLicenses $License
					Write-InformationEventLog -msg "Add license `"$License`" to the user $($user.UserPrincipalName)" -LogPath $LogPath 
				}
			}
		} catch {
			Write-ErrorEventLog -msg "License cannot be assigned to the user $($user.UserPrincipalName). Please check the error: $($error[0].ToString())" -LogPath $LogPath
		}
	}	
}