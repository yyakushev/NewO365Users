[cmdletbinding(SupportsShouldProcess=$True)]

Param(
	[Parameter(Mandatory=$true)]
	[ValidateScript({Test-Path $_ -PathType Leaf})]
	[string] $CSVPath,

	[Parameter(Mandatory=$false)]
	[char] $delimiter = ';',

	[Parameter(Mandatory=$false)]
	[string]$LogPath = $PWD
)

$logName = "ChangePassword.log"

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

function Get-RandomLowercase {[char] (get-random -Minimum 97 -Maximum 122 )}
function Get-RandomUppercase {[char] (get-random -Minimum 65 -Maximum 90 )}
function Get-RandomSpecial {[char] (get-random -InputObject '!','@','#','$','%','&' )}
function Get-RandomNumber {get-random -Minimum 0 -Maximum 9}

ForEach ($user in $Users) {
    try {
		if ($WhatIfPreference) {
			"What if: Performing the operation `"Create new password for $($user.UserPrincipalName)`""
		} else {
			$User.NewPassword = "Lanxess$(Get-RandomSpecial)$(Get-RandomLowercase)$(Get-RandomUppercase)$(Get-RandomNumber)$(Get-RandomNumber)$(Get-RandomNumber)"
			Set-MsolUserPassword -UserPrincipalName $user.UserPrincipalName -NewPassword $User.NewPassword
			Write-InformationEventLog -msg "Password has been reseted for user $($user.UPN)" -LogPath $LogPath
		}
    } catch {
        Write-InformationEventLog -msg "ERROR: Password has not been reseted for user $($user.UserPrincipalName)" -LogPath $LogPath
    }
}

try {
	$Users |select UserPrincipalName, NewPassword | Export-Csv "NewPasswords-$(get-date -Format "dd-MM-yyyy_HH-mm-ss").csv" -Encoding UTF8 -Delimiter $delimiter
	Write-InformationEventLog -msg "Created users list has been exported to CSV" -LogPath $LogPath
} catch {
	Write-ErrorEventLog -msg "Export to CSV failed. Please check the error: $($error[0].ToString())" -LogPath $LogPath
}

