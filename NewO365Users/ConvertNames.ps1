$allusers = Import-Csv .\20171219_ProdUser_Sourcefile.csv -Delimiter ';'
$users = $allusers | Select-Object @{Label='FirstName';Expression = {$_.'first name'}} `
                ,@{Label='LastName';Expression = {$_.'Last name'}} `
                ,@{Label='BusinessUnit';Expression = {$_.'business unit'}} `
                ,@{Label='UserPrincipalName'; `
                    Expression={if (($_.'business unit' -eq 'HPE')-or($_.'business unit' -eq 'TSR')){
                        $_.'user name'.split('\')[1]+'@'+$_.'user name'.split('\')[0].replace('LX-','')+'.arlanxeo.com'
                    } else {
                        $_.'user name'.split('\')[1]+'@'+$_.'user name'.split('\')[0].replace('LX-','')+'.lanxess.com'
                    }}
                } `
| Export-Csv -Path .\LanxessToO365.csv -Delimiter ',' -Encoding UTF8 -NoTypeInformation



$users = import-csv .\LanxessToO365.csv

import-module msonline
$credential = get-credential

Connect-MsolService -Credential $credential

$newO365users = $users | ForEach-Object {
    New-MsolUser -DisplayName "$($_.FirstName) $($_.LastName)" `
    -UserPrincipalName $_.UserPrincipalName `
    -FirstName $_.FirstName `
    -LastName $_.LastName `
    -Department $_.BusinessUnit
}
$newO365users | Export-Csv CreatedO365Users.csv -Encoding UTF8



$users = $UsersToChange | Select-Object @{Label='FirstName';Expression = {$_.'first name'}} `
                ,@{Label='LastName';Expression = {$_.'Last name'}} `
                ,@{Label='BusinessUnit';Expression = {$_.'business unit'}} `
                ,@{Label='UserPrincipalName'; `
                    Expression={if (($_.'business unit' -eq 'HPE')-or($_.'business unit' -eq 'TSR')){
                        $_.'user name'.split('\')[1]+'@'+$_.'user name'.split('\')[0].replace('LX-','')+'.arlanxeo.com'
                    } else {
                        $_.'user name'.split('\')[1]+'@'+$_.'user name'.split('\')[0].replace('LX-','')+'.lanxess.com'
                    }}
                }`
                ,@{Label='PreviousUPN';Expression= {$_.'Primary E-Mail'}}   


$ChangedUsers =  $users | ForEach-Object {
   get-msoluser -UserPrincipalName $_.PreviousUPN | Set-MsolUserPrincipalName -NewUserPrincipalName $_.UserPrincipalName 
}               

$newO365users = $notexist | ForEach-Object {
    New-MsolUser -DisplayName "$($_.FirstName) $($_.LastName)" `
    -UserPrincipalName $_.UserPrincipalName `
    -FirstName $_.FirstName `
    -LastName $_.LastName `
    -Department $_.BusinessUnit
}

$NotChangedUsers =  $users | ForEach-Object {
    get-msoluser -UserPrincipalName $_.PreviousUPN 
}  

$O365rheinchemieUsers = $rheinchemieUsers | ForEach-Object {
    New-MsolUser -DisplayName "$($_.FirstName) $($_.LastName)" `
    -UserPrincipalName $_.UserPrincipalName `
    -FirstName $_.FirstName `
    -LastName $_.LastName `
    -Department $_.BusinessUnit
}


$newO365users  |ForEach-Object {
    $o365user = get-msoluser -UserPrincipalName $_.UserPrincipalName
     Add-MsolGroupMember -GroupObjectId $ProdGroup.ObjectId -GroupMemberType User -GroupMemberObjectId $o365user.ObjectId
     Add-MsolGroupMember -GroupObjectId $TrainingGroup.ObjectId -GroupMemberType User -GroupMemberObjectId $o365user.ObjectId
     Add-MsolGroupMember -GroupObjectId $UATGroup.ObjectId -GroupMemberType User -GroupMemberObjectId $o365user.ObjectId
}

foreach ($user in $newO365users) {
	if ($user.UsageLocation -eq $null) {
		Set-MsolUser -UserPrincipalName $user.userprincipalname -UsageLocation DE
    }
    if ($user.IsLicensed -eq $false) {
		Set-MsolUserLicense -UserPrincipalName $user.userprincipalname -AddLicenses $licenseSkuID
	} 
	
}