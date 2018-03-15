import-module msonline
$credential = get-credential

Connect-MsolService -Credential $credential

$sourcefile = 'C:\temp\20180119_User_PROD-RESET-PW-V05.csv'
$logfile = "PassowrdReset.log"
$users = import-csv $sourcefile -Delimiter ','

$users | ForEach-Object {
    try {
        Set-MsolUserPassword -UserPrincipalName $_.upn -NewPassword $_.NewPassword
        "$(Get-date -Format "dd.MM.yy_HH.mm.ss"): Password has been reseted for user $($_.UPN)" | Out-File $logfile -Append -Encoding unicode
    } catch {
        "$(Get-date -Format "dd.MM.yy_HH.mm.ss"): ERROR: Password has not been reseted for user $($_.UPN)" | Out-File $logfile -Append -Encoding unicode
    }
}