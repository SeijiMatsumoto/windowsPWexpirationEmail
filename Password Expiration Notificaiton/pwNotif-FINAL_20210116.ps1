Import-Module ActiveDirectory
#Get complete list of users whose passwords is set to expire
$MaxPwdAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
$expiredDate = (Get-Date).addDays(-$MaxPwdAge)
$ExpiredUsers = Get-ADUser -Filter {(PasswordLastSet -gt $expiredDate) -and (PasswordNeverExpires -eq $false) -and (Enabled -eq $true)} -Properties PasswordNeverExpires, PasswordLastSet, Mail, EmailAddress | select samaccountname, EmailAddress, PasswordLastSet, @{name = "DaysUntilExpired"; Expression = {$_.PasswordLastSet - $ExpiredDate | select -ExpandProperty Days}} | Sort-Object PasswordLastSet
$date = Get-Date -Format "yyyyMMdd"
$ExpiredUsers | Export-Csv -NoType .\PasswordExpiration\user_list_$date.csv

#Email Setup
$MyEmail = "seiji.a.matsumoto@mitsubishicorp.com"
$SMTP = "smtp.office365.com"
$MyPass = Get-Content ".\Password.txt" | ConvertTo-SecureString
$Creds = New-Object System.Management.Automation.PSCredential ($MyEmail, $MyPass)

#Days
$Days = "7", "3", "1"

# Send an email reminder to users whose passwords expire in 7 days, 3 days, and 1 day
foreach ($day in $Days) {
 $data = Import-Csv .\user_list_$date.csv
 $expiringSoon = $data | where {[int]$_.DaysUntilExpired -eq $day}
 $expiringSoon | Export-CSV .\$daydays_$date.csv
 $emails = Import-CSV .\$daydays_$date.csv | select -ExpandProperty EmailAddress
 $Subject = "Reminder: Windows Password Expiring in $day Days"
 $Body = "This is a reminder that your password is going to expire in $day days. Please connect to the VPN and press Ctrl + Alt + Del to change your password before it expires. If you have any questions, please contact the helpdesk."
 # DEMONSTRATION
 echo "WHAT'S SUPPOSED TO HAPPEN:"
 echo ""
 echo "[$day Days] Send reminder emails to:"
 if ($emails -eq $null) {
     echo ("No accounts expiring in $day days")
 } else{ 
     foreach ($email in $emails) {
         echo "To: $email"
     } 
     echo "Subject: $Subject"
     echo "Body: $Body"
 } echo ""
 Start-Sleep 2
 #REMINDER EMAIL
 <#
 foreach ($email in $emails) {
   Send-MailMessage -To $email -From $MyEmail -Subject $Subject -Body $Body -SmtpServer $SMTP -Credential $Creds -UseSsl -Port 587 -DeliveryNotificationOption never
 } #>
}