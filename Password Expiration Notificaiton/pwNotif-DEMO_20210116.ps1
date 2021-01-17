Import-Module ActiveDirectory
#Get complete list of users whose passwords is set to expire
$MaxPwdAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
$expiredDate = (Get-Date).addDays(-$MaxPwdAge)
$ExpiredUsers = Get-ADUser -Filter {(PasswordLastSet -gt $expiredDate) -and (PasswordNeverExpires -eq $false) -and (Enabled -eq $true)} -Properties PasswordNeverExpires, PasswordLastSet, Mail, EmailAddress | select samaccountname, EmailAddress, PasswordLastSet, @{name = "DaysUntilExpired"; Expression = {$_.PasswordLastSet - $ExpiredDate | select -ExpandProperty Days}} | Sort-Object PasswordLastSet
$date = Get-Date -Format "yyyyMMdd"
$ExpiredUsers | Export-Csv -NoType .\user_list_$date.csv
$ExpiredUsers
echo ""
Start-Sleep 2

#Email Setup
$MyEmail = "seiji.a.matsumoto@mitsubishicorp.com"
$SMTP = "smtp.office365.com"
$MyPass = ConvertTo-SecureString "Macintosh0105!!" -AsPlainText -Force
$Creds = New-Object System.Management.Automation.PSCredential ($MyEmail, $MyPass)


#Days
$Days = "7", "3", "1"

echo "WHAT'S SUPPOSED TO HAPPEN:"
foreach ($day in $Days) {
 $data = Import-Csv .\user_list_$date.csv
 $expiringSoon = $data | where {[int]$_.DaysUntilExpired -eq $day}
 $expiringSoon | Export-CSV .\$daydays_$date.csv
 $emails = Import-CSV .\$daydays_$date.csv | select -ExpandProperty EmailAddress
 $Subject = "Reminder: Windows Password Expiring in $day Day(s)"
 $Body = "This is a reminder that your password is going to expire in $day day(s). Please connect to the VPN and press Ctrl + Alt + Del to change your password before it expires. If you have any questions, please contact the helpdesk."
 # DEMONSTRATION
 echo ""
 echo "[$day Day(s)] Send reminder emails to:"
 if ($emails -eq $null) {
     echo ("No accounts expiring in $day day(s)")
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

