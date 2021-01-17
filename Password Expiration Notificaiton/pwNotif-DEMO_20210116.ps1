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


pause 
#Reminder Email - Expires in 7 days
$data = Import-Csv .\user_list_$date.csv
$expiringSoon7 = $data | where {[int]$_.DaysUntilExpired -eq "7"}
$expiringSoon7 | Export-CSV .\7days_$date.csv
$emails7 = Import-CSV .\7days_$date.csv | select -ExpandProperty EmailAddress
$Subject7 = "Reminder: Windows Password Expiring in 7 Days"
$Body7 = "This is a reminder that your password is going to expire in 7 days. Please connect to the VPN and press Ctrl + Alt + Del to change your password before it expires. If you have any questions, please contact the helpdesk."
# DEMONSTRATION
echo "WHAT'S SUPPOSED TO HAPPEN:"
echo ""
echo "[7 Days] Send reminder emails to:"
if ($emails7 -eq $null) {
    echo ("No accounts expiring in 7 days")
} else{ 
    foreach ($email in $emails7) {
        echo "To: $email"
    } 
    echo "Subject: $Subject7"
    echo "Body: $Body7"
} echo ""
Start-Sleep 2
#REMINDER EMAIL
<#
foreach ($email in $emails7) {
  Send-MailMessage -To $email -From $MyEmail -Subject $Subject -Body $Body -SmtpServer $SMTP -Credential $Creds -UseSsl -Port 587 -DeliveryNotificationOption never
} #>

#Reminder Email - Expires in 3 days
$data = Import-Csv .\user_list_$date.csv
$expiringSoon3 = $data | where {[int]$_.DaysUntilExpired -eq "3"}
$expiringSoon3 | Export-CSV .\3days_$date.csv
$emails3 = Import-CSV .\3days_$date.csv | select -ExpandProperty EmailAddress
$Subject3 = "Reminder: Windows Password Expiring in 3 Days"
$Body3 = "This is a reminder that your password is going to expire in 3 days. Please connect to the VPN and press Ctrl + Alt + Del to change your password before it expires. If you have any questions, please contact the helpdesk."
# DEMONSTRATION
echo "[3 Days] Send reminder emails to:"
if ($emails3 -eq $null) {
    echo ("No accounts expiring in 3 days")
} else{ 
    foreach ($email in $emails3) {
        echo "To: $email"
    } 
    echo "Subject: $Subject3"
    echo "Body: $Body3"
} echo ""
Start-Sleep 2
#REMINDER EMAIL
<#
foreach ($email in $emails3) {
  Send-MailMessage -To $email -From $MyEmail -Subject $Subject -Body $Body -SmtpServer $SMTP -Credential $Creds -UseSsl -Port 587 -DeliveryNotificationOption never
} #>


#Reminder Email - Expires in 1 day
$data = Import-Csv .\user_list_$date.csv
$expiringSoon1 = $data | where {[int]$_.DaysUntilExpired -eq "1"}
$expiringSoon1 | Export-CSV .\1day_$date.csv
$emails1 = Import-CSV .\1day_$date.csv | select -ExpandProperty EmailAddress
$Subject1 = "Reminder: Windows Password Expiring in 1 Day"
$Body1 = "This is a reminder that your password is going to expire in 1 day. Please connect to the VPN and press Ctrl + Alt + Del to change your password before it expires. If you have any questions, please contact the helpdesk."
# DEMONSTRATION
echo "[1 Day] Send reminder emails to:"
if ($emails1 -eq $null) {
    echo ("No accounts expiring in 1 day")
} else{ 
    foreach ($email in $emails1) {
        echo "To: $email"
    } 
    echo "Subject: $Subject1"
    echo "Body: $Body1"
} echo ""
Start-Sleep 2
#REMINDER EMAIL
<#
foreach ($email in $emails1) {
  Send-MailMessage -To $email -From $MyEmail -Subject $Subject -Body $Body -SmtpServer $SMTP -Credential $Creds -UseSsl -Port 587 -DeliveryNotificationOption never
} #>
