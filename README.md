# windowsPWexpirationEmail
Problem:
- With remote environment, the users' PCs are unable to tell when their passwords are expiring. Users often call the helpdesk after their passwords have expired.

Solution: 
- A powershell script that looks for all AD users' Windows password expiration dates, sending an email notification to those expiring in 7 days, 3 days, and 1 day. 
- This script does not change passwords for users. It is still up to the user to change their password after seeing the three email notifications. 
- If user changes password, they shouldn't get subsequent notificaiton emails.

Instructions:
- Place the scripts in a folder on a PC that is joined to the domain, logged in as an administrator's account.
- All CSV files will be saved in that folder. 
