# Gets the list domain users in specified list of OU's that have not logged in since after specified date
import-module activedirectory 
$domain = "host.some.some" 
$DaysInactive = 7
$time = (Get-Date).Adddays(-($DaysInactive))

# Get all AD User with lastLogonTimestamp less than our time and set to enable

$OU=@('ou=TS01,dc=host,dc=some,dc=biz','ou=TS02,dc=host,dc=some,dc=biz')

$ou | foreach { get-aduser -searchbase $_ -Filter {LastLogonTimeStamp -lt $time -and enabled -eq $true} -Properties LastLogonTimeStamp, Description, SamAccountName} |

# Output Name and lastLogonTimestamp into CSV

select-object Name,@{Name="Last login time"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}},Description, SamAccountName | export-csv "C:\Users\some\Documents\Inactive Users Notifier\OLD_Users.csv" -notypeinformation

# Send email

$smtpServer = 'relay.some.co.nz'
$file = "C:\Users\some\Documents\Inactive Users Notifier\OLD_Users.csv"
$att = new-object Net.Mail.Attachment($file)
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "support@some.biz"
$msg.To.Add("techs@some.biz")
$msg.Subject = "Inactive (more then 7 days) hosting users report"
$msg.Attachments.Add($att)
$smtp.Send($msg)
$att.Dispose()





