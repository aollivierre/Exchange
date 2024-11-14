$EmailFrom = "richard@richardd.ca"
$EmailTo = "richard@richardd.ca"
$Subject = "The subject of your email"
$Body = "What do you want your email to say"
$SMTPServer = "smtp.office365.com"
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587)
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential("richard@richardd.ca", "yourpassword");
$SMTPClient.Send($EmailFrom, $EmailTo, $Subject, $Body)