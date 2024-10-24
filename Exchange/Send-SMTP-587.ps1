$EmailFrom = "SMTP@arnpriorhealth.ca"
$EmailTo = "aollivierre@novanetworks.com"
$Subject = "The subject of your email"
$Body = "What do you want your email to say"
$SMTPServer = "smtp.office365.com"
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587)
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential("SMTP@arnpriorhealth.ca", "Kav51862");
$SMTPClient.Send($EmailFrom, $EmailTo, $Subject, $Body)