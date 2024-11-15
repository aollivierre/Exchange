[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$EmailFrom = "SMTP@glebecentre.ca"
$EmailTo = "aollivierre@novanetworks.com"
$Subject = "The subject of your email"
$Body = "What do you want your email to say"
$SMTPServer = "smtp.office365.com"
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587)
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential("SMTP@glebecentre.ca", "Wap33621");
$SMTPClient.Send($EmailFrom, $EmailTo, $Subject, $Body)