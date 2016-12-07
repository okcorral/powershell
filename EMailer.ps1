Function EMailer ($recipients){

[string[]]$emailTo = $recipients.Split(',')

$recipientsCC = "hyps9_support@rccl.com"
#$recipientsCC = ""
[string[]]$emailCC = $recipientsCC.Split(',')

$recipientsBCC = "hyperion_broadcast@rccl.com"
#$recipientsBCC = ""
[string[]]$emailBCC = $recipientsBCC.Split(',')

$attachmentFiles = "D:\Oracle\Middleware\EPMSystem11R1\products\biplus\InstallableApps\weblogic.xml, D:\Oracle\Middleware\logs\log.txt"
#$attachmentFiles = ""
[string[]]$attachments = $attachmentFiles.Split(',')

 
$fromaddress = "HyperionAdmin@rccl.com" 
$toaddress = $emailTo 

$Subject = "Prod EPM Services Alert" 
$body = @"
This is an alert from the EPM Servers. 
 
Thank you, 
Hyperion Administration 
RCCL 
HyperionUpgradeSupport@RCCL.com 
"@  

$smtpserver = "mrmrelay.rccl.com" 
$message = new-object System.Net.Mail.MailMessage 

$message.From = $fromaddress 

for ($i=0; $i -lt $emailTo.length; $i++) {
    $message.To.Add($emailTo[$i]) 
}
if ($recipientsCC.Length -gt 0){
    for ($i=0; $i -lt $emailCC.length; $i++){ 
        $message.CC.Add($emailCC[$i]) 
    }
}
if ($recipientsBCC.Length -gt 0){
    for ($i=0; $i -lt $emailBCC.length; $i++){ 
        $message.Bcc.Add($emailBCC[$i]) 
    }
}

$message.IsBodyHtml = $False 
$message.Subject = $Subject 

if ($attachmentFiles.Length -gt 0){
    for ($i=0; $i -lt $attachments.length; $i++) {
        $attach = new-object Net.Mail.Attachment($attachments[$i]) 
        $message.Attachments.Add($attach) 
    }
}

$message.body = $body 

$smtp = new-object Net.Mail.SmtpClient($smtpserver) 
$smtp.Send($message) 

} 

EMailer("tcorral@rccl.com")
#EMailer("tcorral@rccl.com,hyps9_support@rccl.com")