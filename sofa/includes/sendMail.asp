<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: function to send emails using most popular ASP components          
%>

<%
Function sendMail(fromName, from, rcpt, subject, body)

 pEmailComponent=lcase(pEmailComponent)
 
 on error resume next

if pEmailComponent="persitsaspmail" then

 set mail 	= server.CreateObject("Persits.MailSender")
 
 ' change with your smtp server
 
 mail.IsHTML    = false 
 mail.Host 	= pSmtpServer 
 mail.From 	= from
 mail.FromName 	= fromName
 mail.AddAddress rcpt
 mail.AddReplyTo from
 mail.Subject 	= subject
 mail.Body 	= body
 

 on error resume next
 
 mail.Send

end if

if pEmailComponent="jmail" then

 on error resume next
  
 set mail 	= server.CreateOBject( "JMail.Message" ) 
 mail.Logging 	= true
 mail.silent 	= true
 mail.From 	= from
 mail.FromName 	= fromName
 mail.Subject	= subject

 mail.AddRecipient rcpt

 'mail.MailServerUserName = strUser
 'mail.MailServerPassWord = strPassword

 mail.Body 	= body
 
 
 mail.send(pSmtpServer)
 
end if

if pEmailComponent="jmailhtml" then

 set mail 	= server.CreateOBject( "JMail.Message" ) 
 mail.Logging 	= true
 mail.silent 	= true
 mail.From 	= from
 mail.FromName 	= fromName
 mail.Subject	= subject

 mail.AddRecipient rcpt

 ' if the client cannot read HTML 
 mail.Body 	= body
 
 ' HTML 
 mail.HTMLBody = "<HTML><BODY><font color=""red"">" &pCompany& "</font><br>"
 mail.appendHTML "<img src=""http://" & pStoreLocation & "/store/images/" &pCompanyLogo& """>"
 mail.appendHTML "<br><br>" &body& "</BODY></HTML>"
 
 on error resume next
 
 mail.send(pSmtpServer)  

end if
 
if pEmailComponent="serverobjectsaspmail1" then
  
 set mail = Server.CreateObject("SMTPsvg.mail")

 mail.RemoteHost  = pSmtpServer

 mail.FromName    = fromName
 mail.FromAddress = from
 mail.AddRecipient rcpt, rcpt
 mail.Subject     = subject
 mail.BodyText    = body

 on error resume next
 
 mail.SendMail
 
end if

if pEmailComponent="serverobjectsaspmail2" then
  
 set mail = Server.CreateObject("SMTPsvg.Mailer")

 mail.RemoteHost  = pSmtpServer

 mail.FromName    = fromName
 mail.FromAddress = from
 mail.AddRecipient rcpt, rcpt
 mail.Subject     = subject
 mail.BodyText    = body

 on error resume next
 
 mail.SendMail
 
end if


if pEmailComponent="cdonts" then
 
 on error resume next
 
 Set mail 	 = Server.CreateObject ("CDONTS.NewMail")

 mail.BodyFormat = 1
 mail.MailFormat = 0
 
 on error resume next 
 
 mail.Send from, rcpt, subject, body
  
end if

if pEmailComponent="cdo" then
 
 on error resume next
 Set mail 	= server.CreateObject("CDO.Message")
 
 pMailServerUser	= getSettingKey("pMailServerUser")
 pMailServerPass	= getSettingKey("pMailServerPass")
 if trim(pMailServerUser) <> "" Then
	 Set oMailConfig = Server.CreateObject ("CDO.Configuration") 
	
	   oMailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = pSmtpServer
	   oMailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
	   oMailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
	   oMailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60 
	   oMailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")= 1 'cdoBasic
	   oMailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")= pMailServerUser
	   oMailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")= pMailServerPass
	   oMailConfig.Fields.Update 
	
	   Set mail.Configuration = oMailConfig 
  end if
 mail.From 	= from
 if pBcc="-1" then
 	mail.Bcc = rcpt
 	mail.To = from
 else
 	mail.To 	= rcpt
 end if
 mail.Subject 	= subject
 mail.textBody 	= body
 
 on error resume next
 mail.Send
 
 Set Mail = Nothing
end if

if pEmailComponent="cdontshtml" then

  dim mail
  set mail = Server.CreateObject("CDONTS.NewMail")     
  
  HTML = ""
  HTML = HTML & "<HTML>"
  HTML = HTML & "<HEAD>"
  HTML = HTML & "<TITLE>" &pCompany& "</TITLE>"
  HTML = HTML & "</HEAD>"
  HTML = HTML & "<BODY  bgcolor=""lightyellow"">"
  HTML = HTML & "<TABLE cellpadding=""4"">"
  HTML = HTML & "<TR><TH><FONT color=""darkblue""  SIZE=""4"">"  
  HTML = HTML & subject&"</FONT></TH></TR>" 
  HTML = HTML & "<TR><TD>"
  HTML = HTML & body
  HTML = HTML & "</FONT></TD></TR></TABLE><BR><BR>"
  HTML = HTML & "</BODY>"
  HTML = HTML & "</HTML>"
  
  mail.From    = from
  mail.Subject = subject
  mail.To      = rcpt
        
  ' HTML format
  mail.BodyFormat = 0 
  mail.MailFormat = 0 

  
  mail.Body 	     = HTML
  mail.Send
    
  set mail = nothing
  
end if

if pEmailComponent="cdo2" then

 on error resume next

 set mail = server.CreateObject("CDO.Message")

 set cdoConfig = Server.CreateObject("CDO.Configuration")
 cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
 cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")= pSmtpServer
 cdoConfig.Fields.Update
 set mail.Configuration = cdoConfig

 mail.From = from
 mail.To = rcpt
 mail.Subject = subject
 mail.textBody = body

 on error resume next
 mail.Send

 Set Mail = Nothing
 Set cdoConfig = nothing
end if
 
if pEmailComponent="bamboosmtp" then

 set mail 	= Server.CreateObject("Bamboo.SMTP") 
 mail.Server 	= pSmtpServer
 mail.Rcpt 	= rcpt 
 mail.From 	= from
 mail.FromName 	= fromName 
 mail.Subject 	= subject
 mail.Message 	= body 
 
 on error resume next 
 
 mail.Send 
  
 set mail = nothing 
 
end if

if pEmailComponent="ocxmail" then
 set mail = Server.CreateObject("ASPMAIL.ASPMailCtrl.1") 

 on error resume next 
 
 mail.SendMail pSmtpServer, rcpt, from, subject, body 

end if       

if pEmailComponent="dundasmail" then
 dim objMailer  
 set objEmail 		= Server.CreateObject("Dundas.Mailer")     
 objEmail.TOs.Add rcpt
 objEmail.FromAddress 	= from
 objEmail.Subject 	= subject 
 objEmail.SMTPRelayServers.Add pSmtpServer
 objEmail.Body 		= body
 objEmail.SendMail
 Set objEmail 		= Nothing 
end if
  
if pEmailComponent="aspsmartmail" then
 dim mySmartMail
 set mySmartMail = Server.CreateObject("aspSmartMail.SmartMail")
 ' change with your smtp server
 mySmartMail.Server 		= pSmtpServer
 mySmartMail.SenderName 	= fromName
 mySmartMail.SenderAddress 	= from
 'mySmartMail.Recipients.Add receipt,receipt name
 mySmartMail.Recipients.Add rcpt,""
 mySmartMail.Subject 		= subject
 mySmartMail.Body 		= body
 mySmartMail.SendMail
 set mySmartMail		=Nothing
 
end if

if pEmailComponent="MailEnable" then

 MAIL_SERVER = pSmtpServer
 
 'Create a message object
    Dim objMsg
    Set objMsg = server.CreateObject("MEMail.Message")

    'Populate the message
    With objMsg
        .ContentType = "text/html;"

        'Sender Info
        .MailFrom = from
        .MailFromDisplayName = fromName

        'To
        .MailTo rcpt

        'Body
        .Subject = subject
        .MessageBody  = body

        .Server = MAIL_SERVER
        .SendMessage

        'Pass back any errors
        SendEmailMailEnable = .ErrorString
    End With

    'Cleanup
    Set objMsg = nothing
 
end if
  
if err <> 0 and pDebugEmail ="-1" then  
%>  <html> 
    <br><br>
    Error while sending email:  <%=Err.Description%>. Have you installed and configured the right email component in this server?. If you don't want this message to appear just change the setting in Database Settings.
    <hr>
    <br>From    : <%=from%>    
    <br>To      : <%=rcpt%>
    <br>Subject : <%=subject%>
    <br>Body    : <%=body%>        
    </html>
   <%
 end if
 
end Function
%>

