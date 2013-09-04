<% 
'on error resume next 
Sub SendMail ()


ToEmail = objEmail
'objUserName = "mdnayem"
'objPassword = "Password"
Emailcc = "postmaster@eJuty.com"
EmailBody = "Please DO NOT Respond! "& vbcrlf &"Thank you for signing up with us. You now have the power to search for your right juty."& vbcrlf &"For your convenience, here is your login information."

Set MyMail = Server.CreateObject("CDO.Message") 

MyMail.From = "postmaster@ejuty.com" 
MyMail.To = ToEmail 
MyMail.Cc = Emailcc
MyMail.Subject = "Your Login" 
MyMail.textBody = EmailBody & vbcrlf &"Your User ID = '"& objUserName &"'(Note: User ID case sensitive!)" & vbcrlf & "And Password = '"& objPassword &"'"& vbcrlf &"URL = http://www.eJuty.com/Login.asp."& vbcrlf & "Thanks for choosing ejuty."&  vbcrlf & "Regards" & vbcrlf &"eJuty.com Team" 

'==This section provides the configuration information for the remote SMTP server. 
'==Normally you will only change the server name or IP. 

MyMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

'Name or IP of Remote SMTP Server 

MyMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.ejuty.com" 

'Server port (typically 25) 

MyMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 

MyMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@ejuty.com" 
MyMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "lailabanu" 

MyMail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate").Value = 1 
' The following fields are only used if the previous authentication value is set to 1 or 2 
MyMail.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/smtpaccountname").Value = "postmaster@ejuty.com" 

MyMail.Configuration.Fields.Update 

'==End remote SMTP server configuration section== 

MyMail.Send() 
'Clean up server resources. 
Set MyMail = nothing 


If err.number = 0 then 
response.write "" 
else 
response.write(err.number & " - " & err.description) 
end if 
End sub

%>