<%

'This code has been written with the help of http://blog.tacosoup.com/sending-email-through-gmail-using-classic-asp/
'Note the Google must be configured with Less Secure Apps 'on' for this to work

	Dim Mail				
	Set Mail = CreateObject("CDO.Message")
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="smtp.gmail.com"
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="greensonscreen1@gmail.com" 'You can also use you email address that’s setup through google apps.
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="pptudnqvwynssrid"	'set up in Gmail as an app password
	Mail.Configuration.Fields.Update

 	Mail.To = strTo
 	Mail.Cc = strCc
 	Mail.Bcc = strBcc
 	Mail.From = strFrom
 	Mail.Subject = subject
 	Mail.HTMLBody = message
 	
 On Error Resume Next
 
 	Mail.Send
 	
 	if Err.Number <> 0 then
 		response.write("<p class=""style4boldred"">An error has occured but your process has completed successfully</p>")  
 		response.write("<p class=""style4boldred"">Error details: " & Err.Number & "</p>")
 		response.write("<p class=""style4boldred"">" & Err.Description & "</p>")
 		response.write("<p class=""style4boldred"">" & Err.Source & "</p>")
 		response.write("<p class=""style4boldred"">PLEASE email Steve (steve@greensonscreen.co.uk) to tell him about this message, and if possible, include a screen print.</p>")
 	end if  
 	
 On Error Goto 0
 	
 	Set Mail = Nothing
 	
%>