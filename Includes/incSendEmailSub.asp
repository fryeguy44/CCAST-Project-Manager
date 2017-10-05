<%
Sub SendNotificationEmail(strMessage, strRecipientEmail, strSubject, strFromEmail, strFileAddress)
'	Dim objEMail
'	Dim cdoConfig
'	Dim sch
'	
'    sch = "http://schemas.microsoft.com/cdo/configuration/" 
' 
'	Set cdoConfig = CreateObject("CDO.Configuration") 
'	With cdoConfig
'		.Fields(sch & "smtpserver") = "tgc-wasptest"
'		.Fields(sch & "smtpserverport") = 25
'		.Fields(sch & "sendusing") = 1
'		.Fields(sch & "smtpconnectiontimeout") = 60
'		.Fields.Update 
'	End With 
'	set objEMail = server.CreateObject("CDO.Message")
'	Set objEMail.Configuration = cdoConfig
'
'	'objEMail.BodyFormat = 0
'	'objEMail.MailFormat = 0
'	objEMail.To = strRecipientEmail
'	If strFromEmail = "" Then
'		objEMail.From = "rfrye@pci-tgc.org"
'	Else
'		objEMail.From = strFromEmail
'	End If
'	objEMail.subject = strSubject
'	objEMail.htmlbody = strMessage
'	If strFileAddress <> "0" Then
'		objEMail.AddAttachment strFileAddress 
'	End If
'	objEMail.send
'	set objEMail = nothing
End Sub
%>
