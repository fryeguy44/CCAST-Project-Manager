<%
Dim UserLog
Dim bGrantAccess
Dim strRequestForm
Dim strUserNameCheck

If Session("MM_Username") = "" Then
	strUserNameCheck = "Unknown"
	intUserID = 0
Else
	strUserNameCheck = Session("MM_Username")
	intUserID = Session("UserID")

End If

If MM_grantAccess Then
	bGrantAccess = 1
Else
	bGrantAccess = 0
End If

Set UserLog = Server.CreateObject ("ADODB.Command")
UserLog.ActiveConnection = MM_OBA_STRING
UserLog.CommandText = "INSERT INTO UserLogs (UserID, UserName, AccessTypeID, SIDPage, AccessDateTime, AccessGranted, PageCompleted) VALUES (?, ?, ?, ?, ?, ?, ?)" 
UserLog.Prepared = true
UserLog.Parameters.Append UserLog.CreateParameter("param1", 5, 1, -1, intUserID) ' adVarWChar
UserLog.Parameters.Append UserLog.CreateParameter("param1", 202, 1, 255, strUserNameCheck) ' adVarWChar
UserLog.Parameters.Append UserLog.CreateParameter("param2", 5, 1, -1, lngAccessTypeID) ' adVarWChar
UserLog.Parameters.Append UserLog.CreateParameter("param3", 202, 1, 255, strPageUrl) ' adVarWChar
UserLog.Parameters.Append UserLog.CreateParameter("param4", 135, 1, -1, dteAccessGranted) ' adVarWChar
UserLog.Parameters.Append UserLog.CreateParameter("param5", 5, 1, -1, bGrantAccess) ' adVarWChar
UserLog.Parameters.Append UserLog.CreateParameter("param6", 135, 1, -1, Now()) ' adVarWChar

UserLog.Execute
UserLog.ActiveConnection.Close
    
    
Dim rstNewUserLog
Dim lngNewUserLogID

Set rstNewUserLogcmd = Server.CreateObject ("ADODB.Command")
rstNewUserLogcmd.ActiveConnection = MM_OBA_STRING
rstNewUserLogcmd.CommandText = "SELECT TOP 1 UserLogID FROM UserLogs ORDER BY UserLogID DESC"
rstNewUserLogcmd.Prepared = true
Set rstNewUserLog = rstNewUserLogcmd.Execute

lngNewUserLogID = rstNewUserLog.Fields.Item("UserLogID").Value

rstNewUserLog.Close()
Set rstNewUserLog = Nothing

Set rstUserLogPageParametr = Server.CreateObject ("ADODB.Command")
rstUserLogPageParametr.ActiveConnection = MM_OBA_STRING
rstUserLogPageParametr.Prepared = true

For each inputField in Request.Form
	For each inputValue in Request.Form(inputField)
		If inputValue <> "" AND inputValue <> "Go" Then
			rstUserLogPageParametr.CommandText = "INSERT INTO UserLogPageParameters (UserLogID, InputValue, InputField, InputTypeQuerystring) VALUES (" & lngNewUserLogID & ", '" & Replace(Left(inputValue, 254),"'","") & "', '" & inputField & "', 0)" 
	
			rstUserLogPageParametr.Execute
		End If
	Next
Next
For each inputField in Request.QueryString
	For each inputValue in Request.QueryString(inputField)
		If inputValue <> "" AND inputValue <> "Go" Then
			rstUserLogPageParametr.CommandText = "INSERT INTO UserLogPageParameters (UserLogID, InputValue, InputField, InputTypeQuerystring) VALUES (" & lngNewUserLogID & ", '" & Replace(Left(inputValue, 254),"'","") & "', '" & inputField & "', 1)" 
	
			rstUserLogPageParametr.Execute
		End If
	Next
Next
rstUserLogPageParametr.ActiveConnection.Close
%>