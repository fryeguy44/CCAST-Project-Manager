<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
If Request.Form("tbxUsername") <> "" Then
	MM_valUsername = LCase(CStr(Request.Form("tbxUsername")))
Else
	MM_valUsername = CStr(Request.Querystring("EmailPassword"))
End If
'Response.Write("MM_LoginAction = " & MM_LoginAction & "<br />")
'Response.Write("MM_valUsername = " & MM_valUsername & "<br />")
'Response.Write("Request.QueryString = " & Request.QueryString & "<br />")
'Response.Write("Request.Form(tbxUsername) = " & Request.Form("tbxUsername") & "<br />")
If MM_valUsername <> "" Then
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  bolLoginSuccess = False
 
  MM_redirectLoginFailed = "logon.asp"

  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_OBA_STRING
  MM_rsUser_cmd.CommandText = "SELECT Users.UserID, Users.UserName, Users.PositionID, Users.Password, Users.EmailAddress, Users.VendorID, Users.ClientID, Pages.PageAddress AS LandingPage FROM Users INNER JOIN Pages ON Users.LandingPageID = Pages.PageID WHERE (UserName = ?) AND Users.Active = 1"
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 255, MM_valUsername) ' adVarChar
  'MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 255, Request.Form("tbxPassword")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
  	bolValidUser = True
	intPositionID = MM_rsUser.Fields.Item("PositionID").Value
	lngUserID = MM_rsUser.Fields.Item("UserID").Value
	lngVendorID = MM_rsUser.Fields.Item("VendorID").Value
	lngClientID = MM_rsUser.Fields.Item("ClientID").Value
	strLandingPage = MM_rsUser.Fields.Item("LandingPage").Value
	'Response.Write("MM_User VendorID = " & MM_rsUser.Fields.Item("VendorID").Value & "<br />")
	
  End If
  
  If bolValidUser Then
	' username esixts - this is a valid user
	'Add user password if none currently exists
Response.Write("Password = " & MM_rsUser.Fields.Item("Password").Value & "<br />")
Response.Write("tbxPassword = " & Request.Form("tbxPassword") & "<br />")
  	
	If  MM_rsUser.Fields.Item("Password").Value = "" OR IsNull(MM_rsUser.Fields.Item("Password").Value) Then
		MM_rsUser_cmd.CommandText = "UPDATE Users SET Password = ? WHERE UserID = ?"
		MM_rsUser_cmd.Parameters.Delete 0
		MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 255, Request.Form("tbxPassword")) ' adVarChar
		MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 5, 1, -1, lngUserID) ' adVarChar
		MM_rsUser_cmd.Prepared = true
		MM_rsUser_cmd.Execute
		bolLoginSuccess = True
	Else
	'Authenticate valid user with correct password
	

		strEmailAddress = MM_rsUser.Fields.Item("EmailAddress").Value
		strPassword = MM_rsUser.Fields.Item("Password").Value
		If strPassword = Request.Form("tbxPassword") Then
			  bolLoginSuccess = True
		Else
		  MM_redirectLoginFailed = "logon.asp?logon=" & MM_valUsername
		End If
	End If
	
	If bolLoginSuccess Then
		Session("MM_Username") = MM_valUsername
		Session("UserID") = lngUserID
		Session("PositionID") = intPositionID
		Session("VendorID") = lngVendorID
		Session("ClientID") = lngClientID
		Response.Cookies("MM_Username") = MM_valUsername
		Response.Cookies("MM_Username").Expires =  Date 
		Response.Cookies("UserID") = lngUserID
		Response.Cookies("UserID").Expires =  Date 
		Response.Cookies("PositionID") = intPositionID
		Response.Cookies("PositionID").Expires = Date
		Response.Cookies("VendorID") = lngVendorID
		Response.Cookies("VendorID").Expires = Date
		Response.Cookies("ClientID") = lngClientID
		Response.Cookies("ClientID").Expires = Date
		If MM_IIf(Request.Form("chkKeepLogin"), True, False) Then
			Response.Cookies("MM_UsernameLongTerm") = MM_valUsername
			Response.Cookies("MM_UsernameLongTerm").Expires =  Date + 365
			Response.Cookies("UserIDLongTerm") = lngUserID
			Response.Cookies("UserIDLongTerm").Expires =  Date + 365
		End If
		MM_redirectLoginSuccess = strLandingPage
		if CStr(Request.QueryString("accessdenied")) <> "" Then
		  MM_redirectLoginSuccess = Request.QueryString("accessdenied")
		End If
		Session("intLogonAttempts") = 0
		MM_rsUser.Close
		Set MM_rsUser = Nothing
		Response.Redirect(MM_redirectLoginSuccess)
		'Response.Write("Session VendorID = " & Session("VendorID"))
	End If
  End If
  MM_rsUser.Close
  Set MM_rsUser = Nothing
	If Request.Querystring("EmailPassword") = "" Then
		Response.Redirect(MM_redirectLoginFailed)
	End If
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
strPageTitle="Logon"
%>
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=strPageTitle & " " & strSubTitle%></title>
<link rel="shortcut icon" href="favicon.ico">

<!-- CSS Global -->
<link href="/global/css/global.css"rel="stylesheet" type="text/css" />

<!-- CSS Local -->
<link href="/local/css/local.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" type="text/css" media="print" href="/local/css/print.css" />
<link href="SpryAssets/SpryValidationPassword.css" rel="stylesheet" type="text/css" />
<script src="SpryAssets/SpryValidationPassword.js" type="text/javascript"></script>
</head>

<body>

<!-- Begin Header -->
<div id="wrapper"> <!-- Wrapper div creates sticky footer -->
	<div id="header">
		<table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr>
			<td width="226">&nbsp;</a>
			</td>
			<td width="364" align="center">
				<div id="sid"></div>
			</td>
			<td width="211">
				<div id="ccast"></div>
				</td></tr>
		</table>
	</div>
<!-- End Header -->

<!-- Begin Nav & Search -->
	<div id="nav_bar">
		
		<div id="nav">
			<!-- Quick menu moved to local folder to support different color schemes -->
		</div>
		
		<div id="nav_search">			
            <a href="/help/index.htm?context=80" target="_blank" class="help">Help</a>
  		</div>
	</div>
<!-- End Nav & Search -->

<!-- Begin Content -->

<div style="width:600px;margin:0 auto;clear:both;height:300px;">
			
	<h1><%=strPageTitle%></h1>
		
	<div style="display:block;postion:relative;width:400px;margin:0 auto;">

<form id="frmLogon" name="frmLogon" method="POST" action="<%=MM_LoginAction%>">
	<table width="100%" border="0" cellspacing="0" cellpadding="5">
<%
If Request.QueryString("logon") <> ""  Then
%>  
		<tr>
		  <td colspan="2" class="AlertTagRed">Incorrect Logon Username/Password. Please see your Casino Manager or System Administrator to have your password reset.</a></td>
		</tr>
<%
End If
%>          
<%
If Request.QueryString("EmailPassword") <> ""  Then
	If strEmailAddress = "" OR IsNull(strEmailAddress) Then
		strMessageTo = "bryson@boydgrp.com "
		strTextBody = "The password for "& Request.QueryString("EmailPassword") & " is " & strPassword
		strDisplayMessage = "Because your email address is not in the Database, your password has been sent to the Application Administrator for forwarding to you. Please be patient."
	Else
		strMessageTo = strEmailAddress
		strTextBody = "Your password is " & strPassword
		strDisplayMessage = "Your password has been emailed to the address in your profile. Please fetch and enter it below."
	End If
	Set objMessage = CreateObject("CDO.Message") 
	objMessage.Subject = "B-4 Trucking Application password reset" 
	objMessage.From = "NotificationS@4btrucking.com" 
	objMessage.To = strMessageTo 
	objMessage.TextBody = strTextBody

	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "dedrelay.secureserver.net"
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
	objMessage.Configuration.Fields.Update
	objMessage.Send

%>  
		<tr>
		  <td colspan="2"><%=strDisplayMessage%></td>
		</tr>
<%
End If
%>          
		<tr>
			<td style="vertical-align:middle;"><h2>Username</h2></td>
			<td style="vertical-align:middle;"><input name="tbxUsername" type="text" class="formField" id="tbxUsername" style="font-family:Arial, Helvetica, sans-serif;font-size:18px; padding: 7px; height:26px;" value="<%=MM_valUsername%>" /></td>
		</tr>
		<tr>
			<td style="vertical-align:middle;"><h2>Password</h2></td>
			<td style="vertical-align:middle;"><span id="sprypassword1">
            <input type="password" name="tbxPassword" id="tbxPassword" class="formField" style="font-family:Arial, Helvetica, sans-serif;font-size:18px; padding: 7px; height:26px;" />
            <br /><span class="passwordRequiredMsg">A value is required.</span><span class="passwordInvalidStrengthMsg">The password doesn't meet the specified strength.</span></span></td>
		</tr>
		<tr>
			<td colspan="2" align="center"><input type="checkbox" name="chkKeepLogin" id="chkKeepLogin" />
			Keep me logged into CCAST on this computer until I log off (Not recommended for public or shared computers)</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td><input type="submit" name="btnLogon" id="btnLogon" value="Logon" /></td>
		</tr>
	</table>
	</form>

</div>
</div>
<!-- End Content -->

<!-- Begin Footer -->

	<div id="push"></div> <!-- Push for sticky footer -->

</div><!-- End Wrapper -->
	
	<div id="footer">
	
		<div id="footer_inside">
			<p><a href="Logon.asp">Logon</a></p>
		</div>
</div>

<!-- End Footer -->

<script type="text/javascript">
var sprypassword1 = new Spry.Widget.ValidationPassword("sprypassword1", {minAlphaChars:4});
</script>
</body>

</html>