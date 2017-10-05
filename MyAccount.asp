<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%
Dim strReturnPath

If Request.QueryString("strReturnPath") = "" Then
	strReturnPath = Request.ServerVariables("HTTP_REFERER")
Else
	strReturnPath = Request.QueryString("strReturnPath")
End If
%>
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
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
Dim rstUsers__lngUserID
rstUsers__lngUserID = "1"
If (lngUserID <> "") Then 
  rstUsers__lngUserID = lngUserID
End If
%>
<%
Dim bolPasswordChecks
bolPasswordChecks = True

If (CStr(Request("MM_update")) = "frmEditUser") Then
	If Request.Form("tbxPassword") <> Request.Form("tbxConfirmPassword") Then
		bolPasswordChecks = False
	Else
%>
<%
If (CStr(Request("MM_update")) = "frmEditUser") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "UPDATE dbo.Users SET LandingPageID = ?, EmailAddress = ?, Password = ? WHERE UserID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("cbxLandingPageID")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxEmailAddress")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 20, Request.Form("tbxPassword")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 200, 1, 20, Request.Form("MM_recordId")) ' adVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
	
	'Response.Redirect(Request.Form("cbxLandingPage"))
  End If
End If
%>
<%
	lngAccessTypeID = 2
	End If
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr"><!-- InstanceBegin template="/Templates/master.dwt" codeOutsideHTMLIsLocked="false" -->
<%
If lngAccessTypeID = "" Then
	lngAccessTypeID = 1
End If
%>
<!--#include file="Templates/incMasterSecurity.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=strPageTitle & " " & strSubTitle%></title>
<link rel="shortcut icon" href="favicon.ico" />

<!-- CSS Global -->
<link href="/global/css/global.css"rel="stylesheet" type="text/css" />

<!-- CSS Local -->
<link href="/local/css/local.css" rel="stylesheet" type="text/css" />

<!-- Print Stylesheet -->
<link rel="stylesheet" type="text/css" href="/global/css/print.css" media="print" />

<!-- InstanceBeginEditable name="Head" -->
<%
If (CStr(Request("MM_update")) = "frmEditUser") Then
	Response.Redirect(Request.Form("htbxReturnPath"))
End If
%>

<script src="SpryAssets/SpryValidationTextField.js" type="text/javascript"></script>
<link href="SpryAssets/SpryValidationTextField.css" rel="stylesheet" type="text/css" />
<!-- InstanceEndEditable -->

</head>

<body>

<!-- Begin Header -->
<div id="wrapper"> <!-- Wrapper div creates sticky footer -->
	<div id="header">
		<table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr>
			<td width="256">&nbsp;</td>
			<td width="364" align="center">
				<div id="sid"></div></td>
			<td width="211">
              <div align="right" id="ccast">
			  <p><a href="MyAccount.asp"><%=Session("MM_Username")%> Profile</a> | <a href="logoff.asp">Log Out</a></p>
              </div></td>
			</tr>
		</table>
	</div>
<!-- End Header -->

<!-- Begin Nav & Search -->
	<div id="nav_bar">
		
		<div id="nav">
			<!-- Quick menu moved to local folder to support different color schemes -->
			<!--#include virtual="/local/menu/incQuickMenu.asp" -->
		</div>
		
		<div id="nav_search">			
            <a href="/help/index.htm?context=<%=intHelpContextID%>" target="_blank" class="help">Help</a>
  		</div>
	</div>
<!-- End Nav & Search -->

<!-- Begin Content -->
	
	<div id="content">

<!-- #BeginEditable "content" -->
<%
  lngUserID = Session("UserID")
%>
<%
Dim rstUsers
Dim rstUsers_cmd
Dim rstUsers_numRows

Set rstUsers_cmd = Server.CreateObject ("ADODB.Command")
rstUsers_cmd.ActiveConnection = MM_OBA_STRING
rstUsers_cmd.CommandText = "SELECT UserName, Password, EmailAddress, Title FROM Users WHERE UserID = ?" 
rstUsers_cmd.Prepared = true
rstUsers_cmd.Parameters.Append rstUsers_cmd.CreateParameter("param1", 5, 1, -1, rstUsers__lngUserID) ' adDouble

Set rstUsers = rstUsers_cmd.Execute
rstUsers_numRows = 0
%>

<h1><%=strPageTitle & " <span>" & strSubTitle%></span></h1>

			<table border="0" cellspacing="0" cellpadding="0" class="fixed account">
              <form id="frmEditUser" name="frmEditUser" method="POST" action="<%=MM_editAction%>" autocomplete="off">
                <tr>
                	<td width="5">&nbsp;</td>
                	<td width="194">&nbsp;</td>
                	<td>&nbsp;</td>
               	  </tr>
                
                <tr>
                  <td>&nbsp;</td>
                  <td align="right"><strong>Username</strong></td>
                  <td><%=(rstUsers.Fields.Item("UserName").Value)%></td>
                </tr>
                  
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Title</strong></td>
                    <td><input type="text" name="tbxTitle" id="tbxTitle" /></td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Email</strong></td>
<%
Dim strEmailAddress
If InStr("?",Request.ServerVariables("HTTP_REFERER")) = 0 Then
	strLandingPage = Request.ServerVariables("HTTP_REFERER")
Else
	strLandingPage = Left(Request.ServerVariables("HTTP_REFERER"),InStr("?",Request.ServerVariables("HTTP_REFERER")))
End If
%>

                    <td><input name="tbxEmailAddress" type="email" id="tbxEmailAddress" value="<%=(rstUsers.Fields.Item("EmailAddress").Value)%>" size="35" tabindex="2" required="required" /></td>
                  </tr>
                  
                <%
If Not bolPasswordChecks Then  
%>                
                  <tr>
                    <td colspan="3"><span class="style4">The password typed does not match the confirm password. Please try again</span>.</td>
                  </tr>
<%
End If
%>                  
                <tr>
                  <td>&nbsp;</td>
                  <td align="right"><strong>Password</strong></td>
                  <td><span id="sprytextfield1">
                  <input name="tbxPassword" type="password" id="tbxPassword" value="<%=(rstUsers.Fields.Item("Password").Value)%>" tabindex="3" />
                  <span class="textfieldRequiredMsg">A value is required.</span><span class="textfieldMinCharsMsg">Minimum number of characters not met.</span></span></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td align="right" nowrap="nowrap"><strong>Confirm Password</strong></td>
                  <td><input name="tbxConfirmPassword" type="password" id="tbxConfirmPassword" value="<%=(rstUsers.Fields.Item("Password").Value)%>" tabindex="4" /></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" /></td>
                  <td><input type="submit" name="btnEdit" id="btnEdit" value="Update" /></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <input type="hidden" name="MM_update" value="frmEditUser" />
                <input type="hidden" name="MM_recordId" value="<%=lngUserID%>" />
</form>
		    </table>
			<script type="text/javascript">
<!--
var sprytextfield1 = new Spry.Widget.ValidationTextField("sprytextfield1", "none", {minChars:4});
//-->
            </script>
<!-- #EndEditable -->

</div>

<!-- End Content -->

<!-- Begin Footer -->

	<div id="push"></div> <!-- Push for sticky footer -->

</div><!-- End Wrapper -->
	
<!--#include file="Includes/incFooter.asp" -->
<!-- End Footer -->

</body>

<!-- InstanceEnd --></html>
<%
rstUsers.Close()
Set rstUsers = Nothing
%>