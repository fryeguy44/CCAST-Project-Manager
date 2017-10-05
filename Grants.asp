<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim strSubTitle 
Dim lngElementID

lngElementID = Request.QueryString("lngElementID")

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
If (CStr(Request("MM_insert")) = "frmAdd") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Grants (ElementID, UserID, GrantLevelID) VALUES (?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("htbxElementID"), Request.Form("htbxElementID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxUserID"), Request.Form("cbxUserID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("cbxGrantLevelID"), Request.Form("cbxGrantLevelID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>

<%
Dim rstElementGrants__lngElementID
rstElementGrants__lngElementID = "1"
	If (lngElementID <> "") Then 
	  rstElementGrants__lngElementID = lngElementID
	End If
%>
<%
Dim rstElementGrants
Dim rstElementGrants_cmd
Dim rstElementGrants_numRows

	Set rstElementGrants_cmd = Server.CreateObject ("ADODB.Command")
	rstElementGrants_cmd.ActiveConnection = MM_OBA_STRING
	rstElementGrants_cmd.CommandText = "SELECT Elements.ElementName, Users.DisplayName, Grants.UserID, Grants.ElementID, Grants.GrantID, GrantLevels.LevelName FROM GrantLevels INNER JOIN (((Grants INNER JOIN Users ON Grants.UserID = Users.UserID) INNER JOIN Elements ON Grants.ElementID = Elements.ElementID)) ON GrantLevels.GrantLevelID = Grants.GrantLevelID WHERE Grants.ElementID = ? ORDER BY Users.DisplayName" 
	rstElementGrants_cmd.Prepared = true
	rstElementGrants_cmd.Parameters.Append rstElementGrants_cmd.CreateParameter("param1", 5, 1, -1, rstElementGrants__lngElementID) ' adDouble
	
	Set rstElementGrants = rstElementGrants_cmd.Execute
	rstElementGrants_numRows = 0
%>
<%
Else
	Set rstElementGrants_cmd = Server.CreateObject ("ADODB.Command")
	rstElementGrants_cmd.ActiveConnection = MM_OBA_STRING
	rstElementGrants_cmd.CommandText = "SELECT Elements.ElementName, Users.DisplayName, Grants.UserID, Grants.ElementID, Grants.GrantID, GrantLevels.LevelName FROM GrantLevels INNER JOIN (((Grants INNER JOIN Users ON Grants.UserID = Users.UserID) INNER JOIN Elements ON Grants.ElementID = Elements.ElementID)) ON GrantLevels.GrantLevelID = Grants.GrantLevelID ORDER BY Elements.ElementName, Users.DisplayName" 
	rstElementGrants_cmd.Prepared = true
	
	Set rstElementGrants = rstElementGrants_cmd.Execute
	rstElementGrants_numRows = 0
End If	
%>
<%
Dim rstElements
Dim rstElements_cmd
Dim rstElements_numRows

Set rstElements_cmd = Server.CreateObject ("ADODB.Command")
rstElements_cmd.ActiveConnection = MM_OBA_STRING
rstElements_cmd.CommandText = "SELECT * FROM dbo.Elements ORDER BY ElementName" 
rstElements_cmd.Prepared = true

Set rstElements = rstElements_cmd.Execute
rstElements_numRows = 0
%>
<%
Dim rstUsers
Dim rstUsers_cmd
Dim rstUsers_numRows

Set rstUsers_cmd = Server.CreateObject ("ADODB.Command")
rstUsers_cmd.ActiveConnection = MM_OBA_STRING
rstUsers_cmd.CommandText = "SELECT * FROM dbo.Users ORDER BY DisplayName" 
rstUsers_cmd.Prepared = true

Set rstUsers = rstUsers_cmd.Execute
rstUsers_numRows = 0
%>
<%
Dim rstGrantLevels
Dim rstGrantLevels_cmd
Dim rstGrantLevels_numRows

Set rstGrantLevels_cmd = Server.CreateObject ("ADODB.Command")
rstGrantLevels_cmd.ActiveConnection = MM_OBA_STRING
rstGrantLevels_cmd.CommandText = "SELECT * FROM dbo.GrantLevels " 
rstGrantLevels_cmd.Prepared = true

Set rstGrantLevels = rstGrantLevels_cmd.Execute
rstGrantLevels_numRows = 0
%>

<%
If (CStr(Request("MM_insert")) = "frmAdd") Then
	lngAccessTypeID = 3
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

<h1>Security Grants</h1>

			<table border="0" cellspacing="0" cellpadding="0" class="fixed">
<%
If bolSecurityViewGranted Then

Dim strElementName

	strElementName = ""
	Do While Not rstElementGrants.EOF
		If bolSecurityEditGranted Then
			strEditLink = "<a href=""GrantEdit.asp?lngGrantID=" & (rstElementGrants.Fields.Item("GrantID").Value) & """>"
			strEndEditLink = "</a>&nbsp;"
		Else
			strEditLink = ""
			strEndEditLink = "&nbsp;"
		End If
		If strElementName <> (rstElementGrants.Fields.Item("ElementName").Value) Then
			strElementName = (rstElementGrants.Fields.Item("ElementName").Value)
%>            
              <tr>
                <td>&nbsp;</td>
                <td colspan="2"><strong>Element Name: <%=strEditLink & (rstElementGrants.Fields.Item("ElementName").Value) & strEndEditLink%></strong></td>
              </tr>
              <tr class="column_titles">
                <td>&nbsp;</td>
                <td><h4>User</h4></td>
                <td><h4>Security Level</h4></td>
              </tr>
              <tr class="line">
                <td>&nbsp;</td>
                <td colspan="2"><hr /></td>
              </tr>
<%
		End If
%>	              
              <tr class="tr_hover">
                <td>&nbsp;</td>
                <td><a href="GrantEdit.asp?lngGrantID=<%=(rstElementGrants.Fields.Item("GrantID").Value)%>"><%=(rstElementGrants.Fields.Item("DisplayName").Value)%></a></td>
                <td><%=(rstElementGrants.Fields.Item("LevelName").Value)%></td>
              </tr>
<%
		rstElementGrants.MoveNext
	Loop
	If bolSecurityAddGranted Then
		
%>              
			  <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
              <tr>
                <td>&nbsp;</td>
                <td><select name="cbxUserID" id="cbxUserID">
                  <%
While (NOT rstUsers.EOF)
%>
                  <option value="<%=(rstUsers.Fields.Item("UserID").Value)%>"><%=(rstUsers.Fields.Item("DisplayName").Value)%></option>
                  <%
  rstUsers.MoveNext()
Wend
If (rstUsers.CursorType > 0) Then
  rstUsers.MoveFirst
Else
  rstUsers.Requery
End If
%>
                </select></td>
                <td><select name="cbxGrantLevelID" id="cbxGrantLevelID">
                  <%
While (NOT rstGrantLevels.EOF)
%>
                  <option value="<%=(rstGrantLevels.Fields.Item("GrantLevelID").Value)%>"><%=(rstGrantLevels.Fields.Item("LevelName").Value)%></option>
                  <%
  rstGrantLevels.MoveNext()
Wend
If (rstGrantLevels.CursorType > 0) Then
  rstGrantLevels.MoveFirst
Else
  rstGrantLevels.Requery
End If
%>
                </select></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td><input type="submit" name="btnAdd" id="btnAdd" value="Add" />
                <input name="htbxElementID" type="hidden" id="htbxElementID" value="<%=lngElementID%>" /></td>
              </tr>
            <input type="hidden" name="MM_insert" value="frmAdd" />
            </form>
               
			   

<%
	End If
Else
%>	              
              <tr>
                <td colspan="3">Viewing this list requires certain &quot;Security&quot; permissions</td>
              </tr>
<%
End If
%>            
              <tr>
                <td colspan="3">&nbsp;</td>
              </tr>
            </table>
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
rstElementGrants.Close()
Set rstElementGrants = Nothing
%>
<%
rstElements.Close()
Set rstElements = Nothing
%>
<%
rstUsers.Close()
Set rstUsers = Nothing
%>
<%
rstGrantLevels.Close()
Set rstGrantLevels = Nothing
%>
