<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%
Dim lngGrantID
Dim strReturnPath

lngGrantID = Request.Querystring("lngGrantID")
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
If (CStr(Request("MM_update")) = "frmEdit") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "UPDATE dbo.Grants SET UserID = ?, ElementID = ?, GrantLevelID = ? WHERE GrantID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("cbxUserID"), Request.Form("cbxUserID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxElementID"), Request.Form("cbxElementID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("cbxGrantLevelID"), Request.Form("cbxGrantLevelID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_update")) = "frmEdit") Then
	lngAccessTypeID = 2
End If
%>

<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "DELETE FROM dbo.Grants WHERE GrantID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If

End If
%>
<%
If (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then
	lngAccessTypeID = 4
End If
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr"><!-- InstanceBegin template="/Templates/Edit.dwt" codeOutsideHTMLIsLocked="false" -->
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
If (CStr(Request("MM_update")) = "frmEdit") OR (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then
    Response.Redirect(Request.Form("htbxReturnPath"))
End If
%>

<%
Dim rstGrants__lngGrantID
rstGrants__lngGrantID = "1"
If (lngGrantID <> "") Then 
  rstGrants__lngGrantID = lngGrantID
End If
%>
<%
Dim rstGrants
Dim rstGrants_cmd
Dim rstGrants_numRows

Set rstGrants_cmd = Server.CreateObject ("ADODB.Command")
rstGrants_cmd.ActiveConnection = MM_OBA_STRING
rstGrants_cmd.CommandText = "SELECT * FROM dbo.Grants WHERE GrantID = ?" 
rstGrants_cmd.Prepared = true
rstGrants_cmd.Parameters.Append rstGrants_cmd.CreateParameter("param1", 5, 1, -1, rstGrants__lngGrantID) ' adDouble

Set rstGrants = rstGrants_cmd.Execute
rstGrants_numRows = 0
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
Dim rstGrantLevels
Dim rstGrantLevels_cmd
Dim rstGrantLevels_numRows

Set rstGrantLevels_cmd = Server.CreateObject ("ADODB.Command")
rstGrantLevels_cmd.ActiveConnection = MM_OBA_STRING
rstGrantLevels_cmd.CommandText = "SELECT * FROM dbo.GrantLevels" 
rstGrantLevels_cmd.Prepared = true

Set rstGrantLevels = rstGrantLevels_cmd.Execute
rstGrantLevels_numRows = 0
%>

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
              <div id="ccast">
			  <p><a href="MyAccount.asp"><%=Session("MM_Username")%> Profile</a> | <a href="logoff.asp">Log Out</a></div></td></p>
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

<h1><%=strPageTitle & " " & strSubTitle%></h1>

			<table border="0" cellspacing="0" cellpadding="0" class="box">
<%
If bolSecurityEditGranted Then
%>
                  <form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
				  <tr>
				  	<td>&nbsp;</td>
				  	<td>&nbsp;</td>
				  	<td>&nbsp;</td>
				  	<td>&nbsp;</td>
			  	</tr>
				  <tr>
                    <td width="10">&nbsp;</td>
                    <td><strong>User</strong></td>
                    <td><label>
                    <select name="cbxUserID" id="cbxUserID">
                      <%
While (NOT rstUsers.EOF)
%><option value="<%=(rstUsers.Fields.Item("UserID").Value)%>" <%If (Not isNull((rstGrants.Fields.Item("UserID").Value))) Then If (CStr(rstUsers.Fields.Item("UserID").Value) = CStr((rstGrants.Fields.Item("UserID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstUsers.Fields.Item("DisplayName").Value)%></option>
                      <%
  rstUsers.MoveNext()
Wend
If (rstUsers.CursorType > 0) Then
  rstUsers.MoveFirst
Else
  rstUsers.Requery
End If
%>
                    </select>
                    </label></td>
                    <td>&nbsp;</td>
                  </tr>                  
                  <tr>
                    <td width="10">&nbsp;</td>
                    <td width="128"><strong>Element</strong></td>
                    <td width="298">
<label>
                    <select name="cbxElementID" id="cbxElementID">
                      <%
While (NOT rstElements.EOF)
%><option value="<%=(rstElements.Fields.Item("ElementID").Value)%>" <%If (Not isNull((rstGrants.Fields.Item("ElementID").Value))) Then If (CStr(rstElements.Fields.Item("ElementID").Value) = CStr((rstGrants.Fields.Item("ElementID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstElements.Fields.Item("ElementName").Value)%></option>
                    <%
  rstElements.MoveNext()
Wend
If (rstElements.CursorType > 0) Then
  rstElements.MoveFirst
Else
  rstElements.Requery
End If
%>
                        </select>
                      </label>                    </td>
                    <td>&nbsp;</td>
                  </tr>
                  
                  <tr>
                    <td width="10">&nbsp;</td>
                    <td><strong>Level</strong></td>
  <td><label>
            <select name="cbxGrantLevelID" id="cbxGrantLevelID">
                        <%
While (NOT rstGrantLevels.EOF)
%><option value="<%=(rstGrantLevels.Fields.Item("GrantLevelID").Value)%>" <%If (Not isNull((rstGrants.Fields.Item("GrantLevelID").Value))) Then If (CStr(rstGrantLevels.Fields.Item("GrantLevelID").Value) = CStr((rstGrants.Fields.Item("GrantLevelID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstGrantLevels.Fields.Item("LevelName").Value)%></option>
                      <%
  rstGrantLevels.MoveNext()
Wend
If (rstGrantLevels.CursorType > 0) Then
  rstGrantLevels.MoveFirst
Else
  rstGrantLevels.Requery
End If
%>
                      </select>
                    </label></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="10">&nbsp;</td>
                    <td>&nbsp;</td>
                    <td><label>
                    <input type="submit" name="btnEdit" id="btnEdit" value="Update" />
                    </label></td>
                    <td>&nbsp;</td>
                  </tr>
                  <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
                  <input type="hidden" name="MM_update" value="frmEdit" />
                  <input type="hidden" name="MM_recordId" value="<%= rstGrants.Fields.Item("GrantID").Value %>" />
                  </form>
<%
End If
If bolSecurityDeleteGranted Then
%>                 
                  <tr>
                    <td width="10">&nbsp;</td>
                    <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
                      <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
                    <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
                      <input type="hidden" name="MM_delete" value="frmDelete" />
                      <input type="hidden" name="MM_recordId" value="<%= rstGrants.Fields.Item("GrantID").Value %>" />
                    </form>
                    </td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
<%
End If
%>                  
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
rstGrants.Close()
Set rstGrants = Nothing
%>
<%
rstUsers.Close()
Set rstUsers = Nothing
%>
<%
rstElements.Close()
Set rstElements = Nothing
%>
<%
rstGrantLevels.Close()
Set rstGrantLevels = Nothing
%>
