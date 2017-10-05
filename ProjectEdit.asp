<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngProjectID
Dim strReturnPath

lngProjectID = Request.QueryString("lngProjectID")
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
    MM_editCmd.CommandText = "UPDATE dbo.Projects SET ClientID = ?, ProjectDescription = ?, StartDate = ?, ProjectRate = ?, ProjectPriority = ? WHERE ProjectID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("cbxClientID"), Request.Form("cbxClientID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxProjectDescription")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 135, 1, -1, MM_IIF(Request.Form("tbxStartDate"), Request.Form("tbxStartDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxProjectRate"), Request.Form("tbxProjectRate"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("tbxProjectPriority"), Request.Form("tbxProjectPriority"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "DELETE FROM dbo.Projects WHERE ProjectID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If

End If
%>
<%
Dim rstProjects__lngProjectID
rstProjects__lngProjectID = "1"
If (lngProjectID <> "") Then 
  rstProjects__lngProjectID = lngProjectID
End If
%>
<%
Dim rstProjects
Dim rstProjects_cmd
Dim rstProjects_numRows

Set rstProjects_cmd = Server.CreateObject ("ADODB.Command")
rstProjects_cmd.ActiveConnection = MM_OBA_STRING
rstProjects_cmd.CommandText = "SELECT TOP 1 Projects.ProjectID, Projects.ClientID, Projects.ProjectDescription, Projects.StartDate, Projects.ProjectRate, Projects.ProjectPriority, ProjectDetails.ProjectDetailID FROM Projects LEFT OUTER JOIN ProjectDetails ON Projects.ProjectID = ProjectDetails.ProjectID WHERE Projects.ProjectID = ?" 
rstProjects_cmd.Prepared = true
rstProjects_cmd.Parameters.Append rstProjects_cmd.CreateParameter("param1", 5, 1, -1, rstProjects__lngProjectID) ' adDouble

Set rstProjects = rstProjects_cmd.Execute
rstProjects_numRows = 0
%>
<%
Dim rstClients
Dim rstClients_cmd
Dim rstClients_numRows

Set rstClients_cmd = Server.CreateObject ("ADODB.Command")
rstClients_cmd.ActiveConnection = MM_OBA_STRING
rstClients_cmd.CommandText = "SELECT ClientID, ClientName FROM Clients ORDER BY ClientName" 
rstClients_cmd.Prepared = true

Set rstClients = rstClients_cmd.Execute
rstClients_numRows = 0
%>
<%
If (CStr(Request("MM_update")) = "frmEdit") Then
	lngAccessTypeID = 2
End If
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
If (CStr(Request("MM_update")) = "frmEdit") Then
	Response.Redirect(Request.Form("htbxReturnPath"))
End If
If (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then
	Response.Redirect("Projects.asp")
End If

%>
<!-- jQuery UI -->
<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/themes/base/jquery-ui.css" />
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/jquery-ui.min.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
$(function() {
	$("#tbxStartDate").datepicker({  maxDate: 0 });
});
</script>

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
If bolProjectsEditGranted Then
	If rstProjects.EOF Then
%>  
        <tr>
          <th colspan="4">&nbsp;</th>
        </tr>
        <tr>
            <td colspan="4"><a href="Projects.asp">The Project you are attempting to edit has been deleted. Click here to return to the Project List page</a></td>
        </tr>
<%
	Else
%>     
    	<form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Client</strong></td>
          <td><select name="cbxClientID" id="cbxClientID">
            <%
While (NOT rstClients.EOF)
%>
            <option value="<%=(rstClients.Fields.Item("ClientID").Value)%>" <%If (Not isNull((rstProjects.Fields.Item("ClientID").Value))) Then If (CStr(rstClients.Fields.Item("ClientID").Value) = CStr((rstProjects.Fields.Item("ClientID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstClients.Fields.Item("ClientName").Value)%></option>
            <%
  rstClients.MoveNext()
Wend
If (rstClients.CursorType > 0) Then
  rstClients.MoveFirst
Else
  rstClients.Requery
End If
%>
          </select></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
            <td width="10">&nbsp;</td>
            <td align="right"><strong>Project Description</strong></td>
          <td><input name="tbxProjectDescription" type="text" id="tbxProjectDescription" value="<%=(rstProjects.Fields.Item("ProjectDescription").Value)%>" size="50" maxlength="50" /></td>
		<td>&nbsp;</td>
		</tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Start</strong></td>
          <td><input name="tbxStartDate" type="text" id="tbxStartDate" value="<%=(rstProjects.Fields.Item("StartDate").Value)%>" size="11" /></td>
          <td>&nbsp;</td>
        </tr>
<%
		If bolDeveloperDeleteGranted Then
			strReadOnly = " "
		Else
			strReadOnly = "readonly "
		End If
%>                
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong> Rate</strong></td>
          <td><input name="tbxProjectRate" type="text" id="tbxProjectRate" value="<%=(rstProjects.Fields.Item("ProjectRate").Value)%>" size="8" <%=strReadOnly%>/></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Priority</strong></td>
          <td><input name="tbxProjectPriority" type="text" id="tbxProjectPriority" value="<%=(rstProjects.Fields.Item("ProjectPriority").Value)%>" size="5" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
            <td width="10">&nbsp;</td>
            <td>&nbsp;</td>
            <td><input type="submit" name="btnEdit" id="btnEdit" value="Update" /></td>
            <td>&nbsp;</td>
      </tr>
        <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
        <input type="hidden" name="MM_update" value="frmEdit" />
        <input type="hidden" name="MM_recordId" value="<%= rstProjects.Fields.Item("ProjectID").Value %>" />
        </form>
<%
		If bolDeveloperDeleteGranted AND IsNull(rstProjects.Fields.Item("ProjectDetailID").Value) Then
%>                
      <tr>
        <td width="10">&nbsp;</td>
            <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
              <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
              <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
              <input type="hidden" name="MM_delete" value="frmDelete" />
              <input type="hidden" name="MM_recordId" value="<%= rstProjects.Fields.Item("ProjectID").Value %>" />
            </form>            </td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
      </tr>
<%
		End If
	End If
Else
%>
        <tr>
            <td colspan="4">Certain &quot;Projects&quot; permissions are required to perform this task.</td>
        </tr>

<%

End If
%>
        <tr>
            <td width="10">&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
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
rstProjects.Close()
Set rstProjects = Nothing
%>
<%
rstClients.Close()
Set rstClients = Nothing
%>