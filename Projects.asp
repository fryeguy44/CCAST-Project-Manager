<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
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
    MM_editCmd.CommandText = "INSERT INTO dbo.Projects (ClientID, ProjectDescription, StartDate, ProjectRate, ProjectPriority) VALUES (?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("cbxClientID"), Request.Form("cbxClientID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxProjectDescription")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 135, 1, -1, MM_IIF(Request.Form("tbxStartDate"), Request.Form("tbxStartDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxProjectRate"), Request.Form("tbxProjectRate"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("tbxProjectPriority"), Request.Form("tbxProjectPriority"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim rstProjects
Dim rstProjects_cmd
Dim rstProjects_numRows

Set rstProjects_cmd = Server.CreateObject ("ADODB.Command")
rstProjects_cmd.ActiveConnection = MM_OBA_STRING
rstProjects_cmd.CommandText = "SELECT Projects.ProjectID, Projects.ClientID, Projects.ProjectDescription, Projects.StartDate, Projects.ProjectRate, Projects.ProjectPriority, Clients.ClientName FROM Clients INNER JOIN Projects ON Clients.ClientID = Projects.ClientID ORDER BY Projects.ProjectPriority" 
rstProjects_cmd.Prepared = true

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
<!-- jQuery UI -->
<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/themes/base/jquery-ui.css" />
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/jquery-ui.min.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
$(function() {
	$("#tbxStartDate").datepicker({ minDate:'-30D'});
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
	<h1>Project List</h1>	
	<table border="0" cellspacing="0" cellpadding="0" class="box">
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="left"><h4>Client</h4></th>
	    <th align="left"><h4>Description</h4></th>
	    <th align="center"><h4>Start</h4></th>
	    <th align="right"><h4>Rate</h4></th>
	    <th align="center"><h4>Priority</h4></th>
	    <th align="center"><h4>Project ID</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
<%
If bolProjectsViewGranted Then
    If bolProjectsAddGranted Then
%>
      <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
	  <tr>
	    <td>&nbsp;</td>
	    <td><select name="cbxClientID" id="cbxClientID">
	      <%
While (NOT rstClients.EOF)
%>
	      <option value="<%=(rstClients.Fields.Item("ClientID").Value)%>"><%=(rstClients.Fields.Item("ClientName").Value)%></option>
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
	    <td><input name="tbxProjectDescription" type="text" id="tbxProjectDescription" tabindex="0" size="50" maxlength="50" /></td>
	    <td align="center"><input name="tbxStartDate" type="text" id="tbxStartDate" value="<%=Date%>" size="11" style="text-align: center" /></td>
	    <td align="right"><input name="tbxProjectRate" type="text" id="tbxProjectRate" size="8" style="text-align: right" /></td>
	    <td align="center"><input name="tbxProjectPriority" type="text" id="tbxProjectPriority" size="5" style="text-align: center" /></td>
	    <td align="center"><input type="submit" name="btnAdd" id="btnAdd" value="Add Project" /></td>
	    <td>&nbsp;</td>
      </tr>
      <input type="hidden" name="MM_insert" value="frmAdd" />
      </form>
	  <tr>
	    <td colspan="8"><hr /></td>
      </tr>
<%
    End If
	Do While Not rstProjects.EOF
		If bolProjectsEditGranted Then
			strEdit = "<a href=""ProjectEdit.asp?lngProjectID=" & (rstProjects.Fields.Item("ProjectID").Value) & """>"
			strEditEnd = "</a>"
		Else
			strEdit = ""
			strEditEnd = ""
		End If 
%>      
	  <tr>
        <td><a href="ProjectInformation.asp?lngProjectID=<%=(rstProjects.Fields.Item("ProjectID").Value)%>" class="row_info"></a></td>
        <td><%=(rstProjects.Fields.Item("ClientName").Value)%></td>
		<td><%=strEdit & (rstProjects.Fields.Item("ProjectDescription").Value) & strEditEnd%></td>
		<td align="center"><%=strEdit & (rstProjects.Fields.Item("StartDate").Value) & strEditEnd%></td>
		<td align="right"><%=strEdit & FormatCurrency(rstProjects.Fields.Item("ProjectRate").Value) & strEditEnd%></td>
		<td align="center"><%=strEdit & (rstProjects.Fields.Item("ProjectPriority").Value) & strEditEnd%></td>
		<td align="center"><%=(rstProjects.Fields.Item("ProjectID").Value)%></td>
		<td>&nbsp;</td>
	  </tr>
<%
        rstProjects.MoveNext
    Loop
Else
%>  
        <tr>
            <td colspan="8">Viewing this list requires certain &quot;Projects&quot; permissions</td>
        </tr>

<%
End If
%>
	  <tr>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
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
