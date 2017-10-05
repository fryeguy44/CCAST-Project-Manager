<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngClientVitalID
Dim strReturnPath

lngClientVitalID = Request.QueryString("lngClientVitalID")
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
    MM_editCmd.CommandText = "UPDATE dbo.ClientVitals SET VitalName = ?, Username = ?, Password = ?, Address = ?, Notes = ? WHERE ClientVitalID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxClientVitalName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxUsername")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 100, Request.Form("tbxPassword")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("tbxAddress")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 1000, Request.Form("tbxNotes")) ' adVarWChar
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
    MM_editCmd.CommandText = "DELETE FROM dbo.ClientVitals WHERE ClientVitalID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If

End If
%>
<%
Dim rstClientVitals__lngClientVitalID
rstClientVitals__lngClientVitalID = "1"
If (lngClientVitalID <> "") Then 
  rstClientVitals__lngClientVitalID = lngClientVitalID
End If
%>
<%
Dim rstClientVitals
Dim rstClientVitals_cmd
Dim rstClientVitals_numRows

Set rstClientVitals_cmd = Server.CreateObject ("ADODB.Command")
rstClientVitals_cmd.ActiveConnection = MM_OBA_STRING
rstClientVitals_cmd.CommandText = "SELECT * FROM ClientVitals WHERE ClientVitalID = ?" 
rstClientVitals_cmd.Prepared = true
rstClientVitals_cmd.Parameters.Append rstClientVitals_cmd.CreateParameter("param1", 5, 1, -1, rstClientVitals__lngClientVitalID) ' adDouble

Set rstClientVitals = rstClientVitals_cmd.Execute
rstClientVitals_numRows = 0
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
	Response.Redirect("ClientVitals.asp")
End If

%>
<!-- jQuery UI -->
<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/themes/base/jquery-ui.css" />
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/jquery-ui.min.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
$(function() {
	$("#dteFromDate").datepicker();
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
If bolClientsEditGranted Then
	If rstClientVitals.EOF Then
%>  
        <tr>
          <th colspan="4">&nbsp;</th>
        </tr>
        <tr>
            <td colspan="4"><a href="<%=strReturnPath%>">The Client Vital record you are attempting to edit has been deleted. Click here to return to the Client Information page</a></td>
        </tr>
<%
	Else
%>     
    	<form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
        <tr>
            <td width="10">&nbsp;</td>
            <td align="right"><strong> Name</strong></td>
          <td><input name="tbxClientVitalName" type="text" id="tbxClientVitalName" value="<%=(rstClientVitals.Fields.Item("VitalName").Value)%>" /></td>
		<td>&nbsp;</td>
		</tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>User</strong></td>
          <td><input name="tbxUsername" type="text" id="tbxUsername" value="<%=(rstClientVitals.Fields.Item("Username").Value)%>" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Password</strong></td>
          <td><input name="tbxPassword" type="text" id="tbxPassword" value="<%=(rstClientVitals.Fields.Item("Password").Value)%>" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Address</strong></td>
          <td><input name="tbxAddress" type="text" id="tbxAddress" value="<%=(rstClientVitals.Fields.Item("Address").Value)%>" size="35" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Notes</strong></td>
          <td><textarea name="tbxNotes" cols="50" rows="2" id="tbxNotes"><%=(rstClientVitals.Fields.Item("Notes").Value)%></textarea></td>
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
        <input type="hidden" name="MM_recordId" value="<%= rstClientVitals.Fields.Item("ClientVitalID").Value %>" />
        </form>
<%
		If bolClientsDeleteGranted Then
%>                
      <tr>
        <td width="10">&nbsp;</td>
            <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
              <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
              <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
              <input type="hidden" name="MM_delete" value="frmDelete" />
              <input type="hidden" name="MM_recordId" value="<%= rstClientVitals.Fields.Item("ClientVitalID").Value %>" />
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
            <td colspan="4">Certain &quot;Clients&quot; permissions are required to perform this task.</td>
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
rstClientVitals.Close()
Set rstClientVitals = Nothing
%>
<%
rstClientVitals.Close()
Set rstClientVitals = Nothing
%>
