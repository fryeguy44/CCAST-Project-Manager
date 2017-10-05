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
If (CStr(Request("MM_update")) = "frmEdit") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "UPDATE dbo.ProjectDetails SET VendorID = ?, DetailDescription = ?, StartDate = ?, ProjectStageID = ?, DeveloperNotes = ?, OwnerNotes = ?, Priority = ?, StartTime = ? WHERE ProjectDetailID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("cbxVendorID"), Request.Form("cbxVendorID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxDetailDescription")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 135, 1, -1, MM_IIF(Request.Form("tbxStartDate"), Request.Form("tbxStartDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("cbxProjectStageID"), Request.Form("cbxProjectStageID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 1000, Request.Form("tbxDeveloperNotes")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 1000, Request.Form("tbxOwnerNotes")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("tbxPriority"), Request.Form("tbxPriority"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 135, 1, -1, MM_IIF(Request.Form("tbxStartTime"), Request.Form("tbxStartTime"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
    MM_editCmd.CommandText = "DELETE FROM dbo.ProjectDetails WHERE ProjectDetailID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If

End If
%>
<%
Dim lngProjectDetailID
Dim strReturnPath

lngProjectDetailID = Request.QueryString("lngProjectDetailID")
If Request.QueryString("strReturnPath") = "" Then
	strReturnPath = Request.ServerVariables("HTTP_REFERER")
Else
	strReturnPath = Request.QueryString("strReturnPath")
End If
%>
<%
Dim rstProjectDetails__lngProjectDetailID
rstProjectDetails__lngProjectDetailID = "1"
If (lngProjectDetailID <> "") Then 
  rstProjectDetails__lngProjectDetailID = lngProjectDetailID
End If
%>
<%
Dim rstProjectDetails
Dim rstProjectDetails_cmd
Dim rstProjectDetails_numRows

Set rstProjectDetails_cmd = Server.CreateObject ("ADODB.Command")
rstProjectDetails_cmd.ActiveConnection = MM_OBA_STRING
rstProjectDetails_cmd.CommandText = "SELECT TOP (1) ProjectDetails.ProjectDetailID, ProjectDetails.ProjectID, ProjectDetails.ProjectStageID, ProjectDetails.DetailDescription, ProjectDetails.StartDate, ProjectDetails.Hours, ProjectDetails.StartTime,  ProjectDetails.DeveloperNotes, ProjectDetails.OwnerNotes, ProjectDetails.Priority, ProjectDetails.BilledToClient, ProjectDetails.VendorID,  WorkHistorys.WorkHistoryID FROM ProjectDetails LEFT OUTER JOIN WorkHistorys ON ProjectDetails.ProjectDetailID = WorkHistorys.ProjectDetailID WHERE ProjectDetails.ProjectDetailID = ?" 
rstProjectDetails_cmd.Prepared = true
rstProjectDetails_cmd.Parameters.Append rstProjectDetails_cmd.CreateParameter("param1", 5, 1, -1, rstProjectDetails__lngProjectDetailID) ' adDouble

Set rstProjectDetails = rstProjectDetails_cmd.Execute
rstProjectDetails_numRows = 0
%>
<%
Dim rstProjectStages
Dim rstProjectStages_cmd
Dim rstProjectStages_numRows

Set rstProjectStages_cmd = Server.CreateObject ("ADODB.Command")
rstProjectStages_cmd.ActiveConnection = MM_OBA_STRING
rstProjectStages_cmd.CommandText = "SELECT * FROM ProjectStages ORDER BY SortOrder" 
rstProjectStages_cmd.Prepared = true

Set rstProjectStages = rstProjectStages_cmd.Execute
rstProjectStages_numRows = 0
%>
<%
Dim rstVendors
Dim rstVendors_cmd
Dim rstVendors_numRows

Set rstVendors_cmd = Server.CreateObject ("ADODB.Command")
rstVendors_cmd.ActiveConnection = MM_OBA_STRING
rstVendors_cmd.CommandText = "SELECT VendorID, VendorName FROM Vendors ORDER BY VendorName" 
rstVendors_cmd.Prepared = true

Set rstVendors = rstVendors_cmd.Execute
rstVendors_numRows = 0
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
	Response.Redirect("ProjectsCurrent.asp")
End If

%>
<!-- jQuery UI -->
<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/themes/base/jquery-ui.css">
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/jquery-ui.min.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
function addZero(i) {
    if (i < 10) {
        i = "0" + i;
    }
    return i;
}
$(function() {
	$("#tbxStartDate").datepicker({minDate: new Date(2012,8 - 1,1)});
});
function AddTime() {
    var d = new Date();
    $("#tbxStartTime").val($.datepicker.formatDate('mm/dd/yy', new Date()) + ' ' + addZero(d.getHours()) + ':' + addZero(d.getMinutes()));
     } 
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
	If rstProjectDetails.EOF Then
%>  
        <tr>
          <th colspan="4">&nbsp;</th>
        </tr>
        <tr>
            <td colspan="4"><a href="ProjectDetails.asp">The ProjectDetail you are attempting to edit has been deleted. Click here to return to the Project Detail List page</a></td>
        </tr>
<%
	Else
%>     
    	<form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Vendor</strong></td>
          <td><select name="cbxVendorID" id="cbxVendorID">
            <%
While (NOT rstVendors.EOF)
%>
            <option value="<%=(rstVendors.Fields.Item("VendorID").Value)%>" <%If (Not isNull((rstProjectDetails.Fields.Item("VendorID").Value))) Then If (CStr(rstVendors.Fields.Item("VendorID").Value) = CStr((rstProjectDetails.Fields.Item("VendorID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstVendors.Fields.Item("VendorName").Value)%></option>
            <%
  rstVendors.MoveNext()
Wend
If (rstVendors.CursorType > 0) Then
  rstVendors.MoveFirst
Else
  rstVendors.Requery
End If
%>
          </select></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
            <td width="10">&nbsp;</td>
            <td align="right"><strong>Detail Description</strong></td>
          <td><input name="tbxDetailDescription" type="text" id="tbxDetailDescription" value="<%=(rstProjectDetails.Fields.Item("DetailDescription").Value)%>" size="50" maxlength="50" /></td>
		<td>&nbsp;</td>
		</tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Start</strong></td>
          <td><input name="tbxStartDate" type="text" id="tbxStartDate" value="<%=(rstProjectDetails.Fields.Item("StartDate").Value)%>" size="11" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong> Stage</strong></td>
          <td><select name="cbxProjectStageID" id="cbxProjectStageID">
            <%
While (NOT rstProjectStages.EOF)
%>
            <option value="<%=(rstProjectStages.Fields.Item("ProjectStageID").Value)%>" <%If (Not isNull((rstProjectDetails.Fields.Item("ProjectStageID").Value))) Then If (CStr(rstProjectStages.Fields.Item("ProjectStageID").Value) = CStr((rstProjectDetails.Fields.Item("ProjectStageID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstProjectStages.Fields.Item("StageName").Value)%></option>
            <%
  rstProjectStages.MoveNext()
Wend
If (rstProjectStages.CursorType > 0) Then
  rstProjectStages.MoveFirst
Else
  rstProjectStages.Requery
End If
%>
          </select></td>
          <td>&nbsp;</td>
        </tr>
<%        
        If bolDeveloperEditGranted Then
			strReadOnly = " "
		Else
			strReadOnly = "readonly "
		End If
		If IsDate(rstProjectDetails.Fields.Item("StartTime").Value) Then
			strStartTime = FormatDateTime(rstProjectDetails.Fields.Item("StartTime").Value, 2) & " " & FormatDateTime(rstProjectDetails.Fields.Item("StartTime").Value, 4)
		Else 
			strStartTime = ""
		End If
%>		
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Dev Notes</strong></td>
          <td><textarea name="tbxDeveloperNotes" id="tbxDeveloperNotes" cols="45" rows="3"<%=strReadOnly%>><%=(rstProjectDetails.Fields.Item("DeveloperNotes").Value)%></textarea></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Owner Notes</strong></td>
          <td><textarea name="tbxOwnerNotes" id="tbxOwnerNotes" cols="45" rows="3"><%=(rstProjectDetails.Fields.Item("OwnerNotes").Value)%></textarea></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Priority</strong></td>
          <td><input name="tbxPriority" type="text" id="tbxPriority" value="<%=(rstProjectDetails.Fields.Item("Priority").Value)%>" size="4" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Start Time</strong></td>
          <td><input name="tbxStartTime" type="text" id="tbxStartTime" value="<%=strStartTime%>" size="20" onclick=" AddTime(); " /></td>
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
        <input type="hidden" name="MM_recordId" value="<%= rstProjectDetails.Fields.Item("ProjectDetailID").Value %>" />
        </form>
<%
		If bolDeveloperDeleteGranted AND IsNull(rstProjectDetails.Fields.Item("WorkHistoryID").Value) Then
%>                
      <tr>
        <td width="10">&nbsp;</td>
            <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
              <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
              <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
              <input type="hidden" name="MM_delete" value="frmDelete" />
              <input type="hidden" name="MM_recordId" value="<%= rstProjectDetails.Fields.Item("ProjectDetailID").Value %>" />
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
rstProjectDetails.Close()
Set rstProjectDetails = Nothing
%>
<%
rstProjectStages.Close()
Set rstProjectStages = Nothing
%>
<%
rstVendors.Close()
Set rstVendors = Nothing
%>
