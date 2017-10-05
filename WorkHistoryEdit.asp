<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngWorkHistoryID
Dim strReturnPath

lngWorkHistoryID = Request.QueryString("lngWorkHistoryID")
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
    MM_editCmd.CommandText = "UPDATE dbo.WorkHistorys SET StartTime = ?, VendorID = ?, WorkDate = ?, ProjectDetailID = ?, WorkDescription = ?, Hours = ? WHERE WorkHistoryID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 135, 1, -1, MM_IIF(Request.Form("tbxWorkHistoryName"), Request.Form("tbxWorkHistoryName"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxVendorID"), Request.Form("cbxVendorID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 135, 1, -1, MM_IIF(Request.Form("tbxWorkDate"), Request.Form("tbxWorkDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("cbxProjectDetailID"), Request.Form("cbxProjectDetailID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 1000, Request.Form("tbxWorkDescription")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("tbxHours"), Request.Form("tbxHours"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
    MM_editCmd.CommandText = "DELETE FROM dbo.WorkHistorys WHERE WorkHistoryID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If

End If
%>
<%
Dim rstWorkHistorys__lngWorkHistoryID
rstWorkHistorys__lngWorkHistoryID = "1"
If (lngWorkHistoryID <> "") Then 
  rstWorkHistorys__lngWorkHistoryID = lngWorkHistoryID
End If
%>
<%
Dim rstWorkHistorys
Dim rstWorkHistorys_cmd
Dim rstWorkHistorys_numRows

Set rstWorkHistorys_cmd = Server.CreateObject ("ADODB.Command")
rstWorkHistorys_cmd.ActiveConnection = MM_OBA_STRING
rstWorkHistorys_cmd.CommandText = "SELECT WorkHistorys.WorkHistoryID, WorkHistorys.ProjectDetailID, WorkHistorys.VendorID, WorkHistorys.WorkDate, WorkHistorys.StartTime, WorkHistorys.Hours, WorkHistorys.WorkDescription, WorkHistorys.InvoiceID, ProjectDetails.ProjectID FROM WorkHistorys INNER JOIN ProjectDetails ON WorkHistorys.ProjectDetailID = ProjectDetails.ProjectDetailID WHERE WorkHistoryID = ?" 
rstWorkHistorys_cmd.Prepared = true
rstWorkHistorys_cmd.Parameters.Append rstWorkHistorys_cmd.CreateParameter("param1", 5, 1, -1, rstWorkHistorys__lngWorkHistoryID) ' adDouble

Set rstWorkHistorys = rstWorkHistorys_cmd.Execute
rstWorkHistorys_numRows = 0
%>
<%
If Not rstWorkHistorys.EOF Then
	lngProjectID = rstWorkHistorys.Fields.Item("ProjectID").Value
End If
%>
<%
Dim rstProjectDetails__lngProjectID
rstProjectDetails__lngProjectID = "1"
If (lngProjectID <> "") Then 
  rstProjectDetails__lngProjectID = lngProjectID
End If
%>
<%
Dim rstProjectDetails
Dim rstProjectDetails_cmd
Dim rstProjectDetails_numRows

Set rstProjectDetails_cmd = Server.CreateObject ("ADODB.Command")
rstProjectDetails_cmd.ActiveConnection = MM_OBA_STRING
rstProjectDetails_cmd.CommandText = "SELECT ProjectDetails.ProjectDetailID, ProjectDetails.DetailDescription FROM ProjectDetails WHERE (ProjectDetails.ProjectID = ?)  ORDER BY ProjectDetails.Priority" 
rstProjectDetails_cmd.Prepared = true
rstProjectDetails_cmd.Parameters.Append rstProjectDetails_cmd.CreateParameter("param1", 5, 1, -1, rstProjectDetails__lngProjectID) ' adDouble

Set rstProjectDetails = rstProjectDetails_cmd.Execute
rstProjectDetails_numRows = 0
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
If (CStr(Request("MM_update")) = "frmEdit") OR (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then
	Response.Redirect(Request.Form("htbxReturnPath"))
End If

%>
<!-- jQuery UI -->
<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/themes/base/jquery-ui.css" />
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/jquery-ui.min.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
$(function() {
	$("#tbxWorkDate").datepicker();
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
	If rstWorkHistorys.EOF Then
%>  
        <tr>
          <th colspan="4">&nbsp;</th>
        </tr>
        <tr>
            <td colspan="4"><a href="WorkHistorys.asp">The Work History record you are attempting to edit has been deleted. Click here to return to the Work History List page</a></td>
        </tr>
<%
	Else
%>     
    	<form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
        <tr>
            <td width="10">&nbsp;</td>
            <td align="right"><strong>Start Time</strong></td>
          <td><input name="tbxWorkHistoryName" type="text" id="tbxWorkHistoryName" value="<%=(rstWorkHistorys.Fields.Item("StartTime").Value)%>" /></td>
		<td>&nbsp;</td>
		</tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Vendor</strong></td>
          <td><select name="cbxVendorID" id="cbxVendorID">
            <%
While (NOT rstVendors.EOF)
%>
            <option value="<%=(rstVendors.Fields.Item("VendorID").Value)%>" <%If (Not isNull((rstWorkHistorys.Fields.Item("VendorID").Value))) Then If (CStr(rstVendors.Fields.Item("VendorID").Value) = CStr((rstWorkHistorys.Fields.Item("VendorID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstVendors.Fields.Item("VendorName").Value)%></option>
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
          <td>&nbsp;</td>
          <td align="right"><strong>Work Date</strong></td>
          <td><input name="tbxWorkDate" type="text" id="tbxWorkDate" value="<%=(rstWorkHistorys.Fields.Item("WorkDate").Value)%>" size="11" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Milestone</strong></td>
          <td><select name="cbxProjectDetailID" id="cbxProjectDetailID">
            <%
While (NOT rstProjectDetails.EOF)
%>
            <option value="<%=(rstProjectDetails.Fields.Item("ProjectDetailID").Value)%>" <%If (Not isNull((rstWorkHistorys.Fields.Item("ProjectDetailID").Value))) Then If (CStr(rstProjectDetails.Fields.Item("ProjectDetailID").Value) = CStr((rstWorkHistorys.Fields.Item("ProjectDetailID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstProjectDetails.Fields.Item("DetailDescription").Value)%></option>
            <%
  rstProjectDetails.MoveNext()
Wend
If (rstProjectDetails.CursorType > 0) Then
  rstProjectDetails.MoveFirst
Else
  rstProjectDetails.Requery
End If
%>
          </select></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Work Description</strong></td>
          <td><textarea name="tbxWorkDescription" id="tbxWorkDescription" cols="75" rows="3"><%=(rstWorkHistorys.Fields.Item("WorkDescription").Value)%></textarea></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Hours</strong></td>
          <td><input name="tbxHours" type="text" id="tbxHours" style="text-align:center" value="<%=(rstWorkHistorys.Fields.Item("Hours").Value)%>" size="6" /></td>
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
        <input type="hidden" name="MM_recordId" value="<%= rstWorkHistorys.Fields.Item("WorkHistoryID").Value %>" />
        </form>
<%
		If bolProjectsDeleteGranted Then
%>                
      <tr>
        <td width="10">&nbsp;</td>
            <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
              <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
              <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
              <input type="hidden" name="MM_delete" value="frmDelete" />
              <input type="hidden" name="MM_recordId" value="<%= rstWorkHistorys.Fields.Item("WorkHistoryID").Value %>" />
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
rstWorkHistorys.Close()
Set rstWorkHistorys = Nothing
%>
<%
rstProjectDetails.Close()
Set rstProjectDetails = Nothing
%>
<%
rstVendors.Close()
Set rstVendors = Nothing
%>
