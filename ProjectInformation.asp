<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngProjectID
lngProjectID = Request.QueryString("lngProjectID")
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
    MM_editCmd.CommandText = "INSERT INTO dbo.ProjectDetails (DetailDescription, StartDate, ProjectStageID, DeveloperNotes, OwnerNotes, Priority, VendorID, ProjectID, BilledToClient) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxDetailDescription")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 135, 1, -1, MM_IIF(Request.Form("tbxStartDate"), Request.Form("tbxStartDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("cbxProjectStageID"), Request.Form("cbxProjectStageID"), 1)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 1000, Request.Form("tbxDeveloperNotes")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 1000, Request.Form("tbxOwnerNotes")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("tbxPriority"), Request.Form("tbxPriority"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("cbxVendorID"), Request.Form("cbxVendorID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("htbxProjectID"), Request.Form("htbxProjectID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_update")) = "frmStartTime") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    'Dim MM_editCmd
	dteNow = Now

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "UPDATE dbo.ProjectDetails SET StartTime = ? WHERE ProjectDetailID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 135, 1, -1, dteNow) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
If (CStr(Request("MM_update")) = "frmEndTime") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    'Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "UPDATE dbo.ProjectDetails SET StartTime = NULL WHERE ProjectDetailID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_insert")) = "frmAddWork") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    'Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.WorkHistorys (VendorID, WorkDate, ProjectDetailID, WorkDescription, Hours) VALUES (?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("cbxVendorID"), Request.Form("cbxVendorID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 135, 1, -1, MM_IIF(Request.Form("tbxWorkDate"), Request.Form("tbxWorkDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("cbxProjectDetailID"), Request.Form("cbxProjectDetailID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 1000, Request.Form("tbxWorkDescription")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("tbxHours"), Request.Form("tbxHours"), null)) ' adDouble
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
rstProjects_cmd.CommandText = "SELECT Projects.ProjectID, Projects.ClientID, Projects.ProjectDescription, Projects.StartDate, Projects.ProjectRate, Projects.ProjectPriority, Clients.ClientName FROM Clients INNER JOIN Projects ON Clients.ClientID = Projects.ClientID WHERE Projects.ProjectID = ?" 
rstProjects_cmd.Prepared = true
rstProjects_cmd.Parameters.Append rstProjects_cmd.CreateParameter("param1", 5, 1, -1, rstProjects__lngProjectID) ' adDouble

Set rstProjects = rstProjects_cmd.Execute
rstProjects_numRows = 0
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
rstProjectDetails_cmd.CommandText = "SELECT ProjectDetails.ProjectDetailID, ProjectDetails.ProjectID, ProjectDetails.ProjectStageID, ProjectDetails.DetailDescription, ProjectDetails.StartDate, ProjectDetails.Hours, ProjectDetails.StartTime, ProjectDetails.DeveloperNotes,  ProjectDetails.OwnerNotes, ProjectDetails.Priority, ProjectDetails.BilledToClient, ProjectStages.StageName, Vendors.VendorName, ProjectDetails.VendorID FROM ProjectDetails INNER JOIN ProjectStages ON ProjectDetails.ProjectStageID = ProjectStages.ProjectStageID LEFT OUTER JOIN Vendors ON ProjectDetails.VendorID = Vendors.VendorID WHERE ProjectID = ? ORDER BY Priority" 
rstProjectDetails_cmd.Prepared = true
rstProjectDetails_cmd.Parameters.Append rstProjectDetails_cmd.CreateParameter("param1", 5, 1, -1, rstProjectDetails__lngProjectID) ' adDouble

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
Dim rstWorkHistorys__lngProjectID
rstWorkHistorys__lngProjectID = "1"
If (lngProjectID <> "") Then 
  rstWorkHistorys__lngProjectID = lngProjectID
End If
%>
<%
Dim rstWorkHistorys
Dim rstWorkHistorys_cmd
Dim rstWorkHistorys_numRows

Set rstWorkHistorys_cmd = Server.CreateObject ("ADODB.Command")
rstWorkHistorys_cmd.ActiveConnection = MM_OBA_STRING
rstWorkHistorys_cmd.CommandText = "SELECT WorkHistorys.WorkHistoryID, WorkHistorys.ProjectDetailID, Vendors.VendorName, WorkHistorys.WorkDate, WorkHistorys.StartTime, WorkHistorys.Hours, WorkHistorys.WorkDescription, WorkHistorys.InvoiceID,  CASE WHEN InvoiceDetails.WorkHistoryID IS NULL THEN 'No' ELSE CAST(InvoiceDetails.InvoiceDate AS nvarchar)  END AS BilledToClient, ProjectDetails.DetailDescription FROM WorkHistorys INNER JOIN ProjectDetails ON WorkHistorys.ProjectDetailID = ProjectDetails.ProjectDetailID INNER JOIN Vendors ON WorkHistorys.VendorID = Vendors.VendorID LEFT OUTER JOIN (SELECT InvoiceDetails.WorkHistoryID, Invoices.InvoiceDate FROM Invoices INNER JOIN InvoiceDetails ON Invoices.InvoiceID = InvoiceDetails.InvoiceID) AS InvoiceDetails ON WorkHistorys.WorkHistoryID = InvoiceDetails.WorkHistoryID WHERE (ProjectDetails.ProjectID = ?) ORDER BY WorkHistorys.WorkDate DESC" 
rstWorkHistorys_cmd.Prepared = true
rstWorkHistorys_cmd.Parameters.Append rstWorkHistorys_cmd.CreateParameter("param1", 5, 1, -1, rstWorkHistorys__lngProjectID) ' adDouble

Set rstWorkHistorys = rstWorkHistorys_cmd.Execute
rstWorkHistorys_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr"><!-- InstanceBegin template="/Templates/Information.dwt" codeOutsideHTMLIsLocked="false" -->
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
<link rel="stylesheet" href="//code.jquery.com/ui/1.12.0/themes/base/jquery-ui.css">
<link href="SpryAssets/SpryValidationTextarea.css" rel="stylesheet" type="text/css" />
<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.0/jquery-ui.js"></script>
<script src="SpryAssets/SpryValidationTextarea.js" type="text/javascript"></script>

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
$(function() {
	$("#tbxWorkDate").datepicker({minDate: new Date(2012,8 - 1,1)});
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
	<h1><%=strPageTitle & " " & strSubTitle%></h1>	
<%
If bolProjectsViewGranted OR (bolClientOnlyViewGranted AND CLng(Session("ClientID")) = CLng(rstProjects.Fields.Item("ClientID").Value)) Then
	If bolProjectsEditGranted Then
		strEditLink = "<a href=""ProjectEdit.asp?lngProjectID=" & (rstProjects.Fields.Item("ProjectID").Value) & """>"
		strEndEditLink = "</a>&nbsp;"
	Else
		strEditLink = ""
		strEndEditLink = "&nbsp;"
	End If
%>
                <table border="0" cellspacing="0" cellpadding="0" class="info">
                  <tr>
                    <th align="left"><h4>&nbsp;</h4></th>
                    <th align="left"><h4>&nbsp;</h4></th>
                    <th align="left"><h4>&nbsp;</h4></th>
                    <th align="left"><h4>&nbsp;</h4></th>
                    <th align="left"><h4>&nbsp;</h4></th>
                    <th align="left"><h4>&nbsp;</h4></th>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Client</strong></td>
                    <td><a href="ClientInformation.asp?lngClientID=<%=(rstProjects.Fields.Item("ClientID").Value)%>"  class="info_link"><%=(rstProjects.Fields.Item("ClientName").Value)%></a></td>
                    <td align="right">&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Description</strong></td>
                    <td><%=strEditLink & (rstProjects.Fields.Item("ProjectDescription").Value) & strEndEditLink%></td>
                    <td align="right"><strong>Start</strong></td>
                    <td><%=strEditLink & (rstProjects.Fields.Item("StartDate").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Rate</strong></td>
                    <td><%=strEditLink & FormatCurrency(rstProjects.Fields.Item("ProjectRate").Value) & strEndEditLink%></td>
                    <td align="right"><strong>Priority</strong></td>
                    <td><%=strEditLink & (rstProjects.Fields.Item("ProjectPriority").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                </table>
    			<table border="0" cellspacing="0" cellpadding="0" class="box">
                   <tr>
                     <th colspan="10" align="left"><h3>Project Milestones</h3></th>
                   </tr>
                   <tr>
                    <th align="left"><h4>&nbsp;</h4></th>
                    <th align="left"><h4>Vendor</h4></th>
                    <th align="left"><h4>Description</h4></th>
                    <th align="left"><h4>Start</h4></th>
                    <th align="left"><h4>Stage</h4></th>
                    <th align="left"><h4>Dev Notes</h4></th>
                    <th align="left"><h4>Owner Notes</h4></th>
                    <th align="center"><h4>Priority</h4></th>
                    <th align="center"><h4>Start Time</h4></th>
                    <th align="left"><h4>&nbsp;</h4></th>
                  </tr>
<%
	If bolProjectsAddGranted Then

%>                  
                  <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
                  <tr>
                     <td>&nbsp;</td>
                     <td><select name="cbxVendorID" id="cbxVendorID">
                       <%
While (NOT rstVendors.EOF)
%>
                       <option value="<%=(rstVendors.Fields.Item("VendorID").Value)%>" <%If (Not isNull("5")) Then If (CStr(rstVendors.Fields.Item("VendorID").Value) = CStr("5")) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstVendors.Fields.Item("VendorName").Value)%></option>
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
                     <td><input name="tbxDetailDescription" type="text" id="tbxDetailDescription" size="50" maxlength="50" /></td>
                     <td><input name="tbxStartDate" type="text" id="tbxStartDate" size="11" value="<%=Date%>" /></td>
                     <td><select name="cbxProjectStageID" id="cbxProjectStageID">
                       <%
While (NOT rstProjectStages.EOF)
%>
                       <option value="<%=(rstProjectStages.Fields.Item("ProjectStageID").Value)%>"><%=(rstProjectStages.Fields.Item("StageName").Value)%></option>
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
                     <td><textarea name="tbxDeveloperNotes" id="tbxDeveloperNotes" cols="45" rows="3"></textarea></td>
                     <td><textarea name="tbxOwnerNotes" id="tbxOwnerNotes" cols="45" rows="3"></textarea></td>
                     <td><input name="tbxPriority" type="text" id="tbxPriority" size="4" style="text-align:center" /></td>
                     <td><input type="submit" name="btnAdd2" id="btnAdd2" value="Add" />
                     <input name="htbxProjectID" type="hidden" id="htbxProjectID" value="<%=lngProjectID%>" /></td>
                     <td>&nbsp;</td>
      			  </tr>
                  <input type="hidden" name="MM_insert" value="frmAdd" />
                  </form>
                  <tr>
                    <td colspan="10"><hr /></td>
                  </tr>
<%
	End If
	
	Do While Not rstProjectDetails.EOF
		If bolProjectsEditGranted Then
			strEditLink = "<a href=""ProjectDetailEdit.asp?lngProjectDetailID=" & (rstProjectDetails.Fields.Item("ProjectDetailID").Value) & """>"
			strEndEditLink = "</a>&nbsp;"
		Else
			strEditLink = ""
			strEndEditLink = "&nbsp;"
		End If

		If rstProjectDetails.Fields.Item("StartTime").Value = "" OR IsNull(rstProjectDetails.Fields.Item("StartTime").Value) Then
			strStartFormName ="frmStartTime"
			dteStartTime = ""
		Else
			strStartFormName ="frmEndTime"
			dteStartTime = FormatDateTime(rstProjectDetails.Fields.Item("StartTime").Value, vbShortTime) & " EST"
		End If
		
	
		If bolClientOnlyViewGranted Then
			strClientEditLink = "<a href=""ProjectDetailEditClient.asp?lngProjectID=" & (rstProjects.Fields.Item("ProjectID").Value) & """>"
			strClientEndEditLink = "</a>&nbsp;"
			strVendorName = "N/A"
		Else
			strClientEditLink = ""
			strClientEndEditLink = "&nbsp;"
			strVendorName = (rstProjectDetails.Fields.Item("VendorName").Value)
		End If
		
		If (rstProjectDetails.Fields.Item("OwnerNotes").Value) <> "" Then
			strOwnerNotes = (rstProjectDetails.Fields.Item("OwnerNotes").Value)
		Else
			strOwnerNotes = "None Entered"
		End If
		
		
%>                  
                  <tr class="tr_hover">
                     <td>&nbsp;</td>
                     <td><%=strEditLink & strVendorName & strEndEditLink%></td>
                     <td><%=strEditLink & (rstProjectDetails.Fields.Item("DetailDescription").Value) & strEndEditLink%></td>

                     <td><%=strEditLink & (rstProjectDetails.Fields.Item("StartDate").Value) & strEndEditLink%></td>
                     <td><%=strEditLink & (rstProjectDetails.Fields.Item("StageName").Value) & strEndEditLink%></td>
                     <td><%=strEditLink & (rstProjectDetails.Fields.Item("DeveloperNotes").Value) & strEndEditLink%></td>
                     <td><%=strEditLink & strClientEditLink & strOwnerNotes & strEndEditLink & strClientEndEditLink%></td>
                     <td align="center"><%=strEditLink & strClientEditLink & (rstProjectDetails.Fields.Item("Priority").Value) & strEndEditLink & strClientEndEditLink%></td>
<%
		If bolClientOnlyViewGranted Then
%>                     
                     <td>&nbsp;</td>
<%
		Else
%>                     
                     <form id="frmStartTime" name="frmStartTime" method="POST" action="<%=MM_editAction%>">
                     <td align="center" nowrap="nowrap"><%=strEditLink & dteStartTime & strEndEditLink%>
                         <input type="submit" name="btnStartTime" id="btnStartTime" value="*" />
                     </td>
                     <input type="hidden" name="MM_update" value="<%=strStartFormName%>" />
                     <input type="hidden" name="MM_recordId" value="<%= rstProjectDetails.Fields.Item("ProjectDetailID").Value %>" />
                     </form>
<%		
		End If
%>                     
                     <td>&nbsp;</td>
	  </tr>
<%
		rstProjectDetails.MoveNext
	Loop
	If (rstProjectDetails.CursorType > 0) Then
	  rstProjectDetails.MoveFirst
	Else
	  rstProjectDetails.Requery
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
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                </table>
    			<table border="0" cellspacing="0" cellpadding="0" class="box">
                   <tr>
                     <th colspan="8" align="left"><h3>Work History</h3></th>
                   </tr>
                   <tr>
                    <th align="left"><h4>&nbsp;</h4></th>
                    <th align="left"><h4>Vendor</h4></th>
                    <th align="left"><h4>Work Date</h4></th>
                    <th align="left"><h4>Project Milestone</h4></th>
                    <th align="left"><h4>Work Description</h4></th>
                    <th align="center"><h4>Hours</h4></th>
                    <th align="center"><h4>Billed</h4></th>
                    <th align="left"><h4>&nbsp;</h4></th>
                  </tr>
<%
	If bolProjectsAddGranted Then

%>                  
                  <form id="frmAddWork" name="frmAddWork" method="POST" action="<%=MM_editAction%>">
                  <tr>
                     <td>&nbsp;</td>
                     <td><select name="cbxVendorID" id="cbxVendorID">
                       <%
While (NOT rstVendors.EOF)
%>
                       <option value="<%=(rstVendors.Fields.Item("VendorID").Value)%>" <%If (Not isNull("5")) Then If (CStr(rstVendors.Fields.Item("VendorID").Value) = CStr("5")) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstVendors.Fields.Item("VendorName").Value)%></option>
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
                     <td><input name="tbxWorkDate" type="text" id="tbxWorkDate" value="<%=Date%>" size="11" /></td>
                     <td><select name="cbxProjectDetailID" id="cbxProjectDetailID">
                       <%
	bolIncompleteMilestonesRemain = 0
	While (NOT rstProjectDetails.EOF)
		If rstProjectDetails.Fields.Item("StageName").Value <> "Complete" Then
			bolIncompleteMilestonesRemain = 1
%>
                       <option value="<%=(rstProjectDetails.Fields.Item("ProjectDetailID").Value)%>"><%=(rstProjectDetails.Fields.Item("DetailDescription").Value)%></option>
                       <%
		End If
		rstProjectDetails.MoveNext()
	Wend
	If bolIncompleteMilestonesRemain = 0 Then
		strDisableAddButton = " disabled=""disabled"""
	End IF
%>
                     </select></td>
                     <td><span id="sprytextarea1">
                     <textarea name="tbxWorkDescription" id="tbxWorkDescription" cols="75" rows="3"></textarea>
                    Characters Remaining:<span id="countsprytextarea1">&nbsp;</span><span class="textareaRequiredMsg">A value is required.</span><span class="textareaMaxCharsMsg">Exceeded maximum number of characters.</span></span></td>
                     <td align="center"><input name="tbxHours" type="text" id="tbxHours" style="text-align:center" value="0" size="6" /></td>
                     <td><input type="submit" name="btnAdd" id="btnAdd" value="Add"<%=strDisableAddButton%> /></td>
                     <td>&nbsp;</td>
      			  </tr>
                  <input type="hidden" name="MM_insert" value="frmAddWork" />
                  </form>
                  <tr>
                    <td colspan="8"><hr /></td>
                  </tr>
<%
	End If
	
	Do While Not rstWorkHistorys.EOF
		If bolProjectsEditGranted Then
			strEditLink = "<a href=""WorkHistoryEdit.asp?lngWorkHistoryID=" & (rstWorkHistorys.Fields.Item("WorkHistoryID").Value) & """>"
			strEndEditLink = "</a>&nbsp;"
		Else
			strEditLink = ""
			strEndEditLink = "&nbsp;"
		End If
%>                  
                  <tr class="tr_hover">
                     <td>&nbsp;</td>
                     <td><%=strEditLink & (rstWorkHistorys.Fields.Item("VendorName").Value) & strEndEditLink%></td>
                     <td><%=strEditLink & (rstWorkHistorys.Fields.Item("WorkDate").Value) & strEndEditLink%></td>
                     <td><%=strEditLink & (rstWorkHistorys.Fields.Item("DetailDescription").Value) & strEndEditLink%></td>
                     <td><%=strEditLink & (rstWorkHistorys.Fields.Item("WorkDescription").Value) & strEndEditLink%></td>
                    <td align="center"><%=strEditLink & (rstWorkHistorys.Fields.Item("Hours").Value) & strEndEditLink%></td>
                     <td align="center" nowrap="nowrap"><%=strEditLink & (rstWorkHistorys.Fields.Item("BilledToClient").Value) & strEndEditLink%></td>
                     <td>&nbsp;</td>
	  </tr>
<%
		rstWorkHistorys.MoveNext
	Loop
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
                
<%
Else
%>               

                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="12">Certain &quot;Projects&quot; permissions are required to view this information. ClientID = <%=Session("ClientID")%></td>
                  </tr>
<%
End If
%>
                </table>
    <script type="text/javascript">
var sprytextarea1 = new Spry.Widget.ValidationTextarea("sprytextarea1", {maxChars:1000, counterId:"countsprytextarea1", counterType:"chars_remaining"});
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
rstProjects.Close()
Set rstProjects = Nothing
%>
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
<%
rstWorkHistorys.Close()
Set rstWorkHistorys = Nothing
%>
