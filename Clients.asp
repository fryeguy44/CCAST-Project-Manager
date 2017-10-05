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
    MM_editCmd.CommandText = "INSERT INTO dbo.Clients (ClientName, Source, CurrentRate, Skype, Teamviewer, Phone, Email, Notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxClientName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxSource")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("tbxCurrentRate"), Request.Form("tbxCurrentRate"), 45)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("tbxSkype")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("tbxTeamviewer"), Request.Form("tbxTeamviewer"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 20, Request.Form("tbxPhone")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 202, 1, 50, Request.Form("tbxEmail")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 202, 1, 1000, Request.Form("tbxNotes")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim rstClients
Dim rstClients_cmd
Dim rstClients_numRows

Set rstClients_cmd = Server.CreateObject ("ADODB.Command")
rstClients_cmd.ActiveConnection = MM_OBA_STRING
rstClients_cmd.CommandText = "SELECT * FROM Clients ORDER BY ClientName" 
rstClients_cmd.Prepared = true

Set rstClients = rstClients_cmd.Execute
rstClients_numRows = 0
%>
<%
Dim rstWorkHistory
Dim rstWorkHistory_cmd
Dim rstWorkHistory_numRows

Set rstWorkHistory_cmd = Server.CreateObject ("ADODB.Command")
rstWorkHistory_cmd.ActiveConnection = MM_OBA_STRING
rstWorkHistory_cmd.CommandText = "SELECT Clients.ClientName, WorkHistorys.WorkHistoryID, WorkHistorys.ProjectDetailID, Vendors.VendorName, WorkHistorys.WorkDate, WorkHistorys.Hours, LEFT(WorkHistorys.WorkDescription, 150) AS WorkDescription, ProjectDetails.DetailDescription,  Clients.CurrentRate FROM WorkHistorys INNER JOIN ProjectDetails ON WorkHistorys.ProjectDetailID = ProjectDetails.ProjectDetailID INNER JOIN Vendors ON WorkHistorys.VendorID = Vendors.VendorID INNER JOIN Projects ON ProjectDetails.ProjectID = Projects.ProjectID INNER JOIN Clients ON Projects.ClientID = Clients.ClientID LEFT OUTER JOIN InvoiceDetails ON WorkHistorys.WorkHistoryID = InvoiceDetails.WorkHistoryID WHERE (InvoiceDetails.InvoiceDetailID IS NULL) ORDER BY ClientName, WorkDate" 
rstWorkHistory_cmd.Prepared = true

Set rstWorkHistory = rstWorkHistory_cmd.Execute
rstWorkHistory_numRows = 0
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
	<h1>Client List</h1>	
	<table border="0" cellspacing="0" cellpadding="0" class="fluid">
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="left"><h4>Name</h4></th>
	    <th align="left"><h4>Source</h4></th>
	    <th align="right"><h4>Current Rate</h4></th>
	    <th align="left"><h4>Skype</h4></th>
	    <th align="left"><h4>Teamviewer</h4></th>
	    <th align="left"><h4>Phone</h4></th>
	    <th align="left"><h4>Email</h4></th>
	    <th align="left"><h4>Notes</h4></th>
	    <th align="left"><h4>Entered</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
<%
If bolClientsViewGranted Then
    If bolClientsAddGranted Then
%>
      <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
	  <tr>
	    <td>&nbsp;</td>
	    <td><input type="text" name="tbxClientName" id="tbxClientName" tabindex="0" /></td>
	    <td><input type="text" name="tbxSource" id="tbxSource" /></td>
	    <td align="right">$
	      <input name="tbxCurrentRate" type="text" id="tbxCurrentRate" value="45" size="5" style="text-align:right" /></td>
	    <td><input name="tbxSkype" type="text" id="tbxSkype" size="35" /></td>
	    <td><input name="tbxTeamviewer" type="text" id="tbxTeamviewer" size="10" /></td>
	    <td><input name="tbxPhone" type="text" id="tbxPhone" size="15" /></td>
	    <td><input name="tbxEmail" type="text" id="tbxEmail" size="35" /></td>
	    <td><input type="text" name="tbxNotes" id="tbxNotes" /></td>
	    <td><input type="submit" name="btnAdd" id="btnAdd" value="Add Client" /></td>
	    <td>&nbsp;</td>
      </tr>
      <input type="hidden" name="MM_insert" value="frmAdd" />
      </form>
	  <tr>
	    <td colspan="11"><hr /></td>
      </tr>
<%
    End If
	Do While Not rstClients.EOF
		If bolClientsEditGranted Then
			strEdit = "<a href=""ClientEdit.asp?lngClientID=" & (rstClients.Fields.Item("ClientID").Value) & """>"
			strEditEnd = "</a>"
		Else
			strEdit = ""
			strEditEnd = ""
		End If 
%>      
	  <tr class="tr_hover">
        <td><a href="ClientInformation.asp?lngClientID=<%=(rstClients.Fields.Item("ClientID").Value)%>" class="row_info"></a></td>
		<td><%=strEdit & (rstClients.Fields.Item("ClientName").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstClients.Fields.Item("Source").Value) & strEditEnd%></td>
		<td align="right"><%=strEdit & FormatCurrency(rstClients.Fields.Item("CurrentRate").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstClients.Fields.Item("Skype").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstClients.Fields.Item("Teamviewer").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstClients.Fields.Item("Phone").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstClients.Fields.Item("Email").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstClients.Fields.Item("Notes").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstClients.Fields.Item("DateEntered").Value) & strEditEnd%></td>
		<td>&nbsp;</td>
	  </tr>
<%
        rstClients.MoveNext
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
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	  </tr>
	</table>
	<table border="0" cellspacing="0" cellpadding="0" class="box">
	  <tr>
	    <th colspan="9"><h2>Unbilled Work History</h2></th>
      </tr>
	  <tr>
	    <th>&nbsp;</th>
		<th align="left"><h4>Client</h4></th>
		<th align="left"><h4>Vendor</h4></th>
		<th align="left"><h4>Date</h4></th>
		<th align="left"><h4>Description</h4></th>
		<th align="left"><h4>Milestone</h4></th>
		<th align="right"><h4>Hours</h4></th>
		<th align="right"><h4>Amount</h4></th>
		<th>&nbsp;</th>
	  </tr>
<%
	If bolInvoicesViewGranted Then
		dblTotalTime = 0
		curTotalAmount = 0
		Do While Not rstWorkHistory.EOF
			If bolInvoicesEditGranted Then
				strEdit = "<a href=""WorkHistoryEdit.asp?lngWorkHistoryID=" & (rstWorkHistory.Fields.Item("WorkHistoryID").Value) & """>"
				strEditEnd = "</a>"
			Else
				strEdit = ""
				strEditEnd = ""
			End If 
		
			dblTotalTime = dblTotalTime + CDbl(rstWorkHistory.Fields.Item("Hours").Value)
			curTotalAmount = curTotalAmount + CDbl(rstWorkHistory.Fields.Item("Hours").Value) * CDbl(rstWorkHistory.Fields.Item("CurrentRate").Value)
%>      
	  <tr class="tr_hover">
	    <td>&nbsp;</td>
	    <td nowrap="nowrap"><%=strEdit & (rstWorkHistory.Fields.Item("ClientName").Value) & strEditEnd%></td>
	    <td nowrap="nowrap"><%=strEdit & (rstWorkHistory.Fields.Item("VendorName").Value) & strEditEnd%></td>
	    <td><%=strEdit & (rstWorkHistory.Fields.Item("WorkDate").Value) & strEditEnd%></td>
	    <td nowrap="nowrap"><%=strEdit & (rstWorkHistory.Fields.Item("WorkDescription").Value) & strEditEnd%></td>
	    <td nowrap="nowrap"><%=strEdit & (rstWorkHistory.Fields.Item("DetailDescription").Value) & strEditEnd%></td>
	    <td align="right"><%=strEdit & FormatNumber(rstWorkHistory.Fields.Item("Hours").Value, 1) & strEditEnd%></td>
	    <td align="right"><%=FormatCurrency((rstWorkHistory.Fields.Item("Hours").Value) * (rstWorkHistory.Fields.Item("CurrentRate").Value))%></td>
	    <td>&nbsp;</td>
      </tr>
<%
			rstWorkHistory.MoveNext
		Loop
	End If
%>      
	  <tr>
	    <td colspan="9"><hr /></td>
      </tr>
	  <tr>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td align="right"><strong>Totals:</strong></td>
	    <td align="right"><%=dblTotalTime%></td>
	    <td align="right"><%=FormatCurrency(curTotalAmount)%></td>
	    <td>&nbsp;</td>
      </tr>
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
	  </tr>


<%	
	
Else
%>  
        <tr>
            <td colspan="9">Viewing this list requires certain &quot;Clients&quot; permissions</td>
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
rstClients.Close()
Set rstClients = Nothing
%>
<%
rstWorkHistory.Close()
Set rstWorkHistory = Nothing
%>
