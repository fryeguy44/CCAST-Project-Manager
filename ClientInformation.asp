<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngClientID
lngClientID = Request.QueryString("lngClientID")
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
If (CStr(Request("MM_insert")) = "frmAddPayment") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.ClientPayments (PaymentDate, PaymentMethodID, Amount, CreditedAmount, ClientID) VALUES (?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 135, 1, -1, MM_IIF(Request.Form("tbxPaymentDate"), Request.Form("tbxPaymentDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxPaymentMethodID"), Request.Form("cbxPaymentMethodID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("tbxAmount"), Request.Form("tbxAmount"), 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxAmount2"), Request.Form("tbxAmount2"), 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("htbxClientID"), Request.Form("htbxClientID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_insert")) = "frmAddInvoice") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    'Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Invoices (InvoiceDate, PaymentMethodID, ClientID) VALUES (?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 135, 1, -1, MM_IIF(Request.Form("tbxInvoiceDate"), Request.Form("tbxInvoiceDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxPaymentMethodID"), Request.Form("cbxPaymentMethodID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("htbxClientID"), Request.Form("htbxClientID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_insert")) = "frmAddContact") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    'Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Contacts (ContactName, Title, Phone, Email, Skype, TeamViewer, Notes, ClientID) VALUES (?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxContactName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxTitle")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 20, Request.Form("tbxPhone")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("tbxEmail")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 50, Request.Form("tbxSkype")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("tbxTeamViewer"), Request.Form("tbxTeamViewer"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 202, 1, 1000, Request.Form("tbxNotes")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("htbxClientID"), Request.Form("htbxClientID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_insert")) = "frmAddProject") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    'Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Projects (ProjectDescription, StartDate, ProjectRate, ProjectPriority, ClientID) VALUES (?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxProjectDescription")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 135, 1, -1, MM_IIF(Request.Form("tbxStartDate"), Request.Form("tbxStartDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("tbxProjectRate"), Request.Form("tbxProjectRate"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxProjectPriority"), Request.Form("tbxProjectPriority"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("htbxClientID"), Request.Form("htbxClientID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_insert")) = "frmAddVitals") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    'Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.ClientVitals (VitalName, Username, Password, Address, Notes, ClientID) VALUES (?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxVitalName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxUsername")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 100, Request.Form("tbxPassword")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 100, Request.Form("tbxAddress")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 1000, Request.Form("tbxNotes")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("htbxClientID"), Request.Form("htbxClientID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim rstClients__lngClientID
rstClients__lngClientID = "1"
If (lngClientID <> "") Then 
  rstClients__lngClientID = lngClientID
End If
%>
<%
Dim rstClients
Dim rstClients_cmd
Dim rstClients_numRows

Set rstClients_cmd = Server.CreateObject ("ADODB.Command")
rstClients_cmd.ActiveConnection = MM_OBA_STRING
rstClients_cmd.CommandText = "SELECT * FROM Clients WHERE ClientID = ?" 
rstClients_cmd.Prepared = true
rstClients_cmd.Parameters.Append rstClients_cmd.CreateParameter("param1", 5, 1, -1, rstClients__lngClientID) ' adDouble

Set rstClients = rstClients_cmd.Execute
rstClients_numRows = 0
%>
<%
Dim rstContacts__lngClientID
rstContacts__lngClientID = "1"
If (lngClientID <> "") Then 
  rstContacts__lngClientID = lngClientID
End If
%>
<%
Dim rstContacts
Dim rstContacts_cmd
Dim rstContacts_numRows

Set rstContacts_cmd = Server.CreateObject ("ADODB.Command")
rstContacts_cmd.ActiveConnection = MM_OBA_STRING
rstContacts_cmd.CommandText = "SELECT * FROM Contacts WHERE ClientID = ?" 
rstContacts_cmd.Prepared = true
rstContacts_cmd.Parameters.Append rstContacts_cmd.CreateParameter("param1", 5, 1, -1, rstContacts__lngClientID) ' adDouble

Set rstContacts = rstContacts_cmd.Execute
rstContacts_numRows = 0
%>
<%
Dim rstClientVitals__lngClientID
rstClientVitals__lngClientID = "1"
If (lngClientID <> "") Then 
  rstClientVitals__lngClientID = lngClientID
End If
%>
<%
Dim rstClientVitals
Dim rstClientVitals_cmd
Dim rstClientVitals_numRows

Set rstClientVitals_cmd = Server.CreateObject ("ADODB.Command")
rstClientVitals_cmd.ActiveConnection = MM_OBA_STRING
rstClientVitals_cmd.CommandText = "SELECT * FROM ClientVitals WHERE ClientID = ? ORDER BY VitalName" 
rstClientVitals_cmd.Prepared = true
rstClientVitals_cmd.Parameters.Append rstClientVitals_cmd.CreateParameter("param1", 5, 1, -1, rstClientVitals__lngClientID) ' adDouble

Set rstClientVitals = rstClientVitals_cmd.Execute
rstClientVitals_numRows = 0
%>
<%
Dim rstProjects__lngClientID
rstProjects__lngClientID = "1"
If (lngClientID <> "") Then 
  rstProjects__lngClientID = lngClientID
End If
%>
<%
Dim rstProjects
Dim rstProjects_cmd
Dim rstProjects_numRows

Set rstProjects_cmd = Server.CreateObject ("ADODB.Command")
rstProjects_cmd.ActiveConnection = MM_OBA_STRING
rstProjects_cmd.CommandText = "SELECT Projects.ProjectID, Projects.ClientID, Projects.ProjectDescription, Projects.StartDate, Projects.ProjectRate, Projects.ProjectPriority FROM Projects  WHERE Projects.ClientID = ? ORDER BY Projects.ProjectPriority" 
rstProjects_cmd.Prepared = true
rstProjects_cmd.Parameters.Append rstProjects_cmd.CreateParameter("param1", 5, 1, -1, rstProjects__lngClientID) ' adDouble

Set rstProjects = rstProjects_cmd.Execute
rstProjects_numRows = 0
%>
<%
Dim rstInvoices__lngClientID
rstInvoices__lngClientID = "1"
If (lngClientID <> "") Then 
  rstInvoices__lngClientID = lngClientID
End If
%>
<%
Dim rstInvoices
Dim rstInvoices_cmd
Dim rstInvoices_numRows

Set rstInvoices_cmd = Server.CreateObject ("ADODB.Command")
rstInvoices_cmd.ActiveConnection = MM_OBA_STRING
rstInvoices_cmd.CommandText = "SELECT Invoices.InvoiceID, Invoices.ClientID, Invoices.InvoiceDate, Invoices.Sent, Invoices.RolloverHours, COALESCE(SUM(InvoiceDetails.Amount), 0) AS Amount, PaymentMethods.MethodName FROM Invoices LEFT OUTER JOIN  InvoiceDetails ON Invoices.InvoiceID = InvoiceDetails.InvoiceID INNER JOIN  PaymentMethods ON Invoices.PaymentMethodID = PaymentMethods.PaymentMethodID WHERE (Invoices.ClientID = ?) GROUP BY Invoices.InvoiceID, Invoices.ClientID, Invoices.InvoiceDate, Invoices.RolloverHours, Invoices.Sent, PaymentMethods.MethodName ORDER BY Invoices.InvoiceDate DESC" 
rstInvoices_cmd.Prepared = true
rstInvoices_cmd.Parameters.Append rstInvoices_cmd.CreateParameter("param1", 5, 1, -1, rstInvoices__lngClientID) ' adDouble

Set rstInvoices = rstInvoices_cmd.Execute
rstInvoices_numRows = 0
%>
<%
Dim rstPaymentMethods
Dim rstPaymentMethods_cmd
Dim rstPaymentMethods_numRows

Set rstPaymentMethods_cmd = Server.CreateObject ("ADODB.Command")
rstPaymentMethods_cmd.ActiveConnection = MM_OBA_STRING
rstPaymentMethods_cmd.CommandText = "SELECT * FROM PaymentMethods" 
rstPaymentMethods_cmd.Prepared = true

Set rstPaymentMethods = rstPaymentMethods_cmd.Execute
rstPaymentMethods_numRows = 0
%>
<%
Dim rstClientPayments__lngClientID
rstClientPayments__lngClientID = "1"
If (lngClientID <> "") Then 
  rstClientPayments__lngClientID = lngClientID
End If
%>
<%
Dim rstClientPayments
Dim rstClientPayments_cmd
Dim rstClientPayments_numRows

Set rstClientPayments_cmd = Server.CreateObject ("ADODB.Command")
rstClientPayments_cmd.ActiveConnection = MM_OBA_STRING
rstClientPayments_cmd.CommandText = "SELECT ClientPayments.ClientPaymentID, ClientPayments.ClientID, ClientPayments.PaymentMethodID, ClientPayments.PaymentDate, ClientPayments.Amount, ClientPayments.CreditedAmount, PaymentMethods.MethodName, PaymentMethods.Discount FROM ClientPayments INNER JOIN  PaymentMethods ON ClientPayments.PaymentMethodID = PaymentMethods.PaymentMethodID WHERE (ClientPayments.ClientID = ?) ORDER BY ClientPayments.PaymentDate DESC" 
rstClientPayments_cmd.Prepared = true
rstClientPayments_cmd.Parameters.Append rstClientPayments_cmd.CreateParameter("param1", 5, 1, -1, rstClientPayments__lngClientID) ' adDouble

Set rstClientPayments = rstClientPayments_cmd.Execute
rstClientPayments_numRows = 0
%>
<%
Dim rstProjectDetails__lngClientID
rstProjectDetails__lngClientID = "2"
If (lngClientID <> "") Then 
  rstProjectDetails__lngClientID = lngClientID
End If
%>
<%
Dim rstProjectDetails
Dim rstProjectDetails_cmd
Dim rstProjectDetails_numRows

Set rstProjectDetails_cmd = Server.CreateObject ("ADODB.Command")
rstProjectDetails_cmd.ActiveConnection = MM_OBA_STRING
rstProjectDetails_cmd.CommandText = "SELECT Projects.ProjectID, Projects.ProjectDescription, ProjectDetails.ProjectDetailID, ProjectDetails.DetailDescription,  ProjectStages.StageName, Projects.ProjectRate,  Vendors.VendorName  FROM Vendors  INNER JOIN ProjectStages  INNER JOIN ProjectDetails ON ProjectStages.ProjectStageID = ProjectDetails.ProjectStageID ON Vendors.VendorID = ProjectDetails.VendorID  INNER JOIN Projects ON  ProjectDetails.ProjectID = Projects.ProjectID  WHERE ClientID = ? AND ProjectStages.ProjectStageID BETWEEN 2 AND 5  ORDER BY Projects.ProjectPriority, ProjectDetails.Priority, ProjectStages.SortOrder DESC" 
rstProjectDetails_cmd.Prepared = true
rstProjectDetails_cmd.Parameters.Append rstProjectDetails_cmd.CreateParameter("param1", 5, 1, -1, rstProjectDetails__lngClientID) ' adDouble

Set rstProjectDetails = rstProjectDetails_cmd.Execute
rstProjectDetails_numRows = 0
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
<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/themes/base/jquery-ui.css" />
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/jquery-ui.min.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
$(function() {
	$("#tbxPaymentDate").datepicker();
});
$(function() {
	$("#tbxInvoiceDate").datepicker();
});
$(function() {
	$("#tbxStartDate").datepicker();
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
If bolClientsViewGranted Then
	If bolClientsEditGranted Then
		strEditLink = "<a href=""ClientEdit.asp?lngClientID=" & (rstClients.Fields.Item("ClientID").Value) & """>"
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
                    <td align="right"><strong>Name</strong></td>
                    <td><%=strEditLink & (rstClients.Fields.Item("ClientName").Value) & strEndEditLink%></td>
                    <td align="right"><strong>Source</strong></td>
                    <td><%=strEditLink & (rstClients.Fields.Item("Source").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Phone</strong></td>
                    <td><%=strEditLink & (rstClients.Fields.Item("Phone").Value) & strEndEditLink%></td>
                    <td align="right"><strong>Current Rate</strong></td>
                    <td><%=strEditLink & FormatCurrency(rstClients.Fields.Item("CurrentRate").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Email</strong></td>
                    <td><%=strEditLink & (rstClients.Fields.Item("Email").Value) & strEndEditLink%></td>
                    <td align="right"><strong>Skype</strong></td>
                    <td><%=strEditLink & (rstClients.Fields.Item("Skype").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right">&nbsp;</td>
                    <td>&nbsp;</td>
                    <td align="right"><strong>TeamViewer</strong></td>
                    <td><%=strEditLink & (rstClients.Fields.Item("Teamviewer").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Notes</strong></td>
                    <td colspan="3"><%=strEditLink & (rstClients.Fields.Item("Notes").Value) & strEndEditLink%></td>
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
                    <th colspan="10"><h2>Contacts</h2></th>
                  </tr>
                  <tr>
                    <th>&nbsp;</th>
                    <th align="left"><h4>Name</h4></th>
                    <th align="left"><h4>Title</h4></th>
                    <th align="left"><h4>Phone</h4></th>
                    <th align="left"><h4>Email</h4></th>
                    <th align="left"><h4>Skype</h4></th>
                    <th align="left"><h4>TeamViewer</h4></th>
                    <th align="left"><h4>Notes</h4></th>
                    <th>&nbsp;</th>
                    <th>&nbsp;</th>
                  </tr>
                  <form id="frmAddContact" name="frmAddContact" method="POST" action="<%=MM_editAction%>">
                  <tr>
                    <td>&nbsp;</td>
                    <td>
                      <input type="text" name="tbxContactName" id="tbxContactName" /></td>
                    <td><label for="tbxTitle"></label>
                    <input type="text" name="tbxTitle" id="tbxTitle" /></td>
                    <td><input name="tbxPhone" type="text" id="tbxPhone" size="15" /></td>
                    <td><input name="tbxEmail" type="text" id="tbxEmail" size="35" /></td>
                    <td><input type="text" name="tbxSkype" id="tbxSkype" /></td>
                    <td><input name="tbxTeamViewer" type="text" id="tbxTeamViewer" size="10" /></td>
                    <td><input type="text" name="tbxNotes" id="tbxNotes" /></td>
                    <td><input type="submit" name="btnAdd" id="btnAdd" value="Add" />
                    <input name="htbxClientID" type="hidden" id="htbxClientID" value="<%=lngClientID%>" /></td>
                    <td>&nbsp;</td>
                  </tr>
                  <input type="hidden" name="MM_insert" value="frmAddContact" />
                  </form>
                  <tr>
                    <td colspan="10"><hr /></td>
                  </tr>
<%
	Do While Not rstContacts.EOF
		If bolClientsEditGranted Then
			strEditLink = "<a href=""ContactEdit.asp?lngContactID=" & (rstContacts.Fields.Item("ContactID").Value) & """>"
			strEndEditLink = "</a>&nbsp;"
		Else
			strEditLink = ""
			strEndEditLink = "&nbsp;"
		End If
%>                  
                  <tr>
                    <td>&nbsp;</td>
                    <td><%=strEditLink & (rstContacts.Fields.Item("ContactName").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstContacts.Fields.Item("Title").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstContacts.Fields.Item("Phone").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstContacts.Fields.Item("Email").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstContacts.Fields.Item("Skype").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstContacts.Fields.Item("TeamViewer").Value) & strEndEditLink%></td>
                    <td colspan="2"><%=strEditLink & (rstContacts.Fields.Item("Notes").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                
<%
		rstContacts.MoveNext
	Loop
%>
                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="9">&nbsp;</td>
                  </tr>
                </table>
                <table border="0" cellspacing="0" cellpadding="0" class="box">
                  <tr>
                    <th colspan="8"><h2>Vital Information</h2></th>
                  </tr>
                  <tr>
                    <th>&nbsp;</th>
                    <th align="left"><h4>Name</h4></th>
                    <th align="left"><h4>User</h4></th>
                    <th align="left"><h4>Password</h4></th>
                    <th align="left"><h4>Address</h4></th>
                    <th align="left"><h4>Notes</h4></th>
                    <th>&nbsp;</th>
                    <th>&nbsp;</th>
                  </tr>
                  <form id="frmAddVitals" name="frmAddVitals" method="POST" action="<%=MM_editAction%>">
                  <tr>
                    <td>&nbsp;</td>
                    <td>
                      <input type="text" name="tbxVitalName" id="tbxVitalName" /></td>
                    <td><input type="text" name="tbxUsername" id="tbxUsername" /></td>
                    <td><input name="tbxPassword" type="text" id="tbxPassword" /></td>
                    <td><input name="tbxAddress" type="text" id="tbxAddress" size="35" /></td>
                    <td><textarea name="tbxNotes" cols="50" rows="2" id="tbxNotes"></textarea></td>
                    <td><input type="submit" name="btnAdd" id="btnAdd" value="Add" />
                    <input name="htbxClientID" type="hidden" id="htbxClientID" value="<%=lngClientID%>" /></td>
                    <td>&nbsp;</td>
                  </tr>
                  <input type="hidden" name="MM_insert" value="frmAddVitals" />
                  </form>
                  <tr>
                    <td colspan="8"><hr /></td>
                  </tr>
<%
	Do While Not rstClientVitals.EOF
		If bolClientsEditGranted Then
			strEditLink = "<a href=""ClientVitalEdit.asp?lngClientVitalID=" & (rstClientVitals.Fields.Item("ClientVitalID").Value) & """>"
			strEndEditLink = "</a>&nbsp;"
		Else
			strEditLink = ""
			strEndEditLink = "&nbsp;"
		End If
%>                  
                  <tr class="tr_hover">
                    <td>&nbsp;</td>
                    <td><%=strEditLink & (rstClientVitals.Fields.Item("VitalName").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstClientVitals.Fields.Item("Username").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstClientVitals.Fields.Item("Password").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstClientVitals.Fields.Item("Address").Value) & strEndEditLink%></td>
                    <td colspan="2"><%=strEditLink & (rstClientVitals.Fields.Item("Notes").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                
<%
		rstClientVitals.MoveNext
	Loop
%>
                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="7">&nbsp;</td>
                  </tr>

<%		
	If bolProjectsViewGranted Then
		If bolProjectsAddGranted Then
%>
                </table>
    <table border="0" cellspacing="0" cellpadding="0" class="box">
                  <tr>
                    <th colspan="9"><h2>Projects</h2></th>
                  </tr>
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="left"><h4>Description</h4></th>
	    <th align="center"><h4>Start</h4></th>
	    <th align="right"><h4>Rate</h4></th>
	    <th align="center"><h4>Priority</h4></th>
	    <th align="center"><h4>Project ID</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
      <form id="frmAddProject" name="frmAddProject" method="POST" action="<%=MM_editAction%>">
	  <tr>
	    <td>&nbsp;</td>
	    <td><input name="tbxProjectDescription" type="text" id="tbxProjectDescription" tabindex="0" size="50" maxlength="50" /></td>
	    <td align="center"><input name="tbxStartDate" type="text" id="tbxStartDate" value="<%=Date%>" size="11" style="text-align: center" /></td>
	    <td align="right"><input name="tbxProjectRate" type="text" id="tbxProjectRate" size="8" style="text-align: right" /></td>
	    <td align="center"><input name="tbxProjectPriority" type="text" id="tbxProjectPriority" size="5" style="text-align: center" /></td>
	    <td align="center"><input type="submit" name="btnAdd" id="btnAdd" value="Add Project" />
	      <input name="htbxClientID" type="hidden" id="htbxClientID" value="<%=lngClientID%>" /></td>
	    <td>&nbsp;</td>
      </tr>
      <input type="hidden" name="MM_insert" value="frmAddProject" />
      </form>
	  <tr>
	    <td colspan="7"><hr /></td>
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
%>
                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="7">&nbsp;</td>
                  </tr>
                </table>
    <table border="0" cellspacing="0" cellpadding="0" class="box">
                  <tr>
                    <th colspan="9"><h2>Active Milestones</h2></th>
                  </tr>
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="left"><h4>Project</h4></th>
	    <th align="left"><h4>Milestone</h4></th>
	    <th align="left"><h4>Stage</h4></th>
	    <th align="left"><h4>Rate</h4></th>
	    <th align="left"><h4>Vendor</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
<%
		Do While Not rstProjectDetails.EOF
%>      
	  <tr>
        <td><a href="ProjectInformation.asp?lngProjectID=<%=(rstProjectDetails.Fields.Item("ProjectID").Value)%>" class="row_info"></a></td>
        <td nowrap="nowrap"><%=(rstProjectDetails.Fields.Item("ProjectDescription").Value)%></td>
		<td nowrap="nowrap"><%=(rstProjectDetails.Fields.Item("DetailDescription").Value)%></td>
		<td nowrap="nowrap"><%=(rstProjectDetails.Fields.Item("StageName").Value)%></td>
		<td><%=(rstProjectDetails.Fields.Item("ProjectRate").Value)%></td>
		<td><%=(rstProjectDetails.Fields.Item("VendorName").Value)%></td>
		<td>&nbsp;</td>
	  </tr>
<%
			rstProjectDetails.MoveNext
		Loop
%>
                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="7">&nbsp;</td>
                  </tr>

<%		
	End If
	If bolInvoicesViewGranted Then
		If bolInvoicesAddGranted Then
%>
                </table>
                <table border="0" cellspacing="0" cellpadding="0" class="box">
                  <tr>
                    <th colspan="8"><h2>Invoices</h2></th>
                  </tr>
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="center"><h4>Invoice Date</h4></th>
	    <th align="center"><h4>Invoicing Method</h4></th>
	    <th align="center"><h4>Invoice ID</h4></th>
	    <th align="right"><h4>Amount</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
      <form id="frmAddInvoice" name="frmAddInvoice" method="POST" action="<%=MM_editAction%>">
	  <tr>
	    <td>&nbsp;</td>
	    <td align="center"><input name="tbxInvoiceDate" type="text" id="tbxInvoiceDate" value="<%=Date%>" size="11" style="text-align: center" /></td>
	    <td align="center"><select name="cbxPaymentMethodID" id="cbxPaymentMethodID">
	      <%
While (NOT rstPaymentMethods.EOF)
%>
	      <option value="<%=(rstPaymentMethods.Fields.Item("PaymentMethodID").Value)%>"><%=(rstPaymentMethods.Fields.Item("MethodName").Value)%></option>
	      <%
  rstPaymentMethods.MoveNext()
Wend
If (rstPaymentMethods.CursorType > 0) Then
  rstPaymentMethods.MoveFirst
Else
  rstPaymentMethods.Requery
End If
%>
	      </select></td>
	    <td colspan="2"><input type="submit" name="btnAdd" id="btnAddInvoice" value="Add Invoice" />
	      <input name="htbxClientID" type="hidden" id="htbxClientID" value="<%=lngClientID%>" /></td>
	    <td>&nbsp;</td>
      </tr>
      <input type="hidden" name="MM_insert" value="frmAddInvoice" />
      </form>
	  <tr>
	    <td colspan="6"><hr /></td>
      </tr>
<%
		End If
		curTotalInvoices = 0
		Do While Not rstInvoices.EOF
			If bolInvoicesEditGranted Then
				strEdit = "<a href=""InvoiceEdit.asp?lngInvoiceID=" & (rstInvoices.Fields.Item("InvoiceID").Value) & """>"
				strEditEnd = "</a>"
			Else
				strEdit = ""
				strEditEnd = ""
			End If 
			curTotalInvoices = curTotalInvoices + Round(rstInvoices.Fields.Item("Amount").Value, 2)
%>      
	  <tr class="tr_hover">
        <td><a href="InvoiceInformation.asp?lngInvoiceID=<%=(rstInvoices.Fields.Item("InvoiceID").Value)%>" class="row_info"></a></td>
        <td align="center"><%=strEdit & (rstInvoices.Fields.Item("InvoiceDate").Value) & strEditEnd%></td>
		<td align="center"><%=strEdit & (rstInvoices.Fields.Item("MethodName").Value) & strEditEnd%></td>
		<td align="center"><%=strEdit & (rstInvoices.Fields.Item("InvoiceID").Value) & strEditEnd%></td>
		<td align="right"><%=strEdit & FormatCurrency(rstInvoices.Fields.Item("Amount").Value) & strEditEnd%></td>
		<td>&nbsp;</td>
	  </tr>
<%
			rstInvoices.MoveNext
		Loop
%>
	  <tr>
	    <td colspan="6"><hr /></td>
      </tr>
	  <tr>
	    <td>&nbsp;</td>
	    <td align="center">&nbsp;</td>
	    <td align="center">&nbsp;</td>
	    <td align="right"><strong>Total Invoices:</strong></td>
	    <td align="right"><%=FormatCurrency(curTotalInvoices)%></td>
	    <td>&nbsp;</td>
	    </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="6">&nbsp;</td>
                  </tr>
<%
	End If
	If bolInvoicesViewGranted Then
		If bolInvoicesAddGranted Then
%>
                </table>
    <table border="0" cellspacing="0" cellpadding="0" class="box">
                  <tr>
                    <th colspan="9"><h2>Payments</h2></th>
      </tr>
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="center"><h4>Payment Date</h4></th>
	    <th align="center"><h4>Payment Method</h4></th>
	    <th><h4>Amount Credited</h4></th>
	    <th align="right"><h4>Amount Received</h4></th>
	    <th align="right"><h4>Services Fees</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
      <form id="frmAddPayment" name="frmAddPayment" method="POST" action="<%=MM_editAction%>">
	  <tr>
	    <td>&nbsp;</td>
	    <td align="center"><input name="tbxPaymentDate" type="text" id="tbxPaymentDate" value="<%=Date%>" size="11" style="text-align: center" /></td>
	    <td align="center"><select name="cbxPaymentMethodID" id="cbxPaymentMethodID">
	      <%
While (NOT rstPaymentMethods.EOF)
%>
	      <option value="<%=(rstPaymentMethods.Fields.Item("PaymentMethodID").Value)%>"><%=(rstPaymentMethods.Fields.Item("MethodName").Value)%></option>
	      <%
  rstPaymentMethods.MoveNext()
Wend
If (rstPaymentMethods.CursorType > 0) Then
  rstPaymentMethods.MoveFirst
Else
  rstPaymentMethods.Requery
End If
%>
	      </select></td>
	    <td align="right"><input name="tbxAmount" type="text" id="tbxAmount" size="8" /></td>
	    <td align="right"><input name="tbxAmount2" type="text" id="tbxAmount2" size="8" /></td>
	    <td><input type="submit" name="btnAddpayment" id="btnAddpayment" value="Add Payment" />
          <input name="htbxClientID" type="hidden" id="htbxClientID" value="<%=lngClientID%>" /></td>
	    <td>&nbsp;</td>
      </tr>
      <input type="hidden" name="MM_insert" value="frmAddPayment" />
      </form>
	  <tr>
	    <td colspan="7"><hr /></td>
      </tr>
<%
		End If
		curTotalPayments = 0
		curTotalCredits = 0 
		Do While Not rstClientPayments.EOF
			If bolPaymentsEditGranted Then
				strEdit = "<a href=""ClientPaymentEdit.asp?lngClientPaymentID=" & (rstClientPayments.Fields.Item("ClientPaymentID").Value) & """>"
				strEditEnd = "</a>"
			Else
				strEdit = ""
				strEditEnd = ""
			End If
			curTotalPayments = curTotalPayments + CDbl(rstClientPayments.Fields.Item("Amount").Value)
			curTotalCredits = curTotalCredits +  CDbl(rstClientPayments.Fields.Item("CreditedAmount").Value)
%>      
	  <tr class="tr_hover">
		<td>&nbsp;</td>
        <td align="center"><%=strEdit & (rstClientPayments.Fields.Item("PaymentDate").Value) & strEditEnd%></td>
		<td align="center"><%=strEdit & (rstClientPayments.Fields.Item("MethodName").Value) & strEditEnd%></td>
		<td align="right"><%=strEdit & FormatCurrency(rstClientPayments.Fields.Item("CreditedAmount").Value) & strEditEnd%></td>
		<td align="right"><%=strEdit & FormatCurrency(rstClientPayments.Fields.Item("Amount").Value) & strEditEnd%></td>
		<td align="right"><%=strEdit & FormatCurrency(rstClientPayments.Fields.Item("Amount").Value - rstClientPayments.Fields.Item("CreditedAmount").Value) & strEditEnd%></td>
		<td>&nbsp;</td>
	  </tr>
<%
			rstClientPayments.MoveNext
		Loop
%>
	  <tr>
	    <td colspan="7"><hr /></td>
      </tr>
	  <tr>
	    <td>&nbsp;</td>
	    <td align="center">&nbsp;</td>
	    <td align="right"><strong>Totals:</strong></td>
	    <td align="right"><%=FormatCurrency(curTotalCredits)%></td>
	    <td align="right"><%=FormatCurrency(curTotalPayments)%></td>
	    <td align="right"><%=FormatCurrency(curTotalPayments - curTotalCredits)%></td>
	    <td>&nbsp;</td>
	    </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="7">&nbsp;</td>
                  </tr>
                </table>
                <table border="0" cellspacing="0" cellpadding="0" class="box">
                  <tr>
                    <th colspan="8"><h2> Account Summary</h2></th>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="center"><strong>Account Balance: </strong><%=FormatCurrency(curTotalInvoices - curTotalCredits)%> </td>
                    <td align="center">&nbsp;</td>
                    <td align="center">&nbsp;</td>
                    <td align="right">&nbsp;</td>
                    <td align="left">&nbsp;</td>
                  </tr>
<%
	End If

Else
%>               

                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="12">Certain &quot;Clients&quot; permissions are required to view this information.</td>
                  </tr>
<%
End If
%>
                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="8">&nbsp;</td>
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
rstContacts.Close()
Set rstContacts = Nothing
%>
<%
rstClientVitals.Close()
Set rstClientVitals = Nothing
%>
<%
rstProjects.Close()
Set rstProjects = Nothing
%>
<%
rstInvoices.Close()
Set rstInvoices = Nothing
%>
<%
rstPaymentMethods.Close()
Set rstPaymentMethods = Nothing
%>
<%
rstClientPayments.Close()
Set rstClientPayments = Nothing
%>
<%
rstProjectDetails.Close()
Set rstProjectDetails = Nothing
%>
