<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngVendorID
lngVendorID = Request.QueryString("lngVendorID")
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
    MM_editCmd.CommandText = "INSERT INTO dbo.VendorPayments (PaymentDate, PaymentMethodID, AmountPaid, AmountCredited, VendorID) VALUES (?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 135, 1, -1, MM_IIF(Request.Form("tbxPaymentDate"), Request.Form("tbxPaymentDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxPaymentMethodID"), Request.Form("cbxPaymentMethodID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("tbxAmountPaid"), Request.Form("tbxAmountPaid"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxAmountCredited"), Request.Form("tbxAmountCredited"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("htbxVendorID"), Request.Form("htbxVendorID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_insert")) = "frmWorkHistory") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    'Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.WorkHistorys (WorkDate, ProjectDetailID, WorkDescription, Hours, VendorID) VALUES (?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 135, 1, -1, MM_IIF(Request.Form("tbxWorkDate"), Request.Form("tbxWorkDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxProjectDetailID"), Request.Form("cbxProjectDetailID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 1000, Request.Form("tbxWorkDescription")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxHours"), Request.Form("tbxHours"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("htbxVendorID"), Request.Form("htbxVendorID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_insert")) = "frmInvoice") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    'Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.VendorInvoices (InvoiceDate, PaymentMethodID, VendorRate,Notes,  VendorID) VALUES (?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 135, 1, -1, MM_IIF(Request.Form("tbxInvoiceDate"), Request.Form("tbxInvoiceDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxPaymentMethodID"), Request.Form("cbxPaymentMethodID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("tbxVendorRate"), Request.Form("tbxVendorRate"), 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 500, Request.Form("tbxNotes")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("htbxVendorID"), Request.Form("htbxVendorID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim rstVendors__lngVendorID
rstVendors__lngVendorID = "1"
If (lngVendorID <> "") Then 
  rstVendors__lngVendorID = lngVendorID
End If
%>
<%
Dim rstVendors
Dim rstVendors_cmd
Dim rstVendors_numRows

Set rstVendors_cmd = Server.CreateObject ("ADODB.Command")
rstVendors_cmd.ActiveConnection = MM_OBA_STRING
rstVendors_cmd.CommandText = "SELECT * FROM Vendors WHERE VendorID = ?" 
rstVendors_cmd.Prepared = true
rstVendors_cmd.Parameters.Append rstVendors_cmd.CreateParameter("param1", 5, 1, -1, rstVendors__lngVendorID) ' adDouble

Set rstVendors = rstVendors_cmd.Execute
rstVendors_numRows = 0
%>
<%
Dim rstVendorPayments__lngVendorID
rstVendorPayments__lngVendorID = "1"
If (lngVendorID <> "") Then 
  rstVendorPayments__lngVendorID = lngVendorID
End If
%>
<%
Dim rstVendorPayments
Dim rstVendorPayments_cmd
Dim rstVendorPayments_numRows

Set rstVendorPayments_cmd = Server.CreateObject ("ADODB.Command")
rstVendorPayments_cmd.ActiveConnection = MM_OBA_STRING
rstVendorPayments_cmd.CommandText = "SELECT VendorPayments.VendorPaymentID, VendorPayments.VendorID, VendorPayments.PaymentDate, VendorPayments.AmountPaid,VendorPayments.AmountCredited, PaymentMethods.MethodName  FROM VendorPayments INNER JOIN PaymentMethods ON VendorPayments.PaymentMethodID = PaymentMethods.PaymentMethodID  WHERE (VendorPayments.VendorID = ?) ORDER BY PaymentDate DESC" 
rstVendorPayments_cmd.Prepared = true
rstVendorPayments_cmd.Parameters.Append rstVendorPayments_cmd.CreateParameter("param1", 5, 1, -1, rstVendorPayments__lngVendorID) ' adDouble

Set rstVendorPayments = rstVendorPayments_cmd.Execute
rstVendorPayments_numRows = 0
%>
<%
Dim rstPaymentMethods
Dim rstPaymentMethods_cmd
Dim rstPaymentMethods_numRows

Set rstPaymentMethods_cmd = Server.CreateObject ("ADODB.Command")
rstPaymentMethods_cmd.ActiveConnection = MM_OBA_STRING
rstPaymentMethods_cmd.CommandText = "SELECT PaymentMethodID, MethodName FROM PaymentMethods" 
rstPaymentMethods_cmd.Prepared = true

Set rstPaymentMethods = rstPaymentMethods_cmd.Execute
rstPaymentMethods_numRows = 0
%>
<%
Dim rstVendorInvoices__lngVendoriD
rstVendorInvoices__lngVendoriD = "1"
If (lngVendoriD <> "") Then 
  rstVendorInvoices__lngVendoriD = lngVendoriD
End If
%>
<%
Dim rstVendorInvoices
Dim rstVendorInvoices_cmd
Dim rstVendorInvoices_numRows

Set rstVendorInvoices_cmd = Server.CreateObject ("ADODB.Command")
rstVendorInvoices_cmd.ActiveConnection = MM_OBA_STRING
rstVendorInvoices_cmd.CommandText = "SELECT VendorInvoices.VendorInvoiceID, VendorInvoices.VendorID, VendorInvoices.PaymentMethodID, VendorInvoices.InvoiceDate, VendorInvoices.Notes, VendorInvoices.VendorRate, PaymentMethods.MethodName, SUM(VendorInvoiceDetails.Amount)  AS InvoiceAmount FROM VendorInvoices INNER JOIN PaymentMethods ON VendorInvoices.PaymentMethodID = PaymentMethods.PaymentMethodID LEFT OUTER JOIN VendorInvoiceDetails ON VendorInvoices.VendorInvoiceID = VendorInvoiceDetails.VendorInvoiceID WHERE (VendorInvoices.VendorID = ?) GROUP BY VendorInvoices.VendorInvoiceID, VendorInvoices.Notes, VendorInvoices.VendorID, VendorInvoices.PaymentMethodID, VendorInvoices.InvoiceDate, VendorInvoices.VendorRate, PaymentMethods.MethodName ORDER BY VendorInvoices.InvoiceDate DESC" 
rstVendorInvoices_cmd.Prepared = true
rstVendorInvoices_cmd.Parameters.Append rstVendorInvoices_cmd.CreateParameter("param1", 5, 1, -1, rstVendorInvoices__lngVendoriD) ' adDouble

Set rstVendorInvoices = rstVendorInvoices_cmd.Execute
rstVendorInvoices_numRows = 0
%>
<%
Dim rstWorkHistorys__lngVendorID
rstWorkHistorys__lngVendorID = "1"
If (lngVendorID <> "") Then 
  rstWorkHistorys__lngVendorID = lngVendorID
End If
%>
<%
Dim rstWorkHistorys
Dim rstWorkHistorys_cmd
Dim rstWorkHistorys_numRows

Set rstWorkHistorys_cmd = Server.CreateObject ("ADODB.Command")
rstWorkHistorys_cmd.ActiveConnection = MM_OBA_STRING
rstWorkHistorys_cmd.CommandText = "SELECT WorkHistorys.WorkHistoryID, LEFT(WorkHistorys.WorkDescription, 150) + ' - ' + CAST(WorkHistorys.WorkDate AS nvarchar(11))  AS WorkDescription, ProjectDetails.DetailDescription, WorkHistorys.ProjectDetailID, WorkHistorys.VendorID, WorkHistorys.WorkDate, WorkHistorys.Hours FROM WorkHistorys INNER JOIN ProjectDetails ON WorkHistorys.ProjectDetailID = ProjectDetails.ProjectDetailID INNER JOIN Projects ON ProjectDetails.ProjectID = Projects.ProjectID LEFT OUTER JOIN VendorInvoiceDetails ON WorkHistorys.WorkHistoryID = VendorInvoiceDetails.WorkHistoryID WHERE (WorkHistorys.VendorID =  ?) AND (VendorInvoiceDetails.VendorInvoiceDetailID IS NULL) ORDER BY WorkHistorys.WorkDate DESC" 
rstWorkHistorys_cmd.Prepared = true
rstWorkHistorys_cmd.Parameters.Append rstWorkHistorys_cmd.CreateParameter("param1", 5, 1, -1, rstWorkHistorys__lngVendorID) ' adDouble

Set rstWorkHistorys = rstWorkHistorys_cmd.Execute
rstWorkHistorys_numRows = 0
%>
<%
Dim rstProjectDetails__lngVendorID
rstProjectDetails__lngVendorID = "1"
If (lngVendorID <> "") Then 
  rstProjectDetails__lngVendorID = lngVendorID
End If
%>
<%
Dim rstProjectDetails
Dim rstProjectDetails_cmd
Dim rstProjectDetails_numRows

Set rstProjectDetails_cmd = Server.CreateObject ("ADODB.Command")
rstProjectDetails_cmd.ActiveConnection = MM_OBA_STRING
rstProjectDetails_cmd.CommandText = "SELECT ProjectDetails.ProjectDetailID, Clients.ClientName + ' - ' + ProjectDetails.DetailDescription AS ProjectDescr FROM Projects INNER JOIN ProjectDetails ON Projects.ProjectID = ProjectDetails.ProjectID INNER JOIN Clients ON Projects.ClientID = Clients.ClientID WHERE (ProjectDetails.VendorID = ?)" 
rstProjectDetails_cmd.Prepared = true
rstProjectDetails_cmd.Parameters.Append rstProjectDetails_cmd.CreateParameter("param1", 5, 1, -1, rstProjectDetails__lngVendorID) ' adDouble

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
<link href="/global/jquery/css/ui-lightness/jquery-ui-1.10.3.custom.css" rel="stylesheet" />
<script type="text/javascript" src="/global/jquery/js/jquery-1.9.1.js"></script>
<script type="text/javascript" src="/global/jquery/js/jquery-ui-1.10.3.custom.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
$(function() {
	$("#tbxInvoiceDate").datepicker();
});
$(function() {
	$("#tbxPaymentDate").datepicker();
});
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
If bolVendorsViewGranted Then
	If bolVendorsEditGranted Then
		strEditLink = "<a href=""VendorEdit.asp?lngVendorID=" & (rstVendors.Fields.Item("VendorID").Value) & """>"
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
                    <td><%=strEditLink & (rstVendors.Fields.Item("VendorName").Value) & strEndEditLink%></td>
                    <td align="right"><strong>Skype</strong></td>
                    <td><%=strEditLink & (rstVendors.Fields.Item("SkypeID").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Country</strong></td>
                    <td><%=strEditLink & (rstVendors.Fields.Item("Country").Value) & strEndEditLink%></td>
                    <td align="right"><strong>Teamviewer</strong></td>
                    <td><%=strEditLink & (rstVendors.Fields.Item("TeamViewerID").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Current Rate:</strong></td>
                    <td><%=FormatCurrency(rstVendors.Fields.Item("Rate").Value)%></td>
                    <td align="right"><strong>Phone</strong></td>
                    <td><%=strEditLink & (rstVendors.Fields.Item("Phone").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right">&nbsp;</td>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Email</strong></td>
                    <td><%=strEditLink & (rstVendors.Fields.Item("Email").Value) & strEndEditLink%></td>
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
<%
	If bolInvoicesViewGranted Then
		If bolInvoicesAddGranted Then
%>
                </table>
    <table border="0" cellspacing="0" cellpadding="0" class="box">
                  <tr>
                    <th colspan="9"><h2>Work History</h2></th>
      </tr>
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="center"><h4>Work Date</h4></th>
	    <th align="center"><h4>Project Detail</h4></th>
	    <th align="left"><h4>Work Description</h4></th>
	    <th align="right"><h4>Hours</h4></th>
	    <th><h4>Amount</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
      <form id="frmWorkHistory" name="frmWorkHistory" method="POST" action="<%=MM_editAction%>">
	  <tr>
	    <td>&nbsp;</td>
	    <td align="center"><input name="tbxWorkDate" type="text" id="tbxWorkDate" value="<%=Date%>" size="11" style="text-align: center" /></td>
	    <td align="center"><select name="cbxProjectDetailID" id="cbxProjectDetailID">
	      <%
While (NOT rstProjectDetails.EOF)
%>
	      <option value="<%=(rstProjectDetails.Fields.Item("ProjectDetailID").Value)%>"><%=(rstProjectDetails.Fields.Item("ProjectDescr").Value)%></option>
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
	    <td><textarea name="tbxWorkDescription" id="tbxWorkDescription" cols="45" rows="5"></textarea></td>
	    <td align="right"><input name="tbxHours" type="text" id="tbxHours" size="8" /></td>
	    <td><input type="submit" name="btnAddWorkHistory" id="btnAddWorkHistory" value="Add History" />
	      <input name="htbxVendorID" type="hidden" id="htbxVendorID" value="<%=lngVendorID%>" /></td>
	    <td>&nbsp;</td>
      </tr>
      <input type="hidden" name="MM_insert" value="frmWorkHistory" />
      </form>
	  <tr>
	    <td colspan="7"><hr /></td>
      </tr>
<%
		End If
		curTotalWorkHistoryAmount = 0
		dblTotalWorkHistoryHours = 0
		Do While Not rstWorkHistorys.EOF
			If bolInvoicesEditGranted Then
				strEdit = "<a href=""WorkHistoryEdit.asp?lngWorkHistoryID=" & (rstWorkHistorys.Fields.Item("WorkHistoryID").Value) & """>"
				strEditEnd = "</a>"
			Else
				strEdit = ""
				strEditEnd = ""
			End If
			If IsNull(rstWorkHistorys.Fields.Item("Hours").Value) Then
				curWorkHistoryAmount = 0			
				curWorkHistoryHours = 0			
			Else
				curWorkHistoryAmount = CDbl(rstWorkHistorys.Fields.Item("Hours").Value) * CDbl(rstVendors.Fields.Item("Rate").Value)
				curWorkHistoryHours = CDbl(rstWorkHistorys.Fields.Item("Hours").Value)
			
			End If
			curTotalWorkHistoryAmount = curTotalWorkHistoryAmount + curWorkHistoryAmount
			dblTotalWorkHistoryHours = dblTotalWorkHistoryHours + curWorkHistoryHours
%>      
	  <tr>
		<td>&nbsp;</td>
        <td align="center"><%=strEdit & (rstWorkHistorys.Fields.Item("WorkDate").Value) & strEditEnd%></td>
		<td align="center"><%=strEdit & (rstWorkHistorys.Fields.Item("DetailDescription").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstWorkHistorys.Fields.Item("WorkDescription").Value) & strEditEnd%></td>
		<td align="right"><%=strEdit & curWorkHistoryHours & strEditEnd%></td>
		<td align="center"><%=FormatCurrency(curWorkHistoryAmount)%></td>
		<td>&nbsp;</td>
	  </tr>
<%
			rstWorkHistorys.MoveNext
		Loop
%>
	  <tr>
	    <td colspan="7"><hr /></td>
      </tr>
	  <tr>
	    <td>&nbsp;</td>
	    <td align="center">&nbsp;</td>
	    <td align="center">&nbsp;</td>
	    <td align="right"><strong>Waiting Invoicing:</strong></td>
	    <td><%=(dblTotalWorkHistoryHours)%></td>
	    <td><%=FormatCurrency(curTotalWorkHistoryAmount)%></td>
	    <td>&nbsp;</td>
	    </tr>
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
                    <th colspan="9"><h2>Invoices</h2></th>
      </tr>
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="center"><h4>Invoice Date</h4></th>
	    <th align="center"><h4>Delivery Method</h4></th>
	    <th align="left"><h4>Notes</h4></th>
	    <th align="right"><h4>Vendor Rate</h4></th>
	    <th><h4>Invoice Amount</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
      <form id="frmInvoice" name="frmInvoice" method="POST" action="<%=MM_editAction%>">
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
	    <td><textarea name="tbxNotes" id="tbxNotes" cols="45" rows="5"></textarea></td>
	    <td align="right"><input name="tbxVendorRate" type="text" id="tbxVendorRate" value="<%=(rstVendors.Fields.Item("Rate").Value)%>" size="8" /></td>
	    <td><input type="submit" name="btnAddInvoice" id="btnAddInvoice" value="Add Invoice" />
	      <input name="htbxVendorID" type="hidden" id="htbxVendorID" value="<%=lngVendorID%>" /></td>
	    <td>&nbsp;</td>
      </tr>
      <input type="hidden" name="MM_insert" value="frmInvoice" />
      </form>
	  <tr>
	    <td colspan="7"><hr /></td>
      </tr>
<%
		End If
		curTotalInvoices = 0
		Do While Not rstVendorInvoices.EOF
			If bolInvoicesEditGranted Then
				strEdit = "<a href=""VendorInvoiceEdit.asp?lngVendorInvoiceID=" & (rstVendorInvoices.Fields.Item("VendorInvoiceID").Value) & """>"
				strEditEnd = "</a>"
			Else
				strEdit = ""
				strEditEnd = ""
			End If
			If IsNull(rstVendorInvoices.Fields.Item("InvoiceAmount").Value) Then
				curInvoiceAmount = 0			
			Else
				curInvoiceAmount = CDbl(rstVendorInvoices.Fields.Item("InvoiceAmount").Value)
			
			End If
			curTotalInvoices = curTotalInvoices + curInvoiceAmount
%>      
	  <tr>
        <td><a href="VendorInvoiceInformation.asp?lngVendorInvoiceID=<%=(rstVendorInvoices.Fields.Item("VendorInvoiceID").Value)%>" class="row_info"></a></td>
        <td align="center"><%=strEdit & (rstVendorInvoices.Fields.Item("InvoiceDate").Value) & strEditEnd%></td>
		<td align="center"><%=strEdit & (rstVendorInvoices.Fields.Item("MethodName").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstVendorInvoices.Fields.Item("Notes").Value) & strEditEnd%></td>
		<td align="right"><%=strEdit & (rstVendorInvoices.Fields.Item("VendorRate").Value) & strEditEnd%></td>
		<td align="center"><%=FormatCurrency(curInvoiceAmount)%></td>
		<td>&nbsp;</td>
	  </tr>
<%
			rstVendorInvoices.MoveNext
		Loop
%>
	  <tr>
	    <td colspan="7"><hr /></td>
      </tr>
	  <tr>
	    <td>&nbsp;</td>
	    <td align="center">&nbsp;</td>
	    <td align="center">&nbsp;</td>
	    <td colspan="2" align="right"><strong>Total Invoices:</strong></td>
	    <td><%=FormatCurrency(curTotalInvoices)%></td>
	    <td>&nbsp;</td>
      </tr>
        <tr>
          <td>&nbsp;</td>
          <td colspan="7">&nbsp;</td>
        </tr>
<%
	End If
	If bolPaymentsViewGranted Then
		If bolPaymentsAddGranted Then
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
	    <th align="right"><h4>Amount Paid</h4></th>
	    <th align="right"><h4>Amount Credited</h4></th>
	    <th align="right"><h4>Service Fee</h4></th>
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
	    <td align="right"><input name="tbxAmountPaid" type="text" id="tbxAmountPaid" size="8" /></td>
	    <td align="right"><input name="tbxAmountCredited" type="text" id="tbxAmountCredited" size="8" /></td>
	    <td><input type="submit" name="btnAddpayment" id="btnAddpayment" value="Add Payment" />
	      <input name="htbxVendorID" type="hidden" id="htbxVendorID" value="<%=lngVendorID%>" /></td>
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
		curTotalFees = 0
		Do While Not rstVendorPayments.EOF
			If bolPaymentsEditGranted Then
				strEdit = "<a href=""VendorPaymentEdit.asp?lngVendorPaymentID=" & (rstVendorPayments.Fields.Item("VendorPaymentID").Value) & """>"
				strEditEnd = "</a>"
			Else
				strEdit = ""
				strEditEnd = ""
			End If
			curTotalPayments = curTotalPayments + CDbl(rstVendorPayments.Fields.Item("AmountPaid").Value)
			curTotalCredits = curTotalCredits + CDbl(rstVendorPayments.Fields.Item("AmountCredited").Value)
%>      
	  <tr>
		<td>&nbsp;</td>
        <td align="center"><%=strEdit & (rstVendorPayments.Fields.Item("PaymentDate").Value) & strEditEnd%></td>
		<td align="center"><%=strEdit & (rstVendorPayments.Fields.Item("MethodName").Value) & strEditEnd%></td>
		<td align="right"><%=FormatCurrency(rstVendorPayments.Fields.Item("AmountPaid").Value)%></td>
		<td align="right"><%=FormatCurrency(rstVendorPayments.Fields.Item("AmountCredited").Value)%></td>
		<td align="right"><%=FormatCurrency(CDbl(rstVendorPayments.Fields.Item("AmountPaid").Value) - CDbl(rstVendorPayments.Fields.Item("AmountCredited").Value))%></td>
		<td>&nbsp;</td>
	  </tr>
<%
			rstVendorPayments.MoveNext
		Loop
%>
	  <tr>
	    <td colspan="7"><hr /></td>
      </tr>
	  <tr>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td align="right"><strong>Total Credits:</strong></td>
	    <td><%=FormatCurrency(curTotalCredits)%></td>
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
                    <td align="center"><strong>Account Balance: </strong><%=FormatCurrency(curTotalCredits - curTotalInvoices)%> </td>
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
                    <td colspan="10">Certain &quot;Vendors&quot; permissions are required to view this information.</td>
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
rstVendors.Close()
Set rstVendors = Nothing
%>
<%
rstVendorPayments.Close()
Set rstVendorPayments = Nothing
%>
<%
rstPaymentMethods.Close()
Set rstPaymentMethods = Nothing
%>
<%
rstVendorInvoices.Close()
Set rstVendorInvoices = Nothing
%>
<%
rstWorkHistorys.Close()
Set rstWorkHistorys = Nothing
%>
<%
rstProjectDetails.Close()
Set rstProjectDetails = Nothing
%>
