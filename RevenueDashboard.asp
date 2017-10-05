<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim rstIncomes
Dim rstIncomes_cmd
Dim rstIncomes_numRows

Set rstIncomes_cmd = Server.CreateObject ("ADODB.Command")
rstIncomes_cmd.ActiveConnection = MM_OBA_STRING
rstIncomes_cmd.CommandText = "SELECT MONTH(PaymentDate) AS PaymentMonth, YEAR(PaymentDate) AS PaymentYear, SUM(Amount) AS TotalReceived, COALESCE(Payments.TotalPaid, 0) AS TotalPaid FROM ClientPayments  LEFT OUTER JOIN (SELECT MONTH(PaymentDate) AS PaymentMonth, YEAR(PaymentDate) AS PaymentYear, SUM(Amount) AS TotalPaid FROM VendorPayments GROUP BY YEAR(PaymentDate), MONTH(PaymentDate)) AS Payments ON MONTH(ClientPayments.PaymentDate) = Payments.PaymentMonth AND YEAR(ClientPayments.PaymentDate) = Payments.PaymentYear GROUP BY YEAR(PaymentDate), Month(PaymentDate), Payments.TotalPaid ORDER BY PaymentYear DESC, PaymentMonth DESC" 
rstIncomes_cmd.Prepared = true

Set rstIncomes = rstIncomes_cmd.Execute
rstIncomes_numRows = 0
%>
<%
Dim rstClients
Dim rstClients_cmd
Dim rstClients_numRows

Set rstClients_cmd = Server.CreateObject ("ADODB.Command")
rstClients_cmd.ActiveConnection = MM_OBA_STRING
rstClients_cmd.CommandText = "SELECT Clients.ClientID, Clients.ClientName, SUM(ClientPayments.Amount) AS TotalReceived FROM ClientPayments INNER JOIN  Clients ON ClientPayments.ClientID = Clients.ClientID WHERE (ClientPayments.PaymentDate >= DATEADD(year, - 1, GETDATE())) GROUP BY Clients.ClientID, Clients.ClientName ORDER BY TotalReceived DESC" 
rstClients_cmd.Prepared = true

Set rstClients = rstClients_cmd.Execute
rstClients_numRows = 0
%>
<%
Dim rstVendors
Dim rstVendors_cmd
Dim rstVendors_numRows

Set rstVendors_cmd = Server.CreateObject ("ADODB.Command")
rstVendors_cmd.ActiveConnection = MM_OBA_STRING
rstVendors_cmd.CommandText = "SELECT Vendors.VendorID, Vendors.VendorName, SUM(VendorPayments.Amount) AS TotalPaid FROM VendorPayments INNER JOIN  Vendors ON VendorPayments.VendorID = Vendors.VendorID  WHERE (VendorPayments.PaymentDate >= DATEADD(year, - 1, GETDATE())) GROUP BY Vendors.VendorID, Vendors.VendorName ORDER BY TotalPaid DESC" 
rstVendors_cmd.Prepared = true

Set rstVendors = rstVendors_cmd.Execute
rstVendors_numRows = 0
%>
<%
Dim rstVendorBalances
Dim rstVendorBalances_cmd
Dim rstVendorBalances_numRows

Set rstVendorBalances_cmd = Server.CreateObject ("ADODB.Command")
rstVendorBalances_cmd.ActiveConnection = MM_OBA_STRING
rstVendorBalances_cmd.CommandText = "SELECT Vendors.VendorID, Vendors.VendorName, SUM(Amounts.Amount) AS Balance FROM Vendors INNER JOIN  (SELECT VendorInvoices.VendorID, -(VendorInvoiceDetails.Amount) As Amount FROM VendorInvoiceDetails INNER JOIN VendorInvoices ON VendorInvoiceDetails.VendorInvoiceID = VendorInvoices.VendorInvoiceID UNION ALL SELECT VendorID, Amount - ProcessFee AS Balance FROM VendorPayments) AS Amounts ON Vendors.VendorID = Amounts.VendorID GROUP BY Vendors.VendorID, VendorName HAVING SUM(Amounts.Amount) <> 0" 
rstVendorBalances_cmd.Prepared = true

Set rstVendorBalances = rstVendorBalances_cmd.Execute
rstVendorBalances_numRows = 0
%>
<%
Dim rstClientBalances
Dim rstClientBalances_cmd
Dim rstClientBalances_numRows

Set rstClientBalances_cmd = Server.CreateObject ("ADODB.Command")
rstClientBalances_cmd.ActiveConnection = MM_OBA_STRING
rstClientBalances_cmd.CommandText = "SELECT Clients.ClientID, Clients.ClientName, SUM(Amounts.Amount) AS Balance FROM Clients  INNER JOIN (SELECT Invoices.ClientID, InvoiceDetails.Amount FROM InvoiceDetails  INNER JOIN Invoices ON InvoiceDetails.InvoiceID = Invoices.InvoiceID WHERE Invoices.Sent = 1 UNION ALL SELECT ClientID, -(CreditedAmount) AS Balance FROM ClientPayments) AS Amounts ON Clients.ClientID = Amounts.ClientID  GROUP BY  Clients.ClientID, ClientName HAVING SUM(Amounts.Amount) <> 0 ORDER BY Balance DESC" 
rstClientBalances_cmd.Prepared = true

Set rstClientBalances = rstClientBalances_cmd.Execute
rstClientBalances_numRows = 0
%>
<%
Dim rstUnsentInvoices
Dim rstUnsentInvoices_cmd
Dim rstUnsentInvoices_numRows

Set rstUnsentInvoices_cmd = Server.CreateObject ("ADODB.Command")
rstUnsentInvoices_cmd.ActiveConnection = MM_OBA_STRING
rstUnsentInvoices_cmd.CommandText = "SELECT Invoices.InvoiceID, Invoices.ClientID, Invoices.PaymentMethodID, Invoices.InvoiceDate, Clients.ClientName, PaymentMethods.MethodName, SUM(InvoiceDetails.Amount) AS InvoiceTotal FROM Invoices INNER JOIN Clients ON Invoices.ClientID = Clients.ClientID INNER JOIN PaymentMethods ON Invoices.PaymentMethodID = PaymentMethods.PaymentMethodID INNER JOIN InvoiceDetails ON Invoices.InvoiceID = InvoiceDetails.InvoiceID WHERE (Invoices.Sent = 0) GROUP BY Invoices.InvoiceID, Invoices.ClientID, Invoices.PaymentMethodID, Invoices.InvoiceDate, Clients.ClientName, PaymentMethods.MethodName ORDER BY InvoiceDate" 
rstUnsentInvoices_cmd.Prepared = true

Set rstUnsentInvoices = rstUnsentInvoices_cmd.Execute
rstUnsentInvoices_numRows = 0
%>
<%
Dim rstUnbilledWorkHistorys
Dim rstUnbilledWorkHistorys_cmd
Dim rstUnbilledWorkHistorys_numRows

Set rstUnbilledWorkHistorys_cmd = Server.CreateObject ("ADODB.Command")
rstUnbilledWorkHistorys_cmd.ActiveConnection = MM_OBA_STRING
rstUnbilledWorkHistorys_cmd.CommandText = "SELECT Clients.ClientID, Clients.ClientName, SUM(WorkHistorys.Hours) * Clients.CurrentRate AS Amount FROM Projects INNER JOIN Clients ON Projects.ClientID = Clients.ClientID INNER JOIN WorkHistorys INNER JOIN ProjectDetails ON WorkHistorys.ProjectDetailID = ProjectDetails.ProjectDetailID ON Projects.ProjectID = ProjectDetails.ProjectID LEFT OUTER JOIN InvoiceDetails ON WorkHistorys.WorkHistoryID = InvoiceDetails.WorkHistoryID WHERE (InvoiceDetails.InvoiceDetailID IS NULL) AND Clients.CurrentRate > 0 AND WorkHistorys.Hours > 0 GROUP BY Clients.ClientID, Clients.ClientName, Clients.CurrentRate ORDER BY Clients.ClientName" 
rstUnbilledWorkHistorys_cmd.Prepared = true

Set rstUnbilledWorkHistorys = rstUnbilledWorkHistorys_cmd.Execute
rstUnbilledWorkHistorys_numRows = 0
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
	<h1>Revenue Dashboard</h1>	
<%
If bolFinancialsViewGranted Then

%>

	<table border="0" cellspacing="0" cellpadding="0" class="fluid">
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="left">&nbsp;</th>
	    <th align="left"><h4>&nbsp;</h4></th>
	    <th align="left">&nbsp;</th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
	  <tr>
	    <td>&nbsp;</td>
		<td><table border="0" cellspacing="0" cellpadding="0" class="box3">
		  <tr>
		    <th colspan="6" scope="col"><h2>Monthly Revenue</h2></th>
		    </tr>
		  <tr>
		    <th scope="col">&nbsp;</th>
		    <th align="right"><h4>Month</h4></th>
		    <th align="right"><h4>Received</h4></th>
		    <th align="right"><h4>Paid</h4></th>
		    <th align="right"><h4>Net</h4></th>
		    <th scope="col">&nbsp;</th>
		    </tr>
<%
	curTotalReceived = 0
	curTotalPaid = 0
	Do While Not rstIncomes.EOF
		curTotalReceived = curTotalReceived + CDbl(rstIncomes.Fields.Item("TotalReceived").Value)
		curTotalPaid = curTotalPaid + CDbl(rstIncomes.Fields.Item("TotalPaid").Value)
		curNet = CDbl(rstIncomes.Fields.Item("TotalReceived").Value) - CDbl(rstIncomes.Fields.Item("TotalPaid").Value)
%>            
		  <tr class="tr_hover">
		    <td>&nbsp;</td>
		    <td align="right"><%=(rstIncomes.Fields.Item("PaymentMonth").Value) & "/" & (rstIncomes.Fields.Item("PaymentYear").Value)%></td>
		    <td align="right"><%=FormatCurrency(rstIncomes.Fields.Item("TotalReceived").Value)%></td>
		    <td align="right"><%=FormatCurrency(rstIncomes.Fields.Item("TotalPaid").Value)%></td>
		    <td align="right"><%=FormatCurrency(curNet)%></td>
		    <td>&nbsp;</td>
		    </tr>
<%
        rstIncomes.MoveNext
    Loop
	
%>            
		  <tr>
		    <td colspan="6"><hr /></td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td align="right"><strong>Totals:</strong></td>
		    <td align="right"><%=FormatCurrency(curTotalReceived)%></td>
		    <td align="right"><%=FormatCurrency(curTotalPaid)%></td>
		    <td align="right"><%=FormatCurrency(curTotalReceived - curTotalPaid)%></td>
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
	    </table></td>
		<td><table border="0" cellspacing="0" cellpadding="0" class="box3">
		  <tr>
		    <th colspan="4" scope="col"><h2> 1 Year Revenue By Client</h2></th>
		    </tr>
		  <tr>
		    <th scope="col">&nbsp;</th>
		    <th align="left"><h4>Client</h4></th>
		    <th align="right"><h4>Received</h4></th>
		    <th scope="col">&nbsp;</th>
		    </tr>
<%
	curTotalReceived = 0
	Do While Not rstClients.EOF
		curTotalReceived = curTotalReceived + CDbl(rstClients.Fields.Item("TotalReceived").Value)
%>            
		  <tr class="tr_hover">
		    <td><a href="ClientInformation.asp?lngClientID=<%=rstClients.Fields.Item("ClientID").Value%>" class="row_info"></a></td>
		    <td><%=(rstClients.Fields.Item("ClientName").Value)%></td>
		    <td align="right"><%=FormatCurrency(rstClients.Fields.Item("TotalReceived").Value)%></td>
		    <td>&nbsp;</td>
		    </tr>
<%
        rstClients.MoveNext
    Loop
	
%>            
		  <tr>
		    <td colspan="4"><hr /></td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td align="right"><strong>Totals:</strong></td>
		    <td align="right"><%=FormatCurrency(curTotalReceived)%></td>
		    <td>&nbsp;</td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    </tr>
	    </table>
		<table border="0" cellspacing="0" cellpadding="0" class="box3">
		  <tr>
		    <th colspan="4" scope="col"><h2> Client Balances</h2></th>
		    </tr>
		  <tr>
		    <th scope="col">&nbsp;</th>
		    <th align="left"><h4>Client</h4></th>
		    <th align="right"><h4>Due</h4></th>
		    <th scope="col">&nbsp;</th>
		    </tr>
<%
	curTotalBalance = 0
	Do While Not rstClientBalances.EOF
		curTotalBalance = curTotalBalance + CDbl(rstClientBalances.Fields.Item("Balance").Value)
%>            
		  <tr class="tr_hover">
		    <td><a href="ClientInformation.asp?lngClientID=<%=rstClientBalances.Fields.Item("ClientID").Value%>" class="row_info"></a></td>
		    <td><%=(rstClientBalances.Fields.Item("ClientName").Value)%></td>
		    <td align="right"><%=FormatCurrency(rstClientBalances.Fields.Item("Balance").Value)%></td>
		    <td>&nbsp;</td>
		    </tr>
<%
        rstClientBalances.MoveNext
    Loop
	
%>            
		  <tr>
		    <td colspan="4"><hr /></td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td align="right"><strong>Totals:</strong></td>
		    <td align="right"><%=FormatCurrency(curTotalBalance)%></td>
		    <td>&nbsp;</td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    </tr>
	    </table>        
		<table border="0" cellspacing="0" cellpadding="0" class="box3">
		  <tr>
		    <th colspan="6" scope="col"><h2> Unsent Invoices</h2></th>
		    </tr>
		  <tr>
		    <th scope="col">&nbsp;</th>
		    <th align="left"><h4>Client</h4></th>
		    <th align="right"><h4>Date</h4></th>
		    <th align="right"><h4>Platform</h4></th>
		    <th align="right"><h4>Total</h4></th>
		    <th scope="col">&nbsp;</th>
		    </tr>
<%
	curTotalBalance = 0
	Do While Not rstUnsentInvoices.EOF
		If bolInvoicesEditGranted Then
			strEditLink = "<a href=""InvoiceEdit.asp?lngInvoiceID=" & (rstUnsentInvoices.Fields.Item("InvoiceID").Value) & """>"
			strEndEditLink = "</a>&nbsp;"
		Else
			strEditLink = ""
			strEndEditLink = "&nbsp;"
		End If
		
		curTotalBalance = curTotalBalance + CDbl(rstUnsentInvoices.Fields.Item("InvoiceTotal").Value)
%>            
		  <tr class="tr_hover">
		    <td><a href="InvoiceInformation.asp?lngInvoiceID=<%=rstUnsentInvoices.Fields.Item("InvoiceID").Value%>" class="row_info"></a></td>
		    <td nowrap="nowrap"><%=strEditLink & (rstUnsentInvoices.Fields.Item("ClientName").Value) & strEndEditLink%></td>
		    <td align="right"><%=strEditLink & (rstUnsentInvoices.Fields.Item("InvoiceDate").Value) & strEndEditLink%></td>
		    <td align="right"><%=strEditLink & (rstUnsentInvoices.Fields.Item("MethodName").Value) & strEndEditLink%></td>
		    <td align="right"><%=strEditLink & FormatCurrency(rstUnsentInvoices.Fields.Item("InvoiceTotal").Value) & strEndEditLink%></td>
		    <td>&nbsp;</td>
		    </tr>
<%
        rstUnsentInvoices.MoveNext
    Loop
	
%>            
		  <tr>
		    <td colspan="6"><hr /></td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td align="right">&nbsp;</td>
		    <td align="right">&nbsp;</td>
		    <td align="right"><strong>Total:</strong></td>
		    <td align="right"><%=FormatCurrency(curTotalBalance)%></td>
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
        </td>
		<td><table border="0" cellspacing="0" cellpadding="0" class="box3">
		  <tr>
		    <th colspan="4" scope="col"><h2>1 Year Expense By Vendor</h2></th>
		    </tr>
		  <tr>
		    <th scope="col">&nbsp;</th>
		    <th align="left"><h4>Vendor</h4></th>
		    <th align="right"><h4>Paid</h4></th>
		    <th scope="col">&nbsp;</th>
		    </tr>
<%
	curTotalPaid = 0
	Do While Not rstVendors.EOF
		curTotalPaid = curTotalPaid + CDbl(rstVendors.Fields.Item("TotalPaid").Value)
%>            
		  <tr class="tr_hover">
		    <td><a href="VendorInformation.asp?lngVendorID=<%=rstVendors.Fields.Item("VendorID").Value%>" class="row_info"></a></td>
		    <td><%=(rstVendors.Fields.Item("VendorName").Value)%></td>
		    <td align="right"><%=FormatCurrency(rstVendors.Fields.Item("TotalPaid").Value)%></td>
		    <td>&nbsp;</td>
		    </tr>
<%
        rstVendors.MoveNext
    Loop
	
%>            
		  <tr>
		    <td colspan="4"><hr /></td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td align="right"><strong>Totals:</strong></td>
		    <td align="right"><%=FormatCurrency(curTotalPaid)%></td>
		    <td>&nbsp;</td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    </tr>
	    </table>
		<table border="0" cellspacing="0" cellpadding="0" class="box3">
		  <tr>
		    <th colspan="4" scope="col"><h2>  Vendor Balances</h2></th>
		    </tr>
		  <tr>
		    <th scope="col">&nbsp;</th>
		    <th align="right"><h4>Vendor</h4></th>
		    <th align="right"><h4>Balance</h4></th>
		    <th scope="col">&nbsp;</th>
		    </tr>
<%
	curTotalBalance = 0
	Do While Not rstVendorBalances.EOF
		curTotalBalance = curTotalBalance + CDbl(rstVendorBalances.Fields.Item("Balance").Value)
%>            
		  <tr class="tr_hover">
		    <td><a href="VendorInformation.asp?lngVendorID=<%=rstVendorBalances.Fields.Item("VendorID").Value%>" class="row_info"></a></td>
		    <td align="right"><%=(rstVendorBalances.Fields.Item("VendorName").Value)%></td>
		    <td align="right"><%=FormatCurrency(rstVendorBalances.Fields.Item("Balance").Value)%></td>
		    <td>&nbsp;</td>
		    </tr>
<%
        rstVendorBalances.MoveNext
    Loop
	
%>            
		  <tr>
		    <td colspan="4"><hr /></td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td align="right"><strong>Totals:</strong></td>
		    <td align="right"><%=FormatCurrency(curTotalBalance)%></td>
		    <td>&nbsp;</td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    </tr>
	    </table>        
		<table border="0" cellspacing="0" cellpadding="0" class="box3">
		  <tr>
		    <th colspan="4" scope="col"><h2>Unbilled Work Histories</h2></th>
		    </tr>
		  <tr>
		    <th scope="col">&nbsp;</th>
		    <th align="right"><h4>Client</h4></th>
		    <th align="right"><h4>Amount</h4></th>
		    <th scope="col">&nbsp;</th>
		    </tr>
<%
	curTotalUnbilled = 0
	Do While Not rstUnbilledWorkHistorys.EOF
		curTotalUnbilled = curTotalUnbilled + CDbl(rstUnbilledWorkHistorys.Fields.Item("Amount").Value)
%>            
		  <tr class="tr_hover">
		    <td><a href="ClientInformation.asp?lngClientID=<%=rstUnbilledWorkHistorys.Fields.Item("ClientID").Value%>" class="row_info"></a></td>
		    <td><%=(rstUnbilledWorkHistorys.Fields.Item("ClientName").Value)%></td>
		    <td align="right"><%=FormatCurrency(rstUnbilledWorkHistorys.Fields.Item("Amount").Value)%></td>
		    <td>&nbsp;</td>
		    </tr>
<%
        rstUnbilledWorkHistorys.MoveNext
    Loop
	
%>            
		  <tr>
		    <td colspan="4"><hr /></td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td align="right"><strong>Totals:</strong></td>
		    <td align="right"><%=FormatCurrency(curTotalUnbilled)%></td>
		    <td>&nbsp;</td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    </tr>
	    </table>        
        </td>
		<td>&nbsp;</td>
	  </tr>
<%
Else
%>  
        <tr>
            <td colspan="5">Viewing this list requires certain &quot;Financials&quot; permissions</td>
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
rstIncomes.Close()
Set rstIncomes = Nothing
%>
<%
rstClients.Close()
Set rstClients = Nothing
%>
<%
rstVendors.Close()
Set rstVendors = Nothing
%>
<%
rstVendorBalances.Close()
Set rstVendorBalances = Nothing
%>
<%
rstClientBalances.Close()
Set rstClientBalances = Nothing
%>
<%
rstUnsentInvoices.Close()
Set rstUnsentInvoices = Nothing
%>
<%
rstUnbilledWorkHistorys.Close()
Set rstUnbilledWorkHistorys = Nothing
%>
