<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
strWhere = " WHERE (Invoices.Sent = 0 OR Invoices.Sent = 1)"
If Request.QueryString("tbxInvoiceFromDate") <> "" Then
	strWhere = strWhere & " AND Invoices.InvoiceDate >= '" & Request.QueryString("tbxInvoiceFromDate") & "'"
End If
If Request.QueryString("tbxInvoiceToDate") <> "" Then
	strWhere = strWhere & " AND Invoices.InvoiceDate <= '" & Request.QueryString("tbxInvoiceToDate") & "'"
End If
If Request.QueryString("cbxClientID") <> "" Then
	strWhere = strWhere & " AND Invoices.ClientID = " & CLng(Request.QueryString("cbxClientID"))
End If
If Request.QueryString("cbxPaymentMethodID") <> "" Then
	strWhere = strWhere & " AND Invoices.PaymentMethodID = " & CLng(Request.QueryString("cbxPaymentMethodID"))
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
If (CStr(Request("MM_insert")) = "frmAdd") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Invoices (InvoiceDate, ClientID, PaymentMethodID) VALUES (?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 135, 1, -1, MM_IIF(Request.Form("tbxInvoiceDate"), Request.Form("tbxInvoiceDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxClientID"), Request.Form("cbxClientID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("cbxPaymentMethodID"), Request.Form("cbxPaymentMethodID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim rstInvoices
Dim rstInvoices_cmd
Dim rstInvoices_numRows

Set rstInvoices_cmd = Server.CreateObject ("ADODB.Command")
rstInvoices_cmd.ActiveConnection = MM_OBA_STRING
rstInvoices_cmd.CommandText = "SELECT Invoices.InvoiceID, Invoices.ClientID, Invoices.InvoiceDate, Invoices.Sent, Clients.ClientName, SUM(COALESCE (InvoiceDetails.Amount, 0)) AS InvoiceTotal, PaymentMethods.MethodName FROM Invoices INNER JOIN Clients ON Invoices.ClientID = Clients.ClientID INNER JOIN PaymentMethods ON Invoices.PaymentMethodID = PaymentMethods.PaymentMethodID LEFT OUTER JOIN InvoiceDetails ON Invoices.InvoiceID = InvoiceDetails.InvoiceID" & strWhere & " GROUP BY Invoices.InvoiceID, Invoices.ClientID, Invoices.InvoiceDate, Clients.ClientName, Invoices.Sent, PaymentMethods.MethodName ORDER BY Invoices.InvoiceDate DESC" 
rstInvoices_cmd.Prepared = true

Set rstInvoices = rstInvoices_cmd.Execute
rstInvoices_numRows = 0
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
Dim rstPaymentMethods
Dim rstPaymentMethods_cmd
Dim rstPaymentMethods_numRows

Set rstPaymentMethods_cmd = Server.CreateObject ("ADODB.Command")
rstPaymentMethods_cmd.ActiveConnection = MM_OBA_STRING
rstPaymentMethods_cmd.CommandText = "SELECT * FROM PaymentMethods ORDER BY MethodName" 
rstPaymentMethods_cmd.Prepared = true

Set rstPaymentMethods = rstPaymentMethods_cmd.Execute
rstPaymentMethods_numRows = 0
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
<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/themes/base/jquery-ui.css">
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/jquery-ui.min.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
$(function() {
	$("#tbxInvoiceDate").datepicker();
});
$(function() {
	$("#tbxInvoiceFromDate").datepicker();
});
$(function() {
	$("#tbxInvoiceToDate").datepicker();
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
	<h1>Invoice List</h1>	
	<table border="0" cellspacing="0" cellpadding="0" class="box">
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="left"><h4>Invoice Date</h4></th>
	    <th align="left"><h4>Client</h4></th>
	    <th align="left"><h4>Method</h4></th>
	    <th align="left"><h4>Sent</h4></th>
	    <th align="right"><h4>Invoice Total</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
      <form id="frmFilter" name="frmFilter" method="get" action="">
	  <tr>
	    <td align="left">&nbsp;</td>
	    <td align="left" nowrap="nowrap"><input name="tbxInvoiceFromDate" type="text" id="tbxInvoiceFromDate" tabindex="0" size="11" placeholder="Start Date" value="<%=Request.QueryString("tbxInvoiceFromDate")%>" />
	    -
          <input name="tbxInvoiceToDate" type="text" id="tbxInvoiceToDate" tabindex="0" size="11" placeholder="End Date" value="<%=Request.QueryString("tbxInvoiceToDate")%>" /></td>
	    <td align="left"><select name="cbxClientID" id="cbxClientID">
	      <option value="" <%If (Not isNull(Request.QueryString("cbxClientID"))) Then If ("" = CStr(Request.QueryString("cbxClientID"))) Then Response.Write("selected=""selected""") : Response.Write("")%>>All Clients</option>
	      <%
While (NOT rstClients.EOF)
%>
	      <option value="<%=(rstClients.Fields.Item("ClientID").Value)%>" <%If (Not isNull(Request.QueryString("cbxClientID"))) Then If (CStr(rstClients.Fields.Item("ClientID").Value) = CStr(Request.QueryString("cbxClientID"))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstClients.Fields.Item("ClientName").Value)%></option>
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
	    <td align="left"><select name="cbxPaymentMethodID" id="cbxPaymentMethodID">
	      <option value="" <%If (Not isNull(Request.Querystring("cbxPaymentMethodID"))) Then If ("" = CStr(Request.Querystring("cbxPaymentMethodID"))) Then Response.Write("selected=""selected""") : Response.Write("")%>>All Methods</option>
	      <%
While (NOT rstPaymentMethods.EOF)
%>
	      <option value="<%=(rstPaymentMethods.Fields.Item("PaymentMethodID").Value)%>" <%If (Not isNull(Request.Querystring("cbxPaymentMethodID"))) Then If (CStr(rstPaymentMethods.Fields.Item("PaymentMethodID").Value) = CStr(Request.Querystring("cbxPaymentMethodID"))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstPaymentMethods.Fields.Item("MethodName").Value)%></option>
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
	    <td align="left"><input type="submit" name="btnFilter" id="btnFilter" value="Filter" /></td>
	    <td align="right">&nbsp;</td>
	    <td align="left">&nbsp;</td>
      </tr>
      </form>
<%
If bolInvoicesViewGranted Then
    If bolInvoicesAddGranted Then
%>
      <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
	  <tr>
	    <td colspan="7"><hr /></td>
	    </tr>
	  <tr>
	    <td>&nbsp;</td>
	    <td><input name="tbxInvoiceDate" type="text" id="tbxInvoiceDate" tabindex="0" value="<%=Date%>" size="11" /></td>
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
	    <td><select name="cbxPaymentMethodID" id="cbxPaymentMethodID">
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
	    <td><input type="submit" name="btnAdd" id="btnAdd" value="Add Invoice" /></td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
      </tr>
      <input type="hidden" name="MM_insert" value="frmAdd" />
      </form>
	  <tr>
	    <td colspan="7"><hr /></td>
      </tr>
<%
    End If
	Do While Not rstInvoices.EOF
		'If bolInvoicesEditGranted AND rstInvoices.Fields.Item("Sent").Value <> "True" Then
		If bolInvoicesEditGranted Then
			strEdit = "<a href=""InvoiceEdit.asp?lngInvoiceID=" & (rstInvoices.Fields.Item("InvoiceID").Value) & """>"
			strEditEnd = "</a>"
		Else
			strEdit = ""
			strEditEnd = ""
		End If 
%>      
	  <tr>
        <td><a href="InvoiceInformation.asp?lngInvoiceID=<%=(rstInvoices.Fields.Item("InvoiceID").Value)%>" class="row_info"></a></td>
		<td><%=strEdit & (rstInvoices.Fields.Item("InvoiceDate").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstInvoices.Fields.Item("ClientName").Value) & strEditEnd%></td>
		<td><%=(rstInvoices.Fields.Item("MethodName").Value)%></td>
		<td><%=strEdit & (rstInvoices.Fields.Item("Sent").Value) & strEditEnd%></td>
		<td align="right"><%=FormatCurrency(rstInvoices.Fields.Item("InvoiceTotal").Value)%></td>
		<td>&nbsp;</td>
	  </tr>
<%
        rstInvoices.MoveNext
    Loop
Else
%>  
        <tr>
            <td colspan="7">Viewing this list requires certain &quot;Invoices&quot; permissions</td>
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
rstInvoices.Close()
Set rstInvoices = Nothing
%>
<%
rstClients.Close()
Set rstClients = Nothing
%>
<%
rstPaymentMethods.Close()
Set rstPaymentMethods = Nothing
%>
