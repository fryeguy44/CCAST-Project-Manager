<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngInvoiceID
Dim strReturnPath

lngInvoiceID = Request.QueryString("lngInvoiceID")
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
    MM_editCmd.CommandText = "UPDATE dbo.Invoices SET InvoiceDate = ?, ClientID = ?, PaymentMethodID = ?, Sent = ? WHERE InvoiceID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 135, 1, -1, MM_IIF(Request.Form("tbxInvoiceDate"), Request.Form("tbxInvoiceDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxClientID"), Request.Form("cbxClientID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("cbxPaymentMethodID"), Request.Form("cbxPaymentMethodID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("chkSent"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
    MM_editCmd.CommandText = "DELETE FROM dbo.Invoices WHERE InvoiceID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If

End If
%>
<%
Dim rstInvoices__lngInvoiceID
rstInvoices__lngInvoiceID = "1"
If (lngInvoiceID <> "") Then 
  rstInvoices__lngInvoiceID = lngInvoiceID
End If
%>
<%
Dim rstInvoices
Dim rstInvoices_cmd
Dim rstInvoices_numRows

Set rstInvoices_cmd = Server.CreateObject ("ADODB.Command")
rstInvoices_cmd.ActiveConnection = MM_OBA_STRING
rstInvoices_cmd.CommandText = "SELECT TOP (1) Invoices.InvoiceID, Invoices.ClientID, Invoices.PaymentMethodID, Invoices.InvoiceDate, Invoices.Sent, InvoiceDetails.InvoiceDetailID FROM Invoices LEFT OUTER JOIN InvoiceDetails ON Invoices.InvoiceID = InvoiceDetails.InvoiceID WHERE Invoices.InvoiceID = ?" 
rstInvoices_cmd.Prepared = true
rstInvoices_cmd.Parameters.Append rstInvoices_cmd.CreateParameter("param1", 5, 1, -1, rstInvoices__lngInvoiceID) ' adDouble

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
rstPaymentMethods_cmd.CommandText = "SELECT PaymentMethodID, MethodName FROM PaymentMethods" 
rstPaymentMethods_cmd.Prepared = true

Set rstPaymentMethods = rstPaymentMethods_cmd.Execute
rstPaymentMethods_numRows = 0
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
	Response.Redirect("Invoices.asp")
End If

%>
<!-- jQuery UI -->
<link href="/global/jquery/css/ui-lightness/jquery-ui-1.10.3.custom.css" rel="stylesheet" />
<script type="text/javascript" src="/global/jquery/js/jquery-1.9.1.js"></script>
<script type="text/javascript" src="/global/jquery/js/jquery-ui-1.10.3.custom.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
$(function() {
	$("#tbxInvoiceDate").datepicker();
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
If bolInvoicesEditGranted Then
	If rstInvoices.EOF Then
%>  
        <tr>
          <th colspan="4">&nbsp;</th>
        </tr>
        <tr>
            <td colspan="4"><a href="Invoices.asp">The Invoice you are attempting to edit has been deleted. Click here to return to the Invoice List page</a></td>
        </tr>
<%
	Else
%>     
   	  <form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
        <tr>
            <td width="10">&nbsp;</td>
            <td align="right"><strong>Invoice Date</strong></td>
          <td><input name="tbxInvoiceDate" type="text" id="tbxInvoiceDate" value="<%=(rstInvoices.Fields.Item("InvoiceDate").Value)%>" size="11" /></td>
		<td>&nbsp;</td>
		</tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Client</strong></td>
          <td><select name="cbxClientID" id="cbxClientID">
            <%
While (NOT rstClients.EOF)
%>
            <option value="<%=(rstClients.Fields.Item("ClientID").Value)%>" <%If (Not isNull((rstInvoices.Fields.Item("ClientID").Value))) Then If (CStr(rstClients.Fields.Item("ClientID").Value) = CStr((rstInvoices.Fields.Item("ClientID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstClients.Fields.Item("ClientName").Value)%></option>
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
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Invoice Method</strong></td>
          <td><select name="cbxPaymentMethodID" id="cbxPaymentMethodID">
            <%
While (NOT rstPaymentMethods.EOF)
%>
            <option value="<%=(rstPaymentMethods.Fields.Item("PaymentMethodID").Value)%>" <%If (Not isNull((rstInvoices.Fields.Item("PaymentMethodID").Value))) Then If (CStr(rstPaymentMethods.Fields.Item("PaymentMethodID").Value) = CStr((rstInvoices.Fields.Item("PaymentMethodID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstPaymentMethods.Fields.Item("MethodName").Value)%></option>
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
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Sent</strong></td>
          <td><input <%If (CStr((rstInvoices.Fields.Item("Sent").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="chkSent" id="chkSent" /></td>
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
        <input type="hidden" name="MM_recordId" value="<%= rstInvoices.Fields.Item("InvoiceID").Value %>" />
        </form>
<%
		If bolInvoicesDeleteGranted AND IsNull(rstInvoices.Fields.Item("InvoiceDetailID").Value) Then
%>                
      <tr>
        <td width="10">&nbsp;</td>
            <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
              <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
              <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
              <input type="hidden" name="MM_delete" value="frmDelete" />
              <input type="hidden" name="MM_recordId" value="<%= rstInvoices.Fields.Item("InvoiceID").Value %>" />
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
            <td colspan="4">Certain &quot;Invoices&quot; permissions are required to perform this task.</td>
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
