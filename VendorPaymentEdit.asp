<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngVendorPaymentID
Dim strReturnPath

lngVendorPaymentID = Request.QueryString("lngVendorPaymentID")
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
    MM_editCmd.CommandText = "UPDATE dbo.VendorPayments SET PaymentDate = ?, PaymentMethodID = ?, AmountPaid = ?, AmountCredited = ? WHERE VendorPaymentID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 135, 1, -1, MM_IIF(Request.Form("tbxPaymentDate"), Request.Form("tbxPaymentDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxPaymentMethodID"), Request.Form("cbxPaymentMethodID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("tbxAmountPaid"), Request.Form("tbxAmountPaid"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxAmountCredited"), Request.Form("tbxAmountCredited"), null)) ' adDouble
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
    MM_editCmd.CommandText = "DELETE FROM dbo.VendorPayments WHERE VendorPaymentID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If

End If
%>
<%
Dim rstVendorPayments__lngVendorPaymentID 
rstVendorPayments__lngVendorPaymentID  = "1"
If (lngVendorPaymentID  <> "") Then 
  rstVendorPayments__lngVendorPaymentID  = lngVendorPaymentID 
End If
%>
<%
Dim rstVendorPayments
Dim rstVendorPayments_cmd
Dim rstVendorPayments_numRows

Set rstVendorPayments_cmd = Server.CreateObject ("ADODB.Command")
rstVendorPayments_cmd.ActiveConnection = MM_OBA_STRING
rstVendorPayments_cmd.CommandText = "SELECT VendorPayments.VendorPaymentID, VendorPayments.VendorID, VendorPayments.PaymentMethodID, VendorPayments.PaymentDate, VendorPayments.AmountPaid, VendorPayments.AmountCredited,  Vendors.VendorName FROM VendorPayments  INNER JOIN Vendors ON VendorPayments.VendorID = Vendors.VendorID WHERE VendorPaymentID = ?" 
rstVendorPayments_cmd.Prepared = true
rstVendorPayments_cmd.Parameters.Append rstVendorPayments_cmd.CreateParameter("param1", 5, 1, -1, rstVendorPayments__lngVendorPaymentID) ' adDouble

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
	Response.Redirect("VendorPayments.asp")
End If

%>
<!-- jQuery UI -->
<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/themes/base/jquery-ui.css" />
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/jquery-ui.min.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
$(function() {
	$("#tbxPaymentDate").datepicker();
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
If bolPaymentsEditGranted Then
	If rstVendorPayments.EOF Then
%>  
        <tr>
          <th colspan="4">&nbsp;</th>
        </tr>
        <tr>
            <td colspan="4"><a href="VendorPayments.asp">The VendorPayment you are attempting to edit has been deleted. Click here to return to the VendorPayment List page</a></td>
        </tr>
<%
	Else
%>     
    	<form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Vendor</strong></td>
          <td><%=(rstVendorPayments.Fields.Item("VendorName").Value)%></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Payment Date</strong></td>
          <td><input name="tbxPaymentDate" type="text" id="tbxPaymentDate" value="<%=(rstVendorPayments.Fields.Item("PaymentDate").Value)%>" size="11" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Payment Method</strong></td>
          <td><select name="cbxPaymentMethodID" id="cbxPaymentMethodID">
            <%
While (NOT rstPaymentMethods.EOF)
%>
            <option value="<%=(rstPaymentMethods.Fields.Item("PaymentMethodID").Value)%>" <%If (Not isNull((rstVendorPayments.Fields.Item("PaymentMethodID").Value))) Then If (CStr(rstPaymentMethods.Fields.Item("PaymentMethodID").Value) = CStr((rstVendorPayments.Fields.Item("PaymentMethodID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstPaymentMethods.Fields.Item("MethodName").Value)%></option>
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
          <td align="right"><strong>Amount Paid</strong></td>
          <td><input name="tbxAmountPaid" type="text" id="tbxAmountPaid" value="<%=(rstVendorPayments.Fields.Item("AmountPaid").Value)%>" size="8" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
            <td width="10">&nbsp;</td>
            <td align="right"><strong>Amount Credited</strong></td>
          <td><input name="tbxAmountCredited" type="text" id="tbxAmountCredited" value="<%=(rstVendorPayments.Fields.Item("AmountCredited").Value)%>" size="8" /></td>
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
        <input type="hidden" name="MM_recordId" value="<%= rstVendorPayments.Fields.Item("VendorPaymentID").Value %>" />
        </form>
<%
		If bolPaymentsDeleteGranted Then
%>                
      <tr>
        <td width="10">&nbsp;</td>
            <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
              <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
              <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
              <input type="hidden" name="MM_delete" value="frmDelete" />
              <input type="hidden" name="MM_recordId" value="<%= rstVendorPayments.Fields.Item("VendorPaymentID").Value %>" />
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
            <td colspan="4">Certain &quot;Payments&quot; permissions are required to perform this task.</td>
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
rstVendorPayments.Close()
Set rstVendorPayments = Nothing
%>
<%
rstPaymentMethods.Close()
Set rstPaymentMethods = Nothing
%>
