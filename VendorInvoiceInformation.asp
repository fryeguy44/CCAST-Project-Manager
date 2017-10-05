<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngVendorInvoiceID
lngVendorInvoiceID = Request.QueryString("lngVendorInvoiceID")
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
If (CStr(Request("MM_insert")) = "frmAddWorkHistory") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
	Dim MM_editCmd
	
	Set MM_editCmd = Server.CreateObject ("ADODB.Command")
	MM_editCmd.ActiveConnection = MM_OBA_STRING
		MM_editCmd.CommandText = "INSERT INTO VendorInvoiceDetails  (WorkHistoryID, Time, Amount, VendorInvoiceID) SELECT WorkHistorys.WorkHistoryID, WorkHistorys.Hours, ? * WorkHistorys.Hours AS Amount, ? AS VendorInvoiceID FROM WorkHistorys WHERE (WorkHistorys.WorkHistoryID = ?)"  
    	MM_editCmd.Prepared = true
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("htbxInvoiceRate"), Request.Form("htbxInvoiceRate"), 0)) ' adDouble
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("htbxVendorInvoiceID"), Request.Form("htbxVendorInvoiceID"), null)) ' adDouble
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("cbxWorkHistoryID"), Request.Form("cbxWorkHistoryID"), 0)) ' adDouble
	MM_editCmd.Execute
	MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_insert")) = "frmAddInvoiceDetail") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    'Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
	MM_editCmd.CommandText = "INSERT INTO dbo.VendorInvoiceDetails (DetailDescription, [Time], Amount, VendorInvoiceID) VALUES (?, ?, ?, ?)" 
	MM_editCmd.Prepared = true
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 500, Request.Form("tbxDetailDescription")) ' adVarWChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("tbxTime"), Request.Form("tbxTime"), 0)) ' adDouble
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxAmount"), Request.Form("tbxAmount"), 0)) ' adDouble
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("htbxVendorInvoiceID"), Request.Form("htbxVendorInvoiceID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim rstVendorInvoices__lngVendorInvoiceID
rstVendorInvoices__lngVendorInvoiceID = "1"
If (lngVendorInvoiceID <> "") Then 
  rstVendorInvoices__lngVendorInvoiceID = lngVendorInvoiceID
End If
%>
<%
Dim rstVendorInvoices
Dim rstVendorInvoices_cmd
Dim rstVendorInvoices_numRows

Set rstVendorInvoices_cmd = Server.CreateObject ("ADODB.Command")
rstVendorInvoices_cmd.ActiveConnection = MM_OBA_STRING
rstVendorInvoices_cmd.CommandText = "SELECT VendorInvoices.VendorInvoiceID, VendorInvoices.InvoiceDate, Vendors.VendorName, VendorInvoices.VendorID, VendorInvoices.VendorRate, PaymentMethods.MethodName FROM VendorInvoices INNER JOIN Vendors ON Vendors.VendorID = VendorInvoices.VendorID INNER JOIN PaymentMethods ON VendorInvoices.PaymentMethodID = PaymentMethods.PaymentMethodID WHERE VendorInvoiceID = ?" 
rstVendorInvoices_cmd.Prepared = true
rstVendorInvoices_cmd.Parameters.Append rstVendorInvoices_cmd.CreateParameter("param1", 5, 1, -1, rstVendorInvoices__lngVendorInvoiceID) ' adDouble

Set rstVendorInvoices = rstVendorInvoices_cmd.Execute
rstVendorInvoices_numRows = 0
%>
<%
Dim rstVendorInvoiceDetails__lngVendorInvoiceID
rstVendorInvoiceDetails__lngVendorInvoiceID = "1"
If (lngVendorInvoiceID <> "") Then 
  rstVendorInvoiceDetails__lngVendorInvoiceID = lngVendorInvoiceID
End If
%>
<%
Dim rstVendorInvoiceDetails
Dim rstVendorInvoiceDetails_cmd
Dim rstVendorInvoiceDetails_numRows

Set rstVendorInvoiceDetails_cmd = Server.CreateObject ("ADODB.Command")
rstVendorInvoiceDetails_cmd.ActiveConnection = MM_OBA_STRING
rstVendorInvoiceDetails_cmd.CommandText = "SELECT VendorInvoiceDetails.VendorInvoiceDetailID, VendorInvoiceDetails.VendorInvoiceID, VendorInvoiceDetails.WorkHistoryID, VendorInvoiceDetails.Time, VendorInvoiceDetails.Amount,  VendorInvoiceDetails.DetailDescription, WorkHistorys.WorkDescription, CASE WHEN VendorInvoiceDetails.DetailDescription IS NOT NULL AND VendorInvoiceDetails.DetailDescription <> ''  THEN VendorInvoiceDetails.DetailDescription ELSE WorkHistorys.WorkDescription + ' - ' + CAST(WorkHistorys.WorkDate AS nvarchar(15)) END AS Description FROM VendorInvoiceDetails LEFT OUTER JOIN WorkHistorys ON VendorInvoiceDetails.WorkHistoryID = WorkHistorys.WorkHistoryID WHERE VendorInvoiceDetails.VendorInvoiceID = ? ORDER BY WorkHistorys.WorkDate" 
rstVendorInvoiceDetails_cmd.Prepared = true
rstVendorInvoiceDetails_cmd.Parameters.Append rstVendorInvoiceDetails_cmd.CreateParameter("param1", 5, 1, -1, rstVendorInvoiceDetails__lngVendorInvoiceID) ' adDouble

Set rstVendorInvoiceDetails = rstVendorInvoiceDetails_cmd.Execute
rstVendorInvoiceDetails_numRows = 0
%>
<%
lngVendorID = rstVendorInvoices.Fields.Item("VendorID").Value
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
rstWorkHistorys_cmd.CommandText = "SELECT WorkHistorys.WorkHistoryID, LEFT(WorkHistorys.WorkDescription, 150) + ' - ' + CAST(WorkHistorys.WorkDate AS nvarchar(11))  AS WorkDescription FROM WorkHistorys INNER JOIN ProjectDetails ON WorkHistorys.ProjectDetailID = ProjectDetails.ProjectDetailID INNER JOIN Projects ON ProjectDetails.ProjectID = Projects.ProjectID LEFT OUTER JOIN VendorInvoiceDetails ON WorkHistorys.WorkHistoryID = VendorInvoiceDetails.WorkHistoryID WHERE (WorkHistorys.VendorID =  ?) AND (VendorInvoiceDetails.VendorInvoiceDetailID IS NULL) ORDER BY WorkHistorys.WorkDate" 
rstWorkHistorys_cmd.Prepared = true
rstWorkHistorys_cmd.Parameters.Append rstWorkHistorys_cmd.CreateParameter("param1", 5, 1, -1, rstWorkHistorys__lngVendorID) ' adDouble

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
If bolInvoicesViewGranted Then
	If bolInvoicesEditGranted Then
		strEditLink = "<a href=""VendorInvoiceEdit.asp?lngVendorInvoiceID=" & (rstVendorInvoices.Fields.Item("VendorInvoiceID").Value) & """>"
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
                    <td align="right"><strong>Vendor</strong></td>
                    <td><%=strEditLink & (rstVendorInvoices.Fields.Item("VendorName").Value) & strEndEditLink%></td>
                    <td align="right"><strong>Rate</strong></td>
                    <td><%=strEditLink & FormatCurrency(rstVendorInvoices.Fields.Item("VendorRate").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Invoice Date</strong></td>
                    <td><%=strEditLink & (rstVendorInvoices.Fields.Item("InvoiceDate").Value) & strEndEditLink%></td>
                    <td align="right"><strong>Invoice Delivery Method</strong></td>
                    <td><%=strEditLink & (rstVendorInvoices.Fields.Item("MethodName").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td colspan="6"><hr /></td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="4">&nbsp;
                	  <table border="0" cellspacing="0" cellpadding="0" class="info">
                        <tr>
                          <th align="left"><h4>Project Description</h4></th>
                          <th align="left"><h4>Invoice Detail</h4></th>
                          <th><h4>Time</h4></th>
                          <th align="right"><h4>Amount</h4></th>
                          <th align="right">&nbsp;</th>
                        </tr>
                        <form id="frmAddWorkHistory" name="frmAddWorkHistory" method="POST" action="<%=MM_editAction%>">
                        <tr>
                          <td colspan="4"><select name="cbxWorkHistoryID" id="cbxWorkHistoryID">
                            <%
While (NOT rstWorkHistorys.EOF)
%>
                            <option value="<%=(rstWorkHistorys.Fields.Item("WorkHistoryID").Value)%>"><%=(rstWorkHistorys.Fields.Item("WorkDescription").Value)%></option>
                            <%
  rstWorkHistorys.MoveNext()
Wend
If (rstWorkHistorys.CursorType > 0) Then
  rstWorkHistorys.MoveFirst
Else
  rstWorkHistorys.Requery
End If
%>
                          </select>
                          <input name="htbxInvoiceRate" type="hidden" id="htbxInvoiceRate" value="<%=rstVendorInvoices.Fields.Item("VendorRate").Value%>" /></td>
                          <td><input type="submit" name="btnAddWorkHistory" id="btnAddWorkHistory" value="Add" /></td>
                        </tr>
                        <input type="hidden" name="MM_insert" value="frmAddWorkHistory" />
                        <input name="htbxVendorInvoiceID" type="hidden" id="htbxVendorInvoiceID" value="<%=lngVendorInvoiceID%>" />
                       </form>
                        <tr>
                          <th colspan="6">OR:</th>
                          </tr>
                        <form id="frmAddInvoiceDetail" name="frmAddInvoiceDetail" method="POST" action="<%=MM_editAction%>">
                        <tr>
                          <td>&nbsp;</td>
                          <td><input name="tbxDetailDescription" type="text" id="tbxDetailDescription" size="75" /></td>
                          <td align="center"><input name="tbxTime" type="text" id="tbxTime" size="5" style="text-align:center" /></td>
                          <td align="right"><input name="tbxAmount" type="text" id="tbxAmount" size="10" style="text-align:right;" /></td>
                          <td><input name="htbxVendorInvoiceID" type="hidden" id="htbxVendorInvoiceID" value="<%=lngVendorInvoiceID%>" />
                          <input type="submit" name="btnAdd" id="btnAdd" value="Add" /></td>
                          <td>&nbsp;</td>
                        </tr>
                        <input type="hidden" name="MM_insert" value="frmAddInvoiceDetail" />
                        </form>
                        <tr>
                          <td colspan="6"><hr /></td>
                        </tr>

<%
	curLineTotal = 0
	curVendorInvoiceTotal = 0
	dblTimeTotal = 0
	Do While Not rstVendorInvoiceDetails.EOF
		If bolInvoicesEditGranted Then
			strEditLink = "<a href=""VendorInvoiceDetailEdit.asp?lngVendorInvoiceDetailID=" & (rstVendorInvoiceDetails.Fields.Item("VendorInvoiceDetailID").Value) & """>"
			strEndEditLink = "</a>&nbsp;"
		Else
			strEditLink = ""
			strEndEditLink = "&nbsp;"
		End If
		If IsNull(rstVendorInvoiceDetails.Fields.Item("Time").Value) Then
			dblTime = 0
		Else
			dblTime = rstVendorInvoiceDetails.Fields.Item("Time").Value
			dblTimeTotal = dblTimeTotal + CDbl(rstVendorInvoiceDetails.Fields.Item("Time").Value)
		End If
		
		If IsNull(rstVendorInvoices.Fields.Item("VendorRate").Value) Then
			dblRate = 0
		Else
			dblRate = rstVendorInvoices.Fields.Item("VendorRate").Value
		End If
		
		
		curLineTotal = rstVendorInvoiceDetails.Fields.Item("Amount").Value
		curVendorInvoiceTotal = curVendorInvoiceTotal + curLineTotal
		
%>                        
                        <tr>
                          <td>&nbsp;</td>
                          <td><%=strEditLink &(rstVendorInvoiceDetails.Fields.Item("Description").Value) & strEndEditLink%></td>
                          <td align="center"><%=strEditLink & FormatNumber(dblTime, 1, -1) & strEndEditLink%></td>
                          <td align="right"><%=strEditLink &FormatCurrency(curLineTotal) & strEndEditLink%></td>
                          <td align="right">&nbsp;</td>
                        </tr>                        
<%
		rstVendorInvoiceDetails.MoveNext
	Loop
		
%>                        
                        <tr>
                          <td colspan="5"><hr /></td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                          <td align="right"><strong>Invoice Hours:</strong></td>
                          <td align="right"><%=FormatNumber(dblTimeTotal, 1)%></td>
                          <td align="right"><%=FormatCurrency(curVendorInvoiceTotal)%></td>
                          <td align="right">&nbsp;</td>
                        </tr>
                      </table>
                    
                    </td>
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
Else
%>               

                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="10">Certain &quot;Invoices&quot; permissions are required to view this information.</td>
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
rstVendorInvoices.Close()
Set rstVendorInvoices = Nothing
%>
<%
rstVendorInvoiceDetails.Close()
Set rstVendorInvoiceDetails = Nothing
%>
<%
rstWorkHistorys.Close()
Set rstWorkHistorys = Nothing
%>
