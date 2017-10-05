<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngInvoiceID
lngInvoiceID = Request.QueryString("lngInvoiceID")
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
	MM_editCmd.CommandText = "INSERT INTO InvoiceDetails (WorkHistoryID, [Time], Amount, InvoiceID) SELECT WorkHistorys.WorkHistoryID, WorkHistorys.Hours, WorkHistorys.Hours * Clients.CurrentRate, ? FROM WorkHistorys INNER JOIN ProjectDetails ON WorkHistorys.ProjectDetailID = ProjectDetails.ProjectDetailID INNER JOIN Projects ON ProjectDetails.ProjectID = Projects.ProjectID INNER JOIN Clients ON Projects.ClientID = Clients.ClientID WHERE (WorkHistorys.WorkHistoryID = ?)"  
	MM_editCmd.Prepared = true
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("htbxInvoiceID"), Request.Form("htbxInvoiceID"), null)) ' adDouble
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("cbxWorkHistoryID"), Request.Form("cbxWorkHistoryID"), 0)) ' adDouble
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
	MM_editCmd.CommandText = "INSERT INTO dbo.InvoiceDetails (DetailDescription, [Time], Amount, InvoiceID) VALUES (?, ?, ?, ?)" 
	MM_editCmd.Prepared = true
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 500, Request.Form("tbxDetailDescription")) ' adVarWChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("tbxTime"), Request.Form("tbxTime"), 0)) ' adDouble
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxAmount"), Request.Form("tbxAmount"), 0)) ' adDouble
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("htbxInvoiceID"), Request.Form("htbxInvoiceID"), null)) ' adDouble
	
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
rstInvoices_cmd.CommandText = "SELECT InvoiceID, InvoiceDate, Sent, ClientName, Invoices.ClientID FROM Invoices INNER JOIN Clients ON Clients.ClientID = Invoices.ClientID WHERE InvoiceID = ?" 
rstInvoices_cmd.Prepared = true
rstInvoices_cmd.Parameters.Append rstInvoices_cmd.CreateParameter("param1", 5, 1, -1, rstInvoices__lngInvoiceID) ' adDouble

Set rstInvoices = rstInvoices_cmd.Execute
rstInvoices_numRows = 0
%>
<%
Dim rstInvoiceDetails__lngInvoiceID
rstInvoiceDetails__lngInvoiceID = "1"
If (lngInvoiceID <> "") Then 
  rstInvoiceDetails__lngInvoiceID = lngInvoiceID
End If
%>
<%
Dim rstInvoiceDetails
Dim rstInvoiceDetails_cmd
Dim rstInvoiceDetails_numRows

Set rstInvoiceDetails_cmd = Server.CreateObject ("ADODB.Command")
rstInvoiceDetails_cmd.ActiveConnection = MM_OBA_STRING
rstInvoiceDetails_cmd.CommandText = "SELECT InvoiceDetails.InvoiceDetailID, InvoiceDetails.InvoiceID, InvoiceDetails.WorkHistoryID, InvoiceDetails.Time, InvoiceDetails.Rate, CASE WHEN InvoiceDetails.DetailDescription IS NOT NULL AND  InvoiceDetails.DetailDescription <> '' THEN InvoiceDetails.DetailDescription ELSE WorkHistorys.WorkDescription + ' - ' + CAST(WorkHistorys.WorkDate AS nvarchar(15)) END AS Description, InvoiceDetails.DetailDescription,  WorkHistorys.WorkDescription, InvoiceDetails.Amount FROM InvoiceDetails LEFT OUTER JOIN WorkHistorys ON InvoiceDetails.WorkHistoryID = WorkHistorys.WorkHistoryID WHERE InvoiceDetails.InvoiceID = ? ORDER BY WorkHistorys.WorkDate" 
rstInvoiceDetails_cmd.Prepared = true
rstInvoiceDetails_cmd.Parameters.Append rstInvoiceDetails_cmd.CreateParameter("param1", 5, 1, -1, rstInvoiceDetails__lngInvoiceID) ' adDouble

Set rstInvoiceDetails = rstInvoiceDetails_cmd.Execute
rstInvoiceDetails_numRows = 0
%>
<%
lngClientID = rstInvoices.Fields.Item("ClientID").Value
%>
<%
Dim rstWorkHistorys__lngClientID
rstWorkHistorys__lngClientID = "1"
If (lngClientID <> "") Then 
  rstWorkHistorys__lngClientID = lngClientID
End If
%>
<%
Dim rstWorkHistorys
Dim rstWorkHistorys_cmd
Dim rstWorkHistorys_numRows

Set rstWorkHistorys_cmd = Server.CreateObject ("ADODB.Command")
rstWorkHistorys_cmd.ActiveConnection = MM_OBA_STRING
rstWorkHistorys_cmd.CommandText = "SELECT WorkHistorys.WorkHistoryID, LEFT(WorkHistorys.WorkDescription, 150) + ' - ' + CAST(WorkHistorys.WorkDate AS nvarchar(11))  AS WorkDescription FROM WorkHistorys INNER JOIN ProjectDetails ON WorkHistorys.ProjectDetailID = ProjectDetails.ProjectDetailID INNER JOIN Projects ON ProjectDetails.ProjectID = Projects.ProjectID LEFT OUTER JOIN InvoiceDetails ON WorkHistorys.WorkHistoryID = InvoiceDetails.WorkHistoryID WHERE (Projects.ClientID =  ?) AND (InvoiceDetails.InvoiceDetailID IS NULL) ORDER BY WorkHistorys.WorkDate" 
rstWorkHistorys_cmd.Prepared = true
rstWorkHistorys_cmd.Parameters.Append rstWorkHistorys_cmd.CreateParameter("param1", 5, 1, -1, rstWorkHistorys__lngClientID) ' adDouble

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
		strEditLink = "<a href=""InvoiceEdit.asp?lngInvoiceID=" & (rstInvoices.Fields.Item("InvoiceID").Value) & """>"
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
                    <td><%=(rstInvoices.Fields.Item("ClientName").Value)%></td>
                    <td align="right">&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td align="right"><strong>Invoice Date</strong></td>
                    <td><%=strEditLink & (rstInvoices.Fields.Item("InvoiceDate").Value) & strEndEditLink%></td>
                    <td align="right"><strong>Sent</strong></td>
                    <td><%=strEditLink & (rstInvoices.Fields.Item("Sent").Value) & strEndEditLink%></td>
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
                          <th align="left"><h4>Work History</h4></th>
                          <th align="left"><h4>Invoice Detail Description</h4></th>
                          <th><h4>Qty</h4></th>
                          <th align="right"><h4>Amount</h4></th>
                          <th align="right">&nbsp;</th>
                        </tr>
<%
	If (rstInvoices.Fields.Item("Sent").Value) = "False" Then
		If Not rstWorkHistorys.EOF Then
%>                        
                        <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
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
                          </select></td>
                          <td><input type="submit" name="btnAddWorkHistory" id="btnAddWorkHistory" value="Add" /></td>
                        </tr>
                        <input type="hidden" name="MM_insert" value="frmAddWorkHistory" />
                        <input name="htbxInvoiceID" type="hidden" id="htbxInvoiceID" value="<%=lngInvoiceID%>" />
                       </form>
<%
		End If
%>                       
                        <tr>
                          <th colspan="5">OR:</th>
                          </tr>
                        <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
                        <tr>
                          <td>&nbsp;</td>
                          <td><input name="tbxDetailDescription" type="text" id="tbxDetailDescription" size="75" /></td>
                          <td align="center"><input name="tbxTime" type="text" id="tbxTime" size="5" style="text-align:center" /></td>
                          <td align="right" nowrap="nowrap">$
                            <input name="tbxAmount" type="text" id="tbxAmount" size="8" style="text-align:right;" /></td>
                          <td><input type="submit" name="btnAdd" id="btnAdd" value="Add" />
                          <input name="htbxInvoiceID" type="hidden" id="htbxInvoiceID" value="<%=lngInvoiceID%>" /></td>
                        </tr>
                        <input type="hidden" name="MM_insert" value="frmAddInvoiceDetail" />
                        </form>
                        <tr>
                          <td colspan="5"><hr /></td>
                        </tr>
<%
	End If
	curLineTotal = 0
	curInvoiceTotal = 0
	dblTimeTotal = 0
	Do While Not rstInvoiceDetails.EOF
		If bolInvoicesEditGranted Then
			strEditLink = "<a href=""InvoiceDetailEdit.asp?lngInvoiceDetailID=" & (rstInvoiceDetails.Fields.Item("InvoiceDetailID").Value) & """>"
			strEndEditLink = "</a>&nbsp;"
		Else
			strEditLink = ""
			strEndEditLink = "&nbsp;"
		End If
		
		If IsNull(rstInvoiceDetails.Fields.Item("Time").Value) Then
			dblTime = 0
		Else
			dblTime = rstInvoiceDetails.Fields.Item("Time").Value
			dblTimeTotal = dblTimeTotal + CDbl(rstInvoiceDetails.Fields.Item("Time").Value)
		End If
		
		If IsNull(rstInvoiceDetails.Fields.Item("Amount").Value) Then
			curLineTotal = 0
		Else
			curLineTotal = rstInvoiceDetails.Fields.Item("Amount").Value
			curInvoiceTotal = curInvoiceTotal + curLineTotal
		End If
		
		
		
		
%>                        
                        <tr>
                          <td>&nbsp;</td>
                          <td><%=strEditLink & (rstInvoiceDetails.Fields.Item("Description").Value) & strEndEditLink%></td>
                          <td align="center"><%=strEditLink & FormatNumber(dblTime, 1, -1) & strEndEditLink%></td>
                          <td align="right"><%=FormatCurrency(curLineTotal)%></td>
                          <td align="right">&nbsp;</td>
                        </tr>                        
<%
		rstInvoiceDetails.MoveNext
	Loop
		
%>                        
                        <tr>
                          <td colspan="5"><hr /></td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                          <td align="right"><strong>Totals:</strong></td>
                          <td align="center"><%=FormatNumber(dblTimeTotal, 1)%></td>
                          <td align="right"><%=FormatCurrency(curInvoiceTotal)%></td>
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
rstInvoices.Close()
Set rstInvoices = Nothing
%>
<%
rstInvoiceDetails.Close()
Set rstInvoiceDetails = Nothing
%>
<%
rstWorkHistorys.Close()
Set rstWorkHistorys = Nothing
%>
