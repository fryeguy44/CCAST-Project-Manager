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
    MM_editCmd.CommandText = "INSERT INTO dbo.Vendors (VendorName, Country, SkypeID, TeamViewerID, Phone, Email) VALUES (?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxVendorName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxCountry")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("tbxSkypeID")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxTeamviewerID"), Request.Form("tbxTeamviewerID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 20, Request.Form("tbxPhone")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 50, Request.Form("tbxEmail")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim rstVendors
Dim rstVendors_cmd
Dim rstVendors_numRows

Set rstVendors_cmd = Server.CreateObject ("ADODB.Command")
rstVendors_cmd.ActiveConnection = MM_OBA_STRING
rstVendors_cmd.CommandText = "SELECT * FROM Vendors" 
rstVendors_cmd.Prepared = true

Set rstVendors = rstVendors_cmd.Execute
rstVendors_numRows = 0
%>
<%
Dim rstWorkHistory
Dim rstWorkHistory_cmd
Dim rstWorkHistory_numRows

Set rstWorkHistory_cmd = Server.CreateObject ("ADODB.Command")
rstWorkHistory_cmd.ActiveConnection = MM_OBA_STRING
rstWorkHistory_cmd.CommandText = "SELECT WorkHistorys.WorkHistoryID, WorkHistorys.ProjectDetailID, Vendors.VendorName, WorkHistorys.WorkDate, WorkHistorys.Hours, LEFT(WorkHistorys.WorkDescription, 150) AS WorkDescription, VendorInvoiceDetails.VendorInvoiceDetailID,  ProjectDetails.DetailDescription, Vendors.Rate, Clients.ClientName FROM WorkHistorys INNER JOIN Vendors ON WorkHistorys.VendorID = Vendors.VendorID INNER JOIN ProjectDetails ON WorkHistorys.ProjectDetailID = ProjectDetails.ProjectDetailID INNER JOIN Projects ON ProjectDetails.ProjectID = Projects.ProjectID INNER JOIN Clients ON Projects.ClientID = Clients.ClientID LEFT OUTER JOIN VendorInvoiceDetails ON WorkHistorys.WorkHistoryID = VendorInvoiceDetails.WorkHistoryID WHERE (VendorInvoiceDetails.VendorInvoiceDetailID IS NULL) ORDER BY Vendors.VendorName, Clients.ClientName, WorkHistorys.WorkDate" 
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
	<h1>Vendor List</h1>	
	<table border="0" cellspacing="0" cellpadding="0" class="box">
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="left"><h4>Vendor Name</h4></th>
	    <th align="left"><h4>Country</h4></th>
	    <th align="left"><h4>Skype</h4></th>
	    <th align="left"><h4>Teamviewer</h4></th>
	    <th align="left"><h4>Phone</h4></th>
	    <th align="left"><h4>Email</h4></th>
	    <th align="left"><h4>PayPal</h4></th>
	    <th><h4>Vendor ID</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
<%
If bolVendorsViewGranted Then
    If bolVendorsAddGranted Then
%>
      <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
	  <tr>
	    <td>&nbsp;</td>
	    <td><input name="tbxVendorName" type="text" id="tbxVendorName" tabindex="0" size="20" /></td>
	    <td><input type="text" name="tbxCountry" id="tbxCountry" /></td>
	    <td><input type="text" name="tbxSkypeID" id="tbxSkypeID" /></td>
	    <td><input name="tbxTeamviewerID" type="text" id="tbxTeamviewerID" size="10" /></td>
	    <td><input name="tbxPhone" type="text" id="tbxPhone" size="15" /></td>
	    <td><input name="tbxEmail" type="text" id="tbxEmail" size="35" /></td>
	    <td><input type="text" name="tbxPayPal" id="tbxPayPal" /></td>
	    <td><input type="submit" name="btnAdd" id="btnAdd" value="Add Vendor" /></td>
	    <td>&nbsp;</td>
      </tr>
      <input type="hidden" name="MM_insert" value="frmAdd" />
      </form>
	  <tr>
	    <td colspan="10"><hr /></td>
      </tr>
<%
    End If
	Do While Not rstVendors.EOF
		If bolVendorsEditGranted Then
			strEdit = "<a href=""VendorEdit.asp?lngVendorID=" & (rstVendors.Fields.Item("VendorID").Value) & """>"
			strEditEnd = "</a>"
		Else
			strEdit = ""
			strEditEnd = ""
		End If 
%>      
	  <tr class="tr_hover">
	    <td><a href="VendorInformation.asp?lngVendorID=<%=(rstVendors.Fields.Item("VendorID").Value)%>" class="row_info"></a></td>
		<td><%=strEdit & (rstVendors.Fields.Item("VendorName").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstVendors.Fields.Item("Country").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstVendors.Fields.Item("SkypeID").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstVendors.Fields.Item("TeamViewerID").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstVendors.Fields.Item("Phone").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstVendors.Fields.Item("Email").Value) & strEditEnd%></td>
		<td><%=(rstVendors.Fields.Item("PayPal").Value)%></td>
		<td align="center"><%=strEdit & (rstVendors.Fields.Item("VendorID").Value) & strEditEnd%></td>
		<td>&nbsp;</td>
	  </tr>
<%
        rstVendors.MoveNext
    Loop
%>
	</table>
	<table border="0" cellspacing="0" cellpadding="0" class="box">
	  <tr>
	    <th colspan="9"><h2>Unpaid Work History</h2></th>
      </tr>
	  <tr>
	    <th>&nbsp;</th>
		<th align="left"><h4>Vendor</h4></th>
		<th align="left"><h4>Client</h4></th>
		<th align="left"><h4>Date</h4></th>
		<th align="left"><h4>Description</h4></th>
		<th align="left"><h4>Milestone</h4></th>
		<th align="right"><h4>Hours</h4></th>
		<th align="right"><h4>Amount</h4></th>
		<th>&nbsp;</th>
	  </tr>
<%
	Do While Not rstWorkHistory.EOF
		If bolVendorsEditGranted Then
			strEdit = "<a href=""WorkHistoryEdit.asp?lngWorkHistoryID=" & (rstWorkHistory.Fields.Item("WorkHistoryID").Value) & """>"
			strEditEnd = "</a>"
		Else
			strEdit = ""
			strEditEnd = ""
		End If 
%>      
	  <tr>
	    <td>&nbsp;</td>
	    <td nowrap="nowrap"><%=strEdit & (rstWorkHistory.Fields.Item("VendorName").Value) & strEditEnd%></td>
	    <td nowrap="nowrap"><%=strEdit & (rstWorkHistory.Fields.Item("ClientName").Value) & strEditEnd%></td>
	    <td><%=strEdit & (rstWorkHistory.Fields.Item("WorkDate").Value) & strEditEnd%></td>
	    <td nowrap="nowrap"><%=strEdit & (rstWorkHistory.Fields.Item("WorkDescription").Value) & strEditEnd%></td>
	    <td nowrap="nowrap"><%=strEdit & (rstWorkHistory.Fields.Item("DetailDescription").Value) & strEditEnd%></td>
	    <td align="right"><%=strEdit & FormatNumber(rstWorkHistory.Fields.Item("Hours").Value, 1) & strEditEnd%></td>
	    <td align="right"><%=FormatCurrency((rstWorkHistory.Fields.Item("Hours").Value) * (rstWorkHistory.Fields.Item("Rate").Value))%></td>
	    <td>&nbsp;</td>
      </tr>
<%
		rstWorkHistory.MoveNext
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
            <td colspan="10">Viewing this list requires certain &quot;Vendors&quot; permissions</td>
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
rstVendors.Close()
Set rstVendors = Nothing
%>
<%
rstWorkHistory.Close()
Set rstWorkHistory = Nothing
%>
