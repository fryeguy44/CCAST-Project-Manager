<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngVendorID
Dim strReturnPath

lngVendorID = Request.QueryString("lngVendorID")
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
    MM_editCmd.CommandText = "UPDATE dbo.Vendors SET VendorName = ?, Country = ?, SkypeID = ?, TeamViewerID = ?, Phone = ?, Email = ?, Rate = ? WHERE VendorID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxVendorName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxCountry")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("tbxSkypeID")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxTeamviewerID"), Request.Form("tbxTeamviewerID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 20, Request.Form("tbxPhone")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 50, Request.Form("tbxEmail")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("tbxVendorRate"), Request.Form("tbxVendorRate"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
    MM_editCmd.CommandText = "DELETE FROM dbo.Vendors WHERE VendorID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
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
rstVendors_cmd.CommandText = "SELECT TOP (1) Vendors.VendorID, Vendors.VendorName, Vendors.Country, Vendors.SkypeID, Vendors.TeamViewerID, Vendors.Phone, Vendors.Email, Vendors.PayPal, Vendors.Rate, ProjectDetails.ProjectDetailID,  VendorInvoices.VendorInvoiceID FROM Vendors LEFT OUTER JOIN VendorInvoices ON Vendors.VendorID = VendorInvoices.VendorID LEFT OUTER JOIN ProjectDetails ON Vendors.VendorID = ProjectDetails.VendorID WHERE Vendors.VendorID = ?" 
rstVendors_cmd.Prepared = true
rstVendors_cmd.Parameters.Append rstVendors_cmd.CreateParameter("param1", 5, 1, -1, rstVendors__lngVendorID) ' adDouble

Set rstVendors = rstVendors_cmd.Execute
rstVendors_numRows = 0
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
	Response.Redirect("Vendors.asp")
End If

%>

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
If bolVendorsEditGranted Then
	If rstVendors.EOF Then
%>  
        <tr>
          <th colspan="4">&nbsp;</th>
        </tr>
        <tr>
            <td colspan="4"><a href="Vendors.asp">The Vendor you are attempting to edit has been deleted. Click here to return to the Vendor List page</a></td>
        </tr>
<%
	Else
%>     
    	<form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
        <tr>
            <td width="10">&nbsp;</td>
            <td align="right"><strong>Vendor Name</strong></td>
          <td><input name="tbxVendorName" type="text" id="tbxVendorName" value="<%=(rstVendors.Fields.Item("VendorName").Value)%>" /></td>
		<td>&nbsp;</td>
		</tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Country</strong></td>
          <td><input name="tbxCountry" type="text" id="tbxCountry" value="<%=(rstVendors.Fields.Item("Country").Value)%>" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Skype</strong></td>
          <td><input name="tbxSkypeID" type="text" id="tbxSkypeID" value="<%=(rstVendors.Fields.Item("SkypeID").Value)%>" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Teamviewer</strong></td>
          <td><input name="tbxTeamviewerID" type="text" id="tbxTeamviewerID" value="<%=(rstVendors.Fields.Item("TeamViewerID").Value)%>" size="10" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Phone</strong></td>
          <td><input name="tbxPhone" type="text" id="tbxPhone" value="<%=(rstVendors.Fields.Item("Phone").Value)%>" size="15" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Email</strong></td>
          <td><input name="tbxEmail" type="text" id="tbxEmail" value="<%=(rstVendors.Fields.Item("Email").Value)%>" size="35" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Current Rate</strong></td>
          <td><input name="tbxVendorRate" type="text" id="tbxVendorRate" value="<%=(rstVendors.Fields.Item("Rate").Value)%>" size="8" /></td>
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
        <input type="hidden" name="MM_recordId" value="<%= rstVendors.Fields.Item("VendorID").Value %>" />
        </form>
<%
		If bolVendorsDeleteGranted AND IsNull(rstVendors.Fields.Item("ProjectDetailID").Value) AND IsNull(rstVendors.Fields.Item("VendorInvoiceID").Value) Then
%>                
      <tr>
        <td width="10">&nbsp;</td>
            <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
              <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
              <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
              <input type="hidden" name="MM_delete" value="frmDelete" />
              <input type="hidden" name="MM_recordId" value="<%= rstVendors.Fields.Item("VendorID").Value %>" />
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
            <td colspan="4">Certain &quot;Vendors&quot; permissions are required to perform this task.</td>
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
rstVendors.Close()
Set rstVendors = Nothing
%>
