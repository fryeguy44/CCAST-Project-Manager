<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%
Dim lngPageGroupID
Dim strReturnPath

lngPageGroupID = Request.Querystring("lngPageGroupID")
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
If (CStr(Request("MM_update")) = "frmUpdate") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "UPDATE dbo.PageGroups SET GroupName = ?, TabOrder = ? WHERE PageGroupID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxGroupName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("tbxTabOrder"), Request.Form("tbxTabOrder"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_update")) = "frmUpdate") Then
	lngAccessTypeID = 2
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "DELETE FROM dbo.PageGroups WHERE PageGroupID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If

End If
%>
<%
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
If (CStr(Request("MM_update")) = "frmUpdate") OR (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then
    Response.Redirect(Request.Form("htbxReturnPath"))
End If
%>
<%
Dim rstPageGroups__lngPageGroupId
rstPageGroups__lngPageGroupId = "3"
If (lngPageGroupId <> "") Then 
  rstPageGroups__lngPageGroupId = lngPageGroupId
End If
%>
<%
Dim rstPageGroups
Dim rstPageGroups_cmd
Dim rstPageGroups_numRows

Set rstPageGroups_cmd = Server.CreateObject ("ADODB.Command")
rstPageGroups_cmd.ActiveConnection = MM_OBA_STRING
rstPageGroups_cmd.CommandText = "SELECT * FROM dbo.PageGroups WHERE PageGroupID = ?" 
rstPageGroups_cmd.Prepared = true
rstPageGroups_cmd.Parameters.Append rstPageGroups_cmd.CreateParameter("param1", 5, 1, -1, rstPageGroups__lngPageGroupId) ' adDouble

Set rstPageGroups = rstPageGroups_cmd.Execute
rstPageGroups_numRows = 0
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
              <form id="frmUpdate" name="frmUpdate" method="POST" action="<%=MM_editAction%>">
              <tr>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	</tr>
              <tr>
                <td width="10">&nbsp;</td>
                <td><strong>Group Name</strong></td>
                <td>
                    <input name="tbxGroupName" type="text" id="tbxGroupName" value="<%=(rstPageGroups.Fields.Item("GroupName").Value)%>" />                               </td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><strong>Tab Order</strong></td>
                <td><input name="tbxTabOrder" type="text" id="tbxTabOrder" value="<%=(rstPageGroups.Fields.Item("TabOrder").Value)%>" size="5" /></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />&nbsp;</td>
                <td><input type="submit" name="btnAdd" id="btnAdd" value="Update" /></td>
                <td>&nbsp;</td>
              </tr>
              <input type="hidden" name="MM_update" value="frmUpdate" />
              <input type="hidden" name="MM_recordId" value="<%= rstPageGroups.Fields.Item("PageGroupID").Value %>" />
              </form>  
              <tr>
                <td>&nbsp;</td>
                <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
                  <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
                  <input type="hidden" name="MM_delete" value="frmDelete" />
                  <input type="hidden" name="MM_recordId" value="<%= rstPageGroups.Fields.Item("PageGroupID").Value %>" />
                  <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />                
                </form>
                </td>
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
rstPageGroups.Close()
Set rstPageGroups = Nothing
%>
