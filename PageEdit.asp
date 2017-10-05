<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%
Dim lngPageID
Dim strReturnPath

lngPageID = Request.Querystring("lngPageID")
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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "DELETE FROM dbo.Pages WHERE PageID = ?"
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
<%
If (CStr(Request("MM_update")) = "frmEdit") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "UPDATE dbo.Pages SET PageTitle = ?, PageAddress = ?, PageGroupID = ?, HelpContextID = ?, Active = ?, NavigationPage = ?, MenuSortOrder = ?, DividerBefore = ? WHERE PageID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxPageTitle")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxPageAddress")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("cabPageGroupID"), Request.Form("cabPageGroupID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxHelpContextID"), Request.Form("tbxHelpContextID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("chkActive"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("chkNavigationPage"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("tbxMenuSortOrder"), Request.Form("tbxMenuSortOrder"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("chkDividerBefore"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_update")) = "frmEdit") Then
	lngAccessTypeID = 2
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
If (CStr(Request("MM_update")) = "frmEdit") OR (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then
    Response.Redirect(Request.Form("htbxReturnPath"))
End If
%>

<%
Dim rstPages__lngPageID
rstPages__lngPageID = "1"
If (lngPageID <> "") Then 
  rstPages__lngPageID = lngPageID
End If
%>
<%
Dim rstPages
Dim rstPages_cmd
Dim rstPages_numRows

Set rstPages_cmd = Server.CreateObject ("ADODB.Command")
rstPages_cmd.ActiveConnection = MM_OBA_STRING
rstPages_cmd.CommandText = "SELECT * FROM dbo.Pages WHERE PageID = ?" 
rstPages_cmd.Prepared = true
rstPages_cmd.Parameters.Append rstPages_cmd.CreateParameter("param1", 5, 1, -1, rstPages__lngPageID) ' adDouble

Set rstPages = rstPages_cmd.Execute
rstPages_numRows = 0
%>
<%
Dim rstPageGroups
Dim rstPageGroups_cmd
Dim rstPageGroups_numRows

Set rstPageGroups_cmd = Server.CreateObject ("ADODB.Command")
rstPageGroups_cmd.ActiveConnection = MM_OBA_STRING
rstPageGroups_cmd.CommandText = "SELECT * FROM dbo.PageGroups ORDER BY TabOrder" 
rstPageGroups_cmd.Prepared = true

Set rstPageGroups = rstPageGroups_cmd.Execute
rstPageGroups_numRows = 0
%>
<script src="SpryAssets/SpryValidationTextField.js" type="text/javascript"></script>
<link href="SpryAssets/SpryValidationTextField.css" rel="stylesheet" type="text/css" />
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

			<table border="0" cellspacing="0" cellpadding="0" class="fixed">
<%
If bolDeveloperEditGranted Then
%>
              <form ACTION="<%=MM_editAction%>" METHOD="POST" id="frmEdit" name="frmEdit">
              <tr>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	</tr>
              <tr>
                <td width="10">&nbsp;</td>
                <td width="130" align="right"><strong>Page Title</strong></td>
                <td><span id="sprytextfield1">
                <input name="tbxPageTitle" type="text" tabindex="1" id="tbxPageTitle" value="<%=(rstPages.Fields.Item("PageTitle").Value)%>" />
                <span class="textfieldRequiredMsg">A value is required.</span><span class="textfieldMaxCharsMsg">Exceeded maximum number of characters.</span></span></td>
                <td>Display Name of the Page</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td width="130" align="right"><strong>Page URL</strong></td>
                <td><span id="sprytextfield2">
                <input name="tbxPageAddress" type="text" tabindex="1" id="tbxPageAddress" value="<%=(rstPages.Fields.Item("PageAddress").Value)%>" />
                <span class="textfieldMaxCharsMsg">Exceeded maximum number of characters.</span></span></td>
                <td>Page Address of the Page as typed into the browser address bar</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td align="right"><strong>Page Group</strong></td>
                <td><select name="cabPageGroupID" id="cabPageGroupID">
                  <option value="0" <%If (Not isNull((rstPages.Fields.Item("PageGroupID").Value))) Then If ("0" = CStr((rstPages.Fields.Item("PageGroupID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>Not Listed in Navigation</option>
                  <%
While (NOT rstPageGroups.EOF)
%><option value="<%=(rstPageGroups.Fields.Item("PageGroupID").Value)%>" <%If (Not isNull((rstPages.Fields.Item("PageGroupID").Value))) Then If (CStr(rstPageGroups.Fields.Item("PageGroupID").Value) = CStr((rstPages.Fields.Item("PageGroupID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstPageGroups.Fields.Item("GroupName").Value)%></option>
                  <%
  rstPageGroups.MoveNext()
Wend
If (rstPageGroups.CursorType > 0) Then
  rstPageGroups.MoveFirst
Else
  rstPageGroups.Requery
End If
%>
                </select></td>
                <td>The Navigation Tab the page is listed within.</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td align="right"><strong>Help Context ID</strong></td>
                <td><input name="tbxHelpContextID" type="text" id="tbxHelpContextID" value="<%=(rstPages.Fields.Item("HelpContextID").Value)%>" size="5" maxlength="4" /></td>
                <td>Help System Context ID</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td align="right"><strong>Active</strong></td>
                <td><input <%If (CStr((rstPages.Fields.Item("Active").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" tabindex="2" name="chkActive" id="chkActive" /></td>
                <td>Is the page Active in the Application?</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td width="130" align="right"><strong>Navigation</strong></td>
                <td><input <%If (CStr((rstPages.Fields.Item("NavigationPage").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" tabindex="1" name="chkNavigationPage" id="chkNavigationPage" /></td>
                <td>Display Name of the Page</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td align="right"><strong>Sort Order</strong></td>
                <td><input name="tbxMenuSortOrder" type="text" id="tbxMenuSortOrder" value="<%=(rstPages.Fields.Item("MenuSortOrder").Value)%>" size="5" maxlength="4" /></td>
                <td>Display Order of the menu items</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td align="right"><strong>Divider Before</strong></td>
                <td><input <%If (CStr((rstPages.Fields.Item("DividerBefore").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="chkDividerBefore" id="chkDividerBefore" /></td>
                <td>Places a horizontal divider before the item in the menu</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td width="130"><input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" /></td>
                <td><input type="submit" name="btnEdit" tabindex="1" id="btnEdit" value="Update" /></td>
                <td>&nbsp;</td>
              </tr>
              <input type="hidden" name="MM_update" value="frmEdit" />
              <input type="hidden" name="MM_recordId" value="<%= rstPages.Fields.Item("PageID").Value %>" />
              </form>
<%
End If
If bolDeveloperDeleteGranted Then
%>                           
              <tr>
                <td>&nbsp;</td>
                <td width="130"><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
                  <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
                  <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
                  <input type="hidden" name="MM_delete" value="frmDelete" />
                  <input type="hidden" name="MM_recordId" value="<%= rstPages.Fields.Item("PageID").Value %>" />
                </form>                </td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
<%
End If
%>                           
              <tr>
                <td>&nbsp;</td>
                <td width="130">&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
	        <script type="text/javascript">
<!--
var sprytextfield1 = new Spry.Widget.ValidationTextField("sprytextfield1", "none", {maxChars:49});
var sprytextfield2 = new Spry.Widget.ValidationTextField("sprytextfield2", "none", {maxChars:49, isRequired:false});
//-->
</script>
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
rstPages.Close()
Set rstPages = Nothing
%>
<%
rstPageGroups.Close()
Set rstPageGroups = Nothing
%>
