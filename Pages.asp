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
    MM_editCmd.CommandText = "INSERT INTO dbo.Pages (PageTitle, PageAddress, PageGroupID, HelpContextID, Active, NavigationPage, MenuSortOrder, DividerBefore) VALUES (?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxPageTitle")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxPageAddress")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("cbxPageGroupID"), Request.Form("cbxPageGroupID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("tbxHelpContextID"), Request.Form("tbxHelpContextID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("chkActive"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("chkNavigationPage"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("tbxMenuSortOrder"), Request.Form("tbxMenuSortOrder"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("chkDividerBefore"), 1, 0)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
If (CStr(Request("MM_insert")) = "frmAdd") Then
	lngAccessTypeID = 3
End If
%>
<%
Dim rstPages
Dim rstPages_cmd
Dim rstPages_numRows

Set rstPages_cmd = Server.CreateObject ("ADODB.Command")
rstPages_cmd.ActiveConnection = MM_OBA_STRING
rstPages_cmd.CommandText = "SELECT Pages.PageID, Pages.PageGroupID, Pages.PageTitle, Pages.NavigationPage, Pages.PageAddress, Pages.HelpContextID, Pages.Active, Pages.MenuSortOrder, Pages.DividerBefore, PageGroups.GroupName FROM Pages INNER JOIN PageGroups ON Pages.PageGroupID = PageGroups.PageGroupID ORDER BY PageGroups.GroupName, Pages.NavigationPage DESC, Pages.MenuSortOrder, Pages.PageTitle" 
rstPages_cmd.Prepared = true

Set rstPages = rstPages_cmd.Execute
rstPages_numRows = 0
%>
<%
Dim rstPageGroups
Dim rstPageGroups_cmd
Dim rstPageGroups_numRows

Set rstPageGroups_cmd = Server.CreateObject ("ADODB.Command")
rstPageGroups_cmd.ActiveConnection = MM_OBA_STRING
rstPageGroups_cmd.CommandText = "SELECT * FROM dbo.PageGroups ORDER BY GroupName" 
rstPageGroups_cmd.Prepared = true

Set rstPageGroups = rstPageGroups_cmd.Execute
rstPageGroups_numRows = 0
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

<h1><%=strPageTitle & " " & strSubTitle%></h1>

			<table border="0" cellspacing="0" cellpadding="0" class="fixed">
              <tr>
                <th>&nbsp;</th>
                <th align="left"><h4>Page Title</h4></th>
                <th align="left"><h4>Page URL</h4></th>
                <th align="center" nowrap="nowrap"><h4>Page Group</h4></th>
                <th align="center" nowrap="nowrap"><h4>Help ID</h4></th>
                <th nowrap="nowrap"><h4>Active</h4></th>
                <th nowrap="nowrap"><h4>Navigation</h4></th>
                <th nowrap="nowrap"><h4>Sort Order</h4></th>
                <th nowrap="nowrap"><h4>Divider Before</h4></th>
              </tr>
<%
Do While Not rstPages.EOF
	If bolDeveloperEditGranted Then
		strEdit = "<a href=""PageEdit.asp?lngPageID=" & (rstPages.Fields.Item("PageID").Value) & """>"
		strEditEnd = "</a>"
	Else
		strEdit = ""
		strEditEnd = ""
	
	End If
%>              
              <tr class="tr_hover">
                <td><a href="PageInformation.asp?lngPageID=<%=(rstPages.Fields.Item("PageID").Value)%>" class="row_info"></a></td>
                <td nowrap="nowrap"><%=strEdit & (rstPages.Fields.Item("PageTitle").Value) & strEditEnd%></td>
                <td nowrap="nowrap"><%=strEdit & (rstPages.Fields.Item("PageAddress").Value) & strEditEnd%></td>
                <td align="center"><%=strEdit & (rstPages.Fields.Item("GroupName").Value) & strEditEnd%></td>
                <td align="center"><%=strEdit & (rstPages.Fields.Item("HelpContextID").Value) & strEditEnd%></td>
                <td align="center"><%=strEdit & (rstPages.Fields.Item("Active").Value) & strEditEnd%></td>
                <td align="center"><%=strEdit & (rstPages.Fields.Item("NavigationPage").Value) & strEditEnd%></td>
                <td align="center"><%=strEdit & (rstPages.Fields.Item("MenuSortOrder").Value) & strEditEnd%></td>
                <td align="center"><%=strEdit & (rstPages.Fields.Item("DividerBefore").Value) & strEditEnd%></td>
              </tr>
<%
	rstPages.MoveNext
Loop
%>            
              <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
              <tr>
                <td>&nbsp;</td>
                <td nowrap="nowrap"><input type="text" name="tbxPageTitle" tabindex="1" id="tbxPageTitle" /></td>
                <td nowrap="nowrap"><input name="tbxPageAddress" type="text" id="tbxPageAddress" tabindex="1" size="35" /></td>
                <td align="center"><select name="cbxPageGroupID" id="cbxPageGroupID">
                  <%
While (NOT rstPageGroups.EOF)
%>
                  <option value="<%=(rstPageGroups.Fields.Item("PageGroupID").Value)%>"><%=(rstPageGroups.Fields.Item("GroupName").Value)%></option>
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
                <td align="center"><input name="tbxHelpContextID" type="text" id="tbxHelpContextID" value="10" size="5" maxlength="4" /></td>
                <td align="center"><input type="checkbox" name="chkActive" id="chkActive" /></td>
                <td align="center"><input type="checkbox" name="chkNavigationPage" tabindex="1" id="chkNavigationPage" /></td>
                <td align="center"><input name="tbxMenuSortOrder" type="text" id="tbxMenuSortOrder" value="10" size="5" maxlength="4" /></td>
                <td align="center"><input type="checkbox" name="chkDividerBefore" id="chkDividerBefore" /></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td nowrap="nowrap">&nbsp;</td>
                <td nowrap="nowrap">&nbsp;</td>
                <td>&nbsp;</td>
                <td colspan="5" align="right"><input type="submit" name="btnAdd" tabindex="1" id="btnAdd" value="Add Page" /></td>
              </tr>
              <input type="hidden" name="MM_insert" value="frmAdd" />
              </form>
              <tr>
                <td>&nbsp;</td>
                <td nowrap="nowrap">&nbsp;</td>
                <td nowrap="nowrap">&nbsp;</td>
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
rstPages.Close()
Set rstPages = Nothing
%>
<%
rstPageGroups.Close()
Set rstPageGroups = Nothing
%>
