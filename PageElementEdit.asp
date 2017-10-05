<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%
Dim lngPageElementID
Dim strReturnPath

lngPageElementID = Request.Querystring("lngPageElementID")
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
    MM_editCmd.CommandText = "UPDATE dbo.PageElements SET PageID = ?, ElementID = ? WHERE PageElementID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("cbxPageID"), Request.Form("cbxPageID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxElementID"), Request.Form("cbxElementID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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

<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "DELETE FROM dbo.PageElements WHERE PageElementID = ?"
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
If (CStr(Request("MM_update")) = "frmEdit") OR (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then
	Response.Redirect(Request.Form("htbxReferer"))
End If
%>

<%
Dim rstPageElements__lngPageElementID
rstPageElements__lngPageElementID = "1"
If (lngPageElementID <> "") Then 
  rstPageElements__lngPageElementID = lngPageElementID
End If
%>
<%
Dim rstPageElements
Dim rstPageElements_cmd
Dim rstPageElements_numRows

Set rstPageElements_cmd = Server.CreateObject ("ADODB.Command")
rstPageElements_cmd.ActiveConnection = MM_OBA_STRING
rstPageElements_cmd.CommandText = "SELECT PageElements.* FROM PageElements WHERE PageElementID = ?" 
rstPageElements_cmd.Prepared = true
rstPageElements_cmd.Parameters.Append rstPageElements_cmd.CreateParameter("param1", 5, 1, -1, rstPageElements__lngPageElementID) ' adDouble

Set rstPageElements = rstPageElements_cmd.Execute
rstPageElements_numRows = 0
%>
<%
Dim rstPages
Dim rstPages_cmd
Dim rstPages_numRows

Set rstPages_cmd = Server.CreateObject ("ADODB.Command")
rstPages_cmd.ActiveConnection = MM_OBA_STRING
rstPages_cmd.CommandText = "SELECT * FROM dbo.Pages ORDER BY PageTitle" 
rstPages_cmd.Prepared = true

Set rstPages = rstPages_cmd.Execute
rstPages_numRows = 0
%>
<%
Dim rstElements
Dim rstElements_cmd
Dim rstElements_numRows

Set rstElements_cmd = Server.CreateObject ("ADODB.Command")
rstElements_cmd.ActiveConnection = MM_OBA_STRING
rstElements_cmd.CommandText = "SELECT * FROM dbo.Elements ORDER BY ElementName" 
rstElements_cmd.Prepared = true

Set rstElements = rstElements_cmd.Execute
rstElements_numRows = 0
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
              <form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
              <tr>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	</tr>
              <tr>
                <td width="10">&nbsp;</td>
                <td width="130"><strong>Page</strong></td>
                <td>
                  <select name="cbxPageID" id="cbxPageID">
                    <%
While (NOT rstPages.EOF)
%>
                    <option value="<%=(rstPages.Fields.Item("PageID").Value)%>" <%If (Not isNull((rstPageElements.Fields.Item("PageID").Value))) Then If (CStr(rstPages.Fields.Item("PageID").Value) = CStr((rstPageElements.Fields.Item("PageID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstPages.Fields.Item("PageTitle").Value)%></option>
                    <%
  rstPages.MoveNext()
Wend
If (rstPages.CursorType > 0) Then
  rstPages.MoveFirst
Else
  rstPages.Requery
End If
%>
                  </select>                </td>
                </tr>
              <tr>
                <td width="10">&nbsp;</td>
                <td width="130"><strong>Element</strong></td>
                <td><select name="cbxElementID" id="cbxElementID">
                  <%
While (NOT rstElements.EOF)
%>
                  <option value="<%=(rstElements.Fields.Item("ElementID").Value)%>" <%If (Not isNull((rstPageElements.Fields.Item("ElementID").Value))) Then If (CStr(rstElements.Fields.Item("ElementID").Value) = CStr((rstPageElements.Fields.Item("ElementID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstElements.Fields.Item("ElementName").Value)%></option>
                  <%
  rstElements.MoveNext()
Wend
If (rstElements.CursorType > 0) Then
  rstElements.MoveFirst
Else
  rstElements.Requery
End If
%>
                </select></td>
                </tr>
              <tr>
                <td width="10">&nbsp;</td>
                <td width="130"><input name="htbxReferer" type="hidden" id="htbxReferer" value="<%=strReturnPath%>" /></td>
                <td><input type="submit" name="btnEdit" id="btnEdit" value="Update" /></td>
                </tr>
              <input type="hidden" name="MM_update" value="frmEdit" />
              <input type="hidden" name="MM_recordId" value="<%= rstPageElements.Fields.Item("PageElementID").Value %>" />
              </form>
              <tr>
                <td width="10">&nbsp;</td>
                <td width="130"><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
                  <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
                  <input type="hidden" name="MM_delete" value="frmDelete" />
                  <input type="hidden" name="MM_recordId" value="<%= rstPageElements.Fields.Item("PageElementID").Value %>" />
                  <input name="htbxReferer" type="hidden" id="htbxReferer" value="<%=strReturnPath%>" />
                </form>                </td>
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
rstPageElements.Close()
Set rstPageElements = Nothing
%>
<%
rstPages.Close()
Set rstPages = Nothing
%>
<%
rstElements.Close()
Set rstElements = Nothing
%>
