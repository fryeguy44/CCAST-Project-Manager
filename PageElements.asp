<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%
Dim strSubTitle 
Dim lngPageID

lngPageID = Request.QueryString("lngPageID")

If lngPageID = "" Then
	strSubTitle = ""
Else	
	strSubTitle = " (filtered)"
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
If (CStr(Request("MM_insert")) = "frmAdd") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.PageElements (PageID, ElementID) VALUES (?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("cbxPageID"), Request.Form("cbxPageID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxElementID"), Request.Form("cbxElementID"), null)) ' adDouble
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
If (lngPageID <> "") Then 
%>
<%
Dim rstPageElements__lngPageID
	rstPageElements__lngPageID = "1"
	If (lngPageID <> "") Then 
	  rstPageElements__lngPageID = lngPageID
	End If
%>
<%
Dim rstPageElements
Dim rstPageElements_cmd
Dim rstPageElements_numRows

Set rstPageElements_cmd = Server.CreateObject ("ADODB.Command")
rstPageElements_cmd.ActiveConnection = MM_OBA_STRING
rstPageElements_cmd.CommandText = "SELECT PageElements.PageElementID, PageElements.PageID, PageElements.ElementID, Pages.PageTitle, Elements.ElementName FROM PageElements INNER JOIN Pages ON PageElements.PageID = Pages.PageID INNER JOIN Elements ON PageElements.ElementID = Elements.ElementID WHERE PageElements.PageID = ? ORDER BY Elements.ElementName" 
rstPageElements_cmd.Prepared = true
rstPageElements_cmd.Parameters.Append rstPageElements_cmd.CreateParameter("param1", 5, 1, -1, rstPageElements__lngPageID) ' adDouble

Set rstPageElements = rstPageElements_cmd.Execute
rstPageElements_numRows = 0
%>
<%
Else
	Set rstPageElements_cmd = Server.CreateObject ("ADODB.Command")
	rstPageElements_cmd.ActiveConnection = MM_OBA_STRING
	rstPageElements_cmd.CommandText = "SELECT PageElements.PageElementID, PageElements.PageID, PageElements.ElementID, Pages.PageTitle, Elements.ElementName FROM ((PageElements INNER JOIN Pages ON PageElements.PageID = Pages.PageID) INNER JOIN Elements ON PageElements.ElementID = Elements.ElementID) ORDER BY Pages.PageTitle, Elements.ElementName" 
	rstPageElements_cmd.Prepared = true
	
	Set rstPageElements = rstPageElements_cmd.Execute
	rstPageElements_numRows = 0
End If
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
			<h1><a href="PageElements.asp"><%=strPageTitle & " " & strSubTitle%></a></h1>
			<table border="0" cellpadding="0" cellspacing="0">
<%
Dim strRecPageTitle
strRecPageTitle = ""
Do While Not rstPageElements.EOF
	If strRecPageTitle <> (rstPageElements.Fields.Item("PageTitle").Value) Then
		strRecPageTitle = (rstPageElements.Fields.Item("PageTitle").Value)
%>
              <tr>
                <td width="10">&nbsp;</td>
                <td colspan="4"><strong>Page Title</strong></td>
              </tr>
              <tr>
                <td><a href="PageInformation.asp?lngPageID=<%=(rstPageElements.Fields.Item("PageID").Value)%>" class="row_info">&nbsp;</a></td>
                <td colspan="4"><a href="PageEdit.asp?lngPageID=<%=(rstPageElements.Fields.Item("PageID").Value)%>"><%=(rstPageElements.Fields.Item("PageTitle").Value)%></a></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td width="10">&nbsp;</td>
                <td colspan="2"><strong>Element</strong></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td colspan="3"><hr /></td>
              </tr>
<%
	End If
%>              
              <tr>
                <td>&nbsp;</td>
                <td><a href="ElementInformation.asp?lngElementID=<%=(rstPageElements.Fields.Item("ElementID").Value)%>" class="row_info">&nbsp;</a></td>
                <td colspan="2"><a href="PageElementEdit.asp?lngPageElementID=<%=(rstPageElements.Fields.Item("PageElementID").Value)%>"><%=(rstPageElements.Fields.Item("ElementName").Value)%></a></td>
                <td>&nbsp;</td>
              </tr>
<%
	rstPageElements.MoveNext
Loop
%>              
              <tr>
                <td>&nbsp;</td>
                <td colspan="2">&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td colspan="5"><h2>Add Page Element</h2></td>
              </tr>
              <tr>
                <td colspan="5"><table border="0" cellspacing="0" cellpadding="0">
                  <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
                  <tr>
                    <td>&nbsp;</td>
                    <td><strong>Page</strong></td>
                    <td>
                      <select name="cbxPageID" id="cbxPageID">
                        <%
While (NOT rstPages.EOF)
%><option value="<%=(rstPages.Fields.Item("PageID").Value)%>" <%If (Not isNull(lngPageID)) Then If (CStr(rstPages.Fields.Item("PageID").Value) = CStr(lngPageID)) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstPages.Fields.Item("PageTitle").Value)%></option>
                        <%
  rstPages.MoveNext()
Wend
If (rstPages.CursorType > 0) Then
  rstPages.MoveFirst
Else
  rstPages.Requery
End If
%>
                      </select>                    </td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td><strong>Element&nbsp;</strong></td>
                    <td><select name="cbxElementID" id="cbxElementID">
                      <%
While (NOT rstElements.EOF)
%>
                      <option value="<%=(rstElements.Fields.Item("ElementID").Value)%>"><%=(rstElements.Fields.Item("ElementName").Value)%></option>
                      <%
  rstElements.MoveNext()
Wend
If (rstElements.CursorType > 0) Then
  rstElements.MoveFirst
Else
  rstElements.Requery
End If
%>
                    </select>                    </td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td><input type="submit" name="btnAdd" id="btnAdd" value="Add" /></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
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
                  </tr>
                  <input type="hidden" name="MM_insert" value="frmAdd" />
</form>
                </table></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td colspan="2">&nbsp;</td>
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
