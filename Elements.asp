<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
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
    MM_editCmd.CommandText = "INSERT INTO dbo.Elements (ElementName) VALUES (?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxElementName")) ' adVarWChar
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
Dim rstElements
Dim rstElements_cmd
Dim rstElements_numRows

Set rstElements_cmd = Server.CreateObject ("ADODB.Command")
rstElements_cmd.ActiveConnection = MM_OBA_STRING
rstElements_cmd.CommandText = "SELECT ElementName, ElementID FROM Elements ORDER BY Elements.ElementName" 
rstElements_cmd.Prepared = true

Set rstElements = rstElements_cmd.Execute
rstElements_numRows = 0
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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

    <table border="0" cellspacing="0" cellpadding="0" class="box">
        <tr>
            <th colspan="2"><h2><strong>Elements</strong></h2></th>
        </tr>
<%
If bolDeveloperAddGranted Then
%>
        <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
        <tr>
            <td colspan="2"><strong>New Element Name: </strong>
                <input type="text" name="tbxElementName" id="tbxElementName" />
                <input type="submit" name="btnAdd" id="btnAdd" value="Add Element" />
                <input type="hidden" name="MM_insert" value="frmAdd" />
            </td>
        </tr>
        </form>
        <tr>
            <td colspan="2"><hr /></td>
        </tr>
<%
End If
If bolDeveloperViewGranted Then
	Do While Not rstElements.EOF
		If bolDeveloperEditGranted Then
			strEditLink = "<a href=""ElementEdit.asp?lngElementID=" & (rstElements.Fields.Item("ElementID").Value) & """>"
			strEndEditLink = "</a>&nbsp;"
		Else
			strEditLink = ""
			strEndEditLink = "&nbsp;"
		End If
%>
        <tr class="tr_hover">
            <td><a href="ElementInformation.asp?lngElementID=<%=(rstElements.Fields.Item("ElementID").Value)%>" class="row_info"></a></td>
            <td><%=strEditLink & (rstElements.Fields.Item("ElementName").Value) & strEndEditLink%></td>
        </tr>
                  <%
		rstElements.MoveNext
	Loop
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
rstElements.Close()
Set rstElements = Nothing
%>