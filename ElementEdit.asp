<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%

Dim lngElementID
Dim strReturnPath

lngElementID = Request.Querystring("lngElementID")
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
    MM_editCmd.CommandText = "UPDATE dbo.Elements SET ElementName = ? WHERE ElementID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxElementName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
    MM_editCmd.CommandText = "DELETE FROM Elements WHERE ElementID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If

End If
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
If (CStr(Request("MM_update")) = "frmEdit") OR (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then
	Response.Redirect(Request.Form("htbxReturnPath"))
End If
%>

<%
Dim rstElements__lngElementID
rstElements__lngElementID = "1"
If (lngElementID <> "") Then 
  rstElements__lngElementID = lngElementID
End If
%>
<%
Dim rstElements
Dim rstElements_cmd
Dim rstElements_numRows

Set rstElements_cmd = Server.CreateObject ("ADODB.Command")
rstElements_cmd.ActiveConnection = MM_OBA_STRING
rstElements_cmd.CommandText = "SELECT * FROM Elements WHERE ElementID = ?" 
rstElements_cmd.Prepared = true
rstElements_cmd.Parameters.Append rstElements_cmd.CreateParameter("param1", 5, 1, -1, rstElements__lngElementID) ' adDouble

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


<%
If bolDeveloperEditGranted then
%>            
			<table border="0" cellspacing="0" cellpadding="0" class="box">
              <form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
              <tr>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	</tr>
              <tr>
                <td>&nbsp;</td>
                <td><strong>Element Name&nbsp;</strong></td>
                <td>
            <label>
                    <input name="tbxElementName" type="text" id="tbxElementName" value="<%=(rstElements.Fields.Item("ElementName").Value)%>" size="35" />
                  </label>                </td>
                </tr>
              
              <tr>
                <td>&nbsp;</td>
                <td><input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" /></td>
                <td><label>
                  <input type="submit" name="btnEdit" id="btnEdit" value="Update" />
                </label></td>
                </tr>
              <input type="hidden" name="MM_update" value="frmEdit" />
              <input type="hidden" name="MM_recordId" value="<%= rstElements.Fields.Item("ElementID").Value %>" />
              </form>
<%
	If bolDeveloperDeleteGranted Then
%>              
<tr>
                  	<td>&nbsp;</td>
  <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
                      <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
                      <input type="hidden" name="MM_delete" value="frmDelete" />
                      <input type="hidden" name="MM_recordId" value="<%= rstElements.Fields.Item("ElementID").Value %>" />
                      <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
                    </form>                    </td>
          <td>&nbsp;</td>
                  </tr>
<%
	End If
%>                  
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
<%
End If
%>            
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