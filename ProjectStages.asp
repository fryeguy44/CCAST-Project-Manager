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
    MM_editCmd.CommandText = "INSERT INTO dbo.ProjectStages (StageName, SortOrder) VALUES (?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxStageName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("tbxSortOrder"), Request.Form("tbxSortOrder"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim rstProjectStages
Dim rstProjectStages_cmd
Dim rstProjectStages_numRows

Set rstProjectStages_cmd = Server.CreateObject ("ADODB.Command")
rstProjectStages_cmd.ActiveConnection = MM_OBA_STRING
rstProjectStages_cmd.CommandText = "SELECT * FROM ProjectStages ORDER BY SortOrder" 
rstProjectStages_cmd.Prepared = true

Set rstProjectStages = rstProjectStages_cmd.Execute
rstProjectStages_numRows = 0
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
	<h1>Project Stage List</h1>	
	<table border="0" cellspacing="0" cellpadding="0" class="box3">
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="left"><h4>Stage Name</h4></th>
	    <th align="left"><h4>Sort Order</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
<%
If bolDeveloperViewGranted Then
    If bolDeveloperAddGranted Then
%>
      <form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
	  <tr>
	    <td>&nbsp;</td>
	    <td><input type="text" name="tbxStageName" id="tbxStageName" tabindex="0" /></td>
	    <td><input name="tbxSortOrder" type="text" id="tbxSortOrder" tabindex="1" size="5" /></td>
	    <td><input type="submit" name="btnAdd" id="btnAdd" value="Add Stage" /></td>
	    <td>&nbsp;</td>
      </tr>
      <input type="hidden" name="MM_insert" value="frmAdd" />
      </form>
	  <tr>
	    <td colspan="5"><hr /></td>
      </tr>
<%
    End If
	Do While Not rstProjectStages.EOF
		If bolDeveloperEditGranted Then
			strEdit = "<a href=""ProjectStageEdit.asp?lngProjectStageID=" & (rstProjectStages.Fields.Item("ProjectStageID").Value) & """>"
			strEditEnd = "</a>"
		Else
			strEdit = ""
			strEditEnd = ""
		End If 
%>      
	  <tr>
		<td>&nbsp;</td>
        <td><%=strEdit & (rstProjectStages.Fields.Item("StageName").Value) & strEditEnd%></td>
		<td><%=strEdit & (rstProjectStages.Fields.Item("SortOrder").Value) & strEditEnd%></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	  </tr>
<%
        rstProjectStages.MoveNext
    Loop
Else
%>  
        <tr>
            <td colspan="5">Viewing this list requires certain &quot;Projects&quot; permissions</td>
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
rstProjectStages.Close()
Set rstProjectStages = Nothing
%>
