<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim strPositionTitle
Dim strReturnPath

strPositionName = Request.Querystring("strPositionName")
lngPositionID = Request.Querystring("lngPositionID")
strReturnPath = Request.ServerVariables("HTTP_REFERER")
%>
<%
Dim rstElements__lngPositionID 
rstElements__lngPositionID  = 1
If (lngPositionID  <> "") Then 
  rstElements__lngPositionID  = lngPositionID 
End If
%>
<%
Dim rstElements
Dim rstElements_cmd
Dim rstElements_numRows

Set rstElements_cmd = Server.CreateObject ("ADODB.Command")
rstElements_cmd.ActiveConnection = MM_OBA_STRING
rstElements_cmd.CommandText = "SELECT Elements.ElementName, GrantLevels.LevelName, ISNULL(GrantLevels.GrantLevelID, 0) AS GrantLevelID, Elements.ElementID FROM (SELECT  GrantLevelID, ElementID FROM Grants  WHERE (PositionID = ?)) AS A INNER JOIN GrantLevels ON A.GrantLevelID = GrantLevels.GrantLevelID RIGHT OUTER JOIN Elements ON A.ElementID = Elements.ElementID ORDER BY Elements.ElementName" 
rstElements_cmd.Prepared = true
rstElements_cmd.Parameters.Append rstElements_cmd.CreateParameter("param1", 202, 1, 255, rstElements__lngPositionID) ' adDouble

Set rstElements = rstElements_cmd.Execute
rstElements_numRows = 0
%>
<%
Dim rstGrantLevels
Dim rstGrantLevels_cmd
Dim rstGrantLevels_numRows

Set rstGrantLevels_cmd = Server.CreateObject ("ADODB.Command")
rstGrantLevels_cmd.ActiveConnection = MM_OBA_STRING
rstGrantLevels_cmd.CommandText = "SELECT * FROM GrantLevels" 
rstGrantLevels_cmd.Prepared = true

Set rstGrantLevels = rstGrantLevels_cmd.Execute
rstGrantLevels_numRows = 0
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

<h1><%=strPageTitle%></h1>

			<table border="0" cellspacing="0" cellpadding="0" class="box">
<%	
If bolUsersViewGranted Then
'If True Then
%>
            <tr>
              <th colspan="3"><h3><strong>Security Grants for <%=strPositionName%></strong></h3></th>
              </tr>
			  <tr>
			  	<td>&nbsp;</td>
            	<td>&nbsp;</td>
            	<td>&nbsp;</td>
              </tr>
            <form action="PositionGrantsProcessor.asp" method="post" name="frmSecurityGrants" id="frmSecurityGrants">
<%
	Do While Not rstElements.EOF
%>            

            <tr>
            	<td>&nbsp;</td>
              <td><%=(rstElements.Fields.Item("ElementName").Value)%></td>
              <td>
                <select name="cbxGrantLevelID<%=(rstElements.Fields.Item("ElementID").Value)%>" id="cbxGrantLevelID<%=(rstElements.Fields.Item("ElementID").Value)%>">
                  <option value="0" <%If (Not isNull((rstElements.Fields.Item("GrantLevelID").Value))) Then If ("0" = CStr((rstElements.Fields.Item("GrantLevelID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>No Grant</option>
                  <%
While (NOT rstGrantLevels.EOF)
%>
                  <option value="<%=(rstGrantLevels.Fields.Item("GrantLevelID").Value)%>" <%If (Not isNull((rstElements.Fields.Item("GrantLevelID").Value))) Then If (CStr(rstGrantLevels.Fields.Item("GrantLevelID").Value) = CStr((rstElements.Fields.Item("GrantLevelID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstGrantLevels.Fields.Item("LevelName").Value)%></option>
                  <%
  rstGrantLevels.MoveNext()
Wend
If (rstGrantLevels.CursorType > 0) Then
  rstGrantLevels.MoveFirst
Else
  rstGrantLevels.Requery
End If
%>
                </select>              </td>
            </tr>
<%
		rstElements.MoveNext
	Loop
%>            
            <tr>
            	<td>&nbsp;</td>
              <td><input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
              <input name="htbxPositionTitle" type="hidden" id="htbxPositionTitle" value="<%=strPositionName%>" /></td>
              <input name="htbxPositionID" type="hidden" id="htbxPositionID" value="<%=lngPositionID%>" /></td>
              <td><input type="submit" name="btnUpdateGrants" id="btnUpdateGrants" value="Update" /></td>
            </tr>
            </form>
			    <tr>
			    	<td>&nbsp;</td>
			   	<td>&nbsp;</td>
            	<td>&nbsp;</td>
            	</tr>
<%
Else
%>              
                  <tr>
                    <td colspan="6">Certain &quot;Users&quot; permissions are required to view this information.</td>
                  </tr>
<%
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
<%
rstGrantLevels.Close()
Set rstGrantLevels = Nothing
%>
