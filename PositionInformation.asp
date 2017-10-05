<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngPositionID
lngPositionID = Request.QueryString("lngPositionID")
%>
<%
Dim rstUsers__lngPositionID
rstUsers__lngPositionID = "1"
If (lngPositionID <> "") Then 
  rstUsers__lngPositionID = lngPositionID
End If
%>
<%
Dim rstUsers
Dim rstUsers_cmd
Dim rstUsers_numRows

Set rstUsers_cmd = Server.CreateObject ("ADODB.Command")
rstUsers_cmd.ActiveConnection = MM_OBA_STRING
rstUsers_cmd.CommandText = "SELECT * FROM Users  WHERE Users.PositionID = ? ORDER BY LastName" 
rstUsers_cmd.Prepared = true
rstUsers_cmd.Parameters.Append rstUsers_cmd.CreateParameter("param1", 5, 1, -1, rstUsers__lngPositionID) ' adDouble

Set rstUsers = rstUsers_cmd.Execute
rstUsers_numRows = 0
%>
<%
Dim rstGrants__lngPositionID
rstGrants__lngPositionID = "1"
If (lngPositionID <> "") Then 
  rstGrants__lngPositionID = lngPositionID
End If
%>
<%
Dim rstGrants
Dim rstGrants_cmd
Dim rstGrants_numRows

Set rstGrants_cmd = Server.CreateObject ("ADODB.Command")
rstGrants_cmd.ActiveConnection = MM_OBA_STRING
rstGrants_cmd.CommandText = "SELECT Elements.ElementName, Grants.GrantID, Grants.ElementID, Grants.GrantLevelID, GrantLevels.LevelName  FROM Grants INNER JOIN GrantLevels ON Grants.GrantLevelID = GrantLevels.GrantLevelID INNER JOIN Elements ON Grants.ElementID = Elements.ElementID  WHERE (Grants.PositionID = ? )  ORDER BY Elements.ElementName" 
rstGrants_cmd.Prepared = true
rstGrants_cmd.Parameters.Append rstGrants_cmd.CreateParameter("param1", 5, 1, -1, rstGrants__lngPositionID) ' adDouble

Set rstGrants = rstGrants_cmd.Execute
rstGrants_numRows = 0
%>
<%
Dim rstPages__lngPositionID
rstPages__lngPositionID = 1
If (lngPositionID <> "") Then 
  rstPages__lngPositionID = lngPositionID
End If
%>
<%
Dim rstPages
Dim rstPages_cmd
Dim rstPages_numRows

Set rstPages_cmd = Server.CreateObject ("ADODB.Command")
rstPages_cmd.ActiveConnection = MM_OBA_STRING
rstPages_cmd.CommandText = "SELECT Pages.PageTitle, Pages.PageID FROM Pages INNER JOIN PageElements ON Pages.PageID = PageElements.PageID INNER JOIN Grants ON PageElements.ElementID = Grants.ElementID WHERE (Grants.PositionID = ?) GROUP BY Pages.PageTitle, Pages.PageID ORDER BY Pages.PageTitle" 
rstPages_cmd.Prepared = true
rstPages_cmd.Parameters.Append rstPages_cmd.CreateParameter("param1", 5, 1, -1, rstPages__lngPositionID) ' adVarChar

Set rstPages = rstPages_cmd.Execute
rstPages_numRows = 0
%>
<%
Dim rstPositions__lngPositionID
rstPositions__lngPositionID = "1"
If (lngPositionID <> "") Then 
  rstPositions__lngPositionID = lngPositionID
End If
%>
<%
Dim rstPositions
Dim rstPositions_cmd
Dim rstPositions_numRows

Set rstPositions_cmd = Server.CreateObject ("ADODB.Command")
rstPositions_cmd.ActiveConnection = MM_OBA_STRING
rstPositions_cmd.CommandText = "SELECT * FROM Positions WHERE PositionID = ?" 
rstPositions_cmd.Prepared = true
rstPositions_cmd.Parameters.Append rstPositions_cmd.CreateParameter("param1", 5, 1, -1, rstPositions__lngPositionID) ' adDouble

Set rstPositions = rstPositions_cmd.Execute
rstPositions_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr"><!-- InstanceBegin template="/Templates/Information.dwt" codeOutsideHTMLIsLocked="false" -->
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
<!-- jQuery UI -->
<link href="/global/jquery/css/ui-lightness/jquery-ui-1.10.3.custom.css" rel="stylesheet" />
<script type="text/javascript" src="/global/jquery/js/jquery-1.9.1.js"></script>
<script type="text/javascript" src="/global/jquery/js/jquery-ui-1.10.3.custom.js"></script>
<!-- Datepicker -->


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
		<h1>Positions</h1>
<%
		strPositionName = (rstPositions.Fields.Item("PositionName").Value)

'If bolUsersViewGranted Then
If True Then
%>
            <table border="0" cellpadding="0" cellspacing="0" class="box3">
                  <tr>
                    <th align="left" colspan="6"><h3>Position - <%=strPositionName%></h3></th>
                  </tr>
                  <tr>
                    <th align="left">&nbsp;</th>
                    <th align="left"><h4>Users</h4></th>
                    <th align="left"><h4>Title</h4></th>
                    <th align="left"><h4>UserName</h4></th>
                    <th align="left"><h4>Email</h4></th>
                    <th align="left">&nbsp;</th>
                  </tr>
<%
	Do While Not rstUsers.EOF
		strPositionName = (rstPositions.Fields.Item("PositionName").Value)
%>                  
                  <tr>
                    <td nowrap="nowrap">&nbsp;</td>
                    <td nowrap="nowrap"><%=(rstUsers.Fields.Item("FirstName").Value) & " " & (rstUsers.Fields.Item("LastName").Value)%></td>
                    <td nowrap="nowrap"><%=(rstUsers.Fields.Item("Title").Value)%></td>
                    <td nowrap="nowrap"><%=(rstUsers.Fields.Item("UserName").Value)%></td>
                    <td nowrap="nowrap"><%=(rstUsers.Fields.Item("EmailAddress").Value)%></td>
                    <td nowrap="nowrap">&nbsp;</td>
                  </tr>
<%
		rstUsers.MoveNext
	Loop
%>                  
                </table>
            <table border="0" cellspacing="0" cellpadding="0" class="box3">
                  <tr>
                    <th colspan="7"><h3>CAS Page Access</h3></th>
                  </tr>

<%
	If rstPages.EOF Then
%>
                  <tr>
                    <td>&nbsp;</td>
                    <td colspan="6">This User has no Access.</td>
                  </tr>

<%
	Else
%>               
                  
                  <tr>
                    <td>&nbsp;</td>
                    <td>                  
<%
		Do While Not rstPages.EOF
				Response.Write("<a href=""PageInformation.asp?lngPageID=" & (rstPages.Fields.Item("PageID").Value) & """>" & (rstPages.Fields.Item("PageTitle").Value) & "</a> | ") 
			rstPages.MoveNext
		Loop
%>					</td>
                  </tr>

<%
	End If
%>              
                </table>
    <table border="0" cellpadding="0" cellspacing="0" class="box3">
                  <tr>
                    <th colspan="2" align="left" nowrap="nowrap"><h3><a href="PositionGrants.asp?strPositionName=<%=strPositionName%>&lngPositionID=<%=lngPositionID%>"><%=strPositionName%> Security Grants</a></h3></th>
                  </tr>
<%
	If rstGrants.EOF Then
%>
                  <tr>
                    <td colspan="2"><a href="PositionGrants.asp?strPositionName=<%=strPositionName%>&lngPositionID=<%=lngPositionID%>">This Position has no Grants.</a></td>
                  </tr>
<%
	Else
%>               
                  <tr>
                    <th align="left" valign="middle"><h4><strong>Element</strong></h4></th>
                    <th align="left" nowrap="nowrap"><h4><strong>Grant Level</strong></h4></th>
                  </tr>
<%
		Do While Not rstGrants.EOF
%>	
                  <tr>
                	<td><a href="PositionGrants.asp?strPositionName=<%=strPositionName%>&lngPositionID=<%=lngPositionID%>"><%=(rstGrants.Fields.Item("ElementName").Value)%></a></td>
                    <td><a href="PositionGrants.asp?strPositionName=<%=strPositionName%>&lngPositionID=<%=lngPositionID%>"><%=(rstGrants.Fields.Item("LevelName").Value)%></a></td>
                  </tr>
<%
			rstGrants.MoveNext
		Loop
%>
				  <tr>
                    <td colspan="2">&nbsp;</td>
                  </tr>
<%
	End If
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
rstUsers.Close()
Set rstUsers = Nothing
%>
<%
rstGrants.Close()
Set rstGrants = Nothing
%>
<%
rstPositions.Close()
Set rstPositions = Nothing
%>
<%
rstPages.Close()
Set rstPages = Nothing
%>
