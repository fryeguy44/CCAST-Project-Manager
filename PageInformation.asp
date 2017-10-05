<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%
Dim lngPageID
lngPageID = Request.QueryString("lngPageID")
%>
<%
Dim rstPages__lngPageID
rstPages__lngPageID = "2"
If (lngPageID  <> "") Then 
  rstPages__lngPageID = lngPageID 
End If
%>
<%
Dim rstPages
Dim rstPages_cmd
Dim rstPages_numRows

Set rstPages_cmd = Server.CreateObject ("ADODB.Command")
rstPages_cmd.ActiveConnection = MM_OBA_STRING
rstPages_cmd.CommandText = "SELECT PageID, PageGroupID, PageTitle, NavigationPage, PageAddress, HelpContextID FROM Pages WHERE (Pages.PageID = ?)" 
rstPages_cmd.Prepared = true
rstPages_cmd.Parameters.Append rstPages_cmd.CreateParameter("param1", 5, 1, -1, rstPages__lngPageID) ' adDouble

Set rstPages = rstPages_cmd.Execute
rstPages_numRows = 0
%>
<%
Dim strPageAddress
strPageAddress = (rstPages.Fields.Item("PageAddress").Value)
%>
<%
Dim rstUserLogs__strPageAddress
rstUserLogs__strPageAddress = "PageInformation.asp"
If (strPageAddress <> "") Then 
  rstUserLogs__strPageAddress = strPageAddress
End If
%>
<%
Dim rstUserLogs
Dim rstUserLogs_cmd
Dim rstUserLogs_numRows

Set rstUserLogs_cmd = Server.CreateObject ("ADODB.Command")
rstUserLogs_cmd.ActiveConnection = MM_OBA_STRING
rstUserLogs_cmd.CommandText = "SELECT UserLogs.SIDPage, UserLogs.UserName, AccessTypes.AccessTypeName, MAX(UserLogs.AccessDateTime) AS LastAccess FROM UserLogs INNER JOIN AccessTypes ON UserLogs.AccessTypeID = AccessTypes.AccessTypeID GROUP BY UserLogs.SIDPage, AccessTypes.AccessTypeName, UserLogs.UserName HAVING (UserLogs.SIDPage = ?) ORDER BY LastAccess DESC" 
rstUserLogs_cmd.Prepared = true
rstUserLogs_cmd.Parameters.Append rstUserLogs_cmd.CreateParameter("param1", 202, 1, 255, rstUserLogs__strPageAddress) ' adVarChar

Set rstUserLogs = rstUserLogs_cmd.Execute
rstUserLogs_numRows = 0
%>
<%
Dim rstPageElements__lngPageID
rstPageElements__lngPageID = "5"
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
rstPageElements_cmd.CommandText = "SELECT PageElements.PageElementID, PageElements.PageID, PageElements.ElementID, Elements.ElementName FROM PageElements INNER JOIN Elements ON PageElements.ElementID = Elements.ElementID WHERE PageID = ? ORDER BY Elements.ElementName" 
rstPageElements_cmd.Prepared = true
rstPageElements_cmd.Parameters.Append rstPageElements_cmd.CreateParameter("param1", 5, 1, -1, rstPageElements__lngPageID) ' adDouble

Set rstPageElements = rstPageElements_cmd.Execute
rstPageElements_numRows = 0
%>
<%
Dim rstUsers__lngPageID
rstUsers__lngPageID = "5"
If (lngPageID <> "") Then 
  rstUsers__lngPageID = lngPageID
End If
%>
<%
Dim rstUsers
Dim rstUsers_cmd
Dim rstUsers_numRows

Set rstUsers_cmd = Server.CreateObject ("ADODB.Command")
rstUsers_cmd.ActiveConnection = MM_OBA_STRING
rstUsers_cmd.CommandText = "SELECT     GrantLevels.LevelName, Grants.PositionID, Elements.ElementName, Elements.ElementID, PageElements.PageID, Positions.PositionName FROM Grants INNER JOIN Elements ON Grants.ElementID = Elements.ElementID INNER JOIN PageElements ON Elements.ElementID = PageElements.ElementID INNER JOIN GrantLevels ON Grants.GrantLevelID = GrantLevels.GrantLevelID INNER JOIN Positions ON Positions.PositionID = Grants.PositionID WHERE (PageElements.PageID = ?) ORDER BY Positions.PositionName, Elements.ElementName" 
rstUsers_cmd.Prepared = true
rstUsers_cmd.Parameters.Append rstUsers_cmd.CreateParameter("param1", 5, 1, -1, rstUsers__lngPageID) ' adDouble

Set rstUsers = rstUsers_cmd.Execute
rstUsers_numRows = 0
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
		
	<h1><%=strPageTitle & " " & strSubTitle%></h1>

<%
Dim fs
Dim f
Dim strFileDate
Set fs=Server.CreateObject("Scripting.FileSystemObject")
'response.Write(rstPages.Fields.Item("PageAddress").Value)
'response.Write(Server.MapPath(rstPages.Fields.Item("PageAddress").Value))
Set f=fs.GetFile(Server.MapPath(rstPages.Fields.Item("PageAddress").Value))
strFileDate = f.DateCreated
set f=nothing
set fs=nothing
%>
			<table border="0" cellspacing="0" cellpadding="0" class="box">
              <tr>
                <th colspan="3"><h2><%=strPageTitle & " " & strSubTitle%></h2></th>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><strong>Page Title</strong></td>
                <td><%=(rstPages.Fields.Item("PageTitle").Value)%></td>
              </tr>
              
              <tr>
                <td>&nbsp;</td>
                <td><strong>Page Address</strong></td>
                <td><%=(rstPages.Fields.Item("PageAddress").Value)%></td>
      </tr>
              <tr>
                <td>&nbsp;</td>
                <td><strong>Modified</strong></td>
                <td><%=FormatDateTime(strFileDate,2)%></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><strong>Help Page</strong></td>
                <td><%=(rstPages.Fields.Item("HelpContextID").Value)%></td>
      		  </tr>
            </table>
            
<table cellpadding="0" cellspacing="0" class="box">
              <tr>
                <th colspan="4"><h2 style="text-align:center;"><a href="PageElements.asp?lngPageID=<%=(rstPages.Fields.Item("PageID").Value)%>">Elements On Page</a></h2></th>
              </tr>
<%
If rstPageElements.EOF Then
%>
              <tr>
                <td colspan="4" align="center"><a href="PageElements.asp?lngPageID=<%=(rstPages.Fields.Item("PageID").Value)%>">This Page has no Elements.</a></td>
      </tr>

<%
Else
%>               
              <tr class="column_titles">
                <td align="center"><h4>Element Name</h4></td>
              </tr>
              <tr class="line">
                <td><hr /></td>
              </tr>
<%
	Do While Not rstPageElements.EOF
%>	
              <tr>
                <td align="center" nowrap="nowrap"><a href="PageElementEdit.asp?lngPageElementID=<%=(rstPageElements.Fields.Item("PageElementID").Value)%>"><%=(rstPageElements.Fields.Item("ElementName").Value)%></a></td> 
              </tr>
<%
		rstPageElements.MoveNext
	Loop
End If
%>            
    </table>
<table cellpadding="0" cellspacing="0" class="box">
              <tr>
                <th colspan="4" align="center"><h2 style="text-align:center;">Page Access</h2></th>
              </tr>
<%
If rstUsers.EOF Then
%>
              <tr>
                <td align="center" colspan="3"><a href="PageElements.asp?lngPageID=<%=(rstPages.Fields.Item("PageID").Value)%>">This Page has no Users.</a></td>
              </tr>

<%
Else
%>               
              <tr class="column_titles">
                <td><h4>Position</h4></td>
                <td><h4>Element</h4></td>
                <td><h4>Grant Level</h4></td>
              </tr>
              <tr class="line">
                <td colspan="3"><hr /></td> 
              </tr>
<%
Dim strDisplayName
Dim bolWriteDisplayName
	strDisplayName = ""
	bolWriteDisplayName = False
	Do While Not rstUsers.EOF
		If strDisplayName <>  (rstUsers.Fields.Item("PositionName").Value) Then
			strDisplayName =  (rstUsers.Fields.Item("PositionName").Value)
			bolWriteDisplayName = True
		End If
%>		                  
              <tr>
                <td><a href="PositionGrants.asp?strPositionTitle=<%=(rstUsers.Fields.Item("PositionName").Value)%>"><%If bolWriteDisplayName Then Response.Write(strDisplayName) Else Response.Write("&nbsp") End If %></a></td> 
                <td><%=(rstUsers.Fields.Item("ElementName").Value)%></td>
                <td><%=(rstUsers.Fields.Item("LevelName").Value)%></td>
              </tr>
<%
		rstUsers.MoveNext
		bolWriteDisplayName = False
	Loop
End If
%>              

</table>            
<table class="box" border="1" cellpadding="0" cellspacing="0">
              <tr>
                <th colspan="4" align="center"><h2 style="text-align:center;">Page Access Log</h2></th>
              </tr>
              <tr class="column_titles">
              	<td width="10">&nbsp;</td>
              	<td><h4>User</h4></td>
              	<td><h4>Type</h4></td>
              	<td><h4>Last Access Date</h4></td>
              </tr>
              <tr class="line">
              	<td colspan="4"><hr /></td>
           	  </tr>
<%
Do While Not rstUserLogs.EOF
%>              
              <tr>
              	<td>&nbsp;</td>
              	<td><%=(rstUserLogs.Fields.Item("UserName").Value)%></td>
              	<td><%=(rstUserLogs.Fields.Item("AccessTypeName").Value)%></td>
              	<td><%=(rstUserLogs.Fields.Item("LastAccess").Value)%></td>
              </tr>
<%
	rstUserLogs.MoveNext
Loop
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
rstPages.Close()
Set rstPages = Nothing
%>
<%
rstUserLogs.Close()
Set rstUserLogs = Nothing
%>
<%
rstPageElements.Close()
Set rstPageElements = Nothing
%>
<%
rstUsers.Close()
Set rstUsers = Nothing
%>
