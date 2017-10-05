<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngUserID
lngUserID = Request.QueryString("lngUserID")
%>
<%
If CStr(Request("MM_delete")) = "frmResetPassword" Then

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "UPDATE Users SET Password = null WHERE UserID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.QueryString("htbxUserID")) ' adDouble
Response.Write(MM_editCmd.CommandText) & "<br />"
Response.Write(Request.QueryString("htbxUserID")) & "<br />"
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

	Response.Redirect("Users.asp")
End If
%>
<%
If CStr(Request("MM_delete")) = "frmDelete" Then

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "DELETE FROM dbo.Users WHERE UserID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("htbxUserID")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

	Response.Redirect("Users.asp")
End If
%>
<%
If CStr(Request("MM_delete")) = "frmDeactivate" Then

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "UPDATE dbo.Users SET Active = " & Request.Form("htbxNewValue") & " WHERE UserID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("htbxUserID")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

	Response.Redirect("Users.asp")
End If
%>
<%
Dim rstUsers__lngUserID
rstUsers__lngUserID = 1
If (lngUserID <> "") Then 
  rstUsers__lngUserID = lngUserID
End If
%>
<%
Dim rstUsers
Dim rstUsers_cmd
Dim rstUsers_numRows

Set rstUsers_cmd = Server.CreateObject ("ADODB.Command")
rstUsers_cmd.ActiveConnection = MM_OBA_STRING
rstUsers_cmd.CommandText = "SELECT UserID, UserName, Positions.PositionName AS PositionTitle, Positions.PositionID, FirstName, LastName, Title, Active FROM Users INNER JOIN Positions ON Positions.PositionID = Users.PositionID WHERE UserID = ?" 
rstUsers_cmd.Prepared = true
rstUsers_cmd.Parameters.Append rstUsers_cmd.CreateParameter("param1", 5, 1, -1, rstUsers__lngUserID) ' adVarChar

Set rstUsers = rstUsers_cmd.Execute
rstUsers_numRows = 0
%>
<%
intPositionID = rstUsers.Fields.Item("PositionID").Value
strPositionTitle = (rstUsers.Fields.Item("PositionTitle").Value)
strUserName = (rstUsers.Fields.Item("UserName").Value)
%>
<%
Dim rstPages__intPositionID
rstPages__intPositionID = 1
If (intPositionID <> "") Then 
  rstPages__intPositionID = intPositionID
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
rstPages_cmd.Parameters.Append rstPages_cmd.CreateParameter("param1", 5, 1, -1, rstPages__intPositionID) ' adVarChar

Set rstPages = rstPages_cmd.Execute
rstPages_numRows = 0
%>
<%
Dim rstUserLog__lngUserID
rstUserLog__lngUserID = 1
If (lngUserID <> "") Then 
  rstUserLog__lngUserID = lngUserID
End If
%>
<%
Dim rstUserLog
Dim rstUserLog_cmd
Dim rstUserLog_numRows

Set rstUserLog_cmd = Server.CreateObject ("ADODB.Command")
rstUserLog_cmd.ActiveConnection = MM_OBA_STRING
rstUserLog_cmd.CommandText = "SELECT TOP (300) UserLogs.UserLogID, UserLogs.SIDPage, UserLogs.AccessDateTime, UserLogs.AccessGranted, AccessTypes.AccessTypeName,  UserLogs.UserName FROM UserLogs INNER JOIN AccessTypes ON UserLogs.AccessTypeID = AccessTypes.AccessTypeID INNER JOIN Users ON Users.UserID = UserLogs.UserID WHERE (UserLogs.UserID = ?) ORDER BY UserLogs.AccessDateTime DESC" 
rstUserLog_cmd.Prepared = true
rstUserLog_cmd.Parameters.Append rstUserLog_cmd.CreateParameter("param1", 5, 1, -1, rstUserLog__lngUserID) ' adVarChar

Set rstUserLog = rstUserLog_cmd.Execute
rstUserLog_numRows = 0
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
		
			  <h1>User Information</h1>

<%
If bolUsersViewGranted Then
	If rstUsers.Fields.Item("Active").Value = True Then
		strActiveButtonLabel = "Deactivate"
		strNewValue = 0
	Else
		strActiveButtonLabel = "Activate"
		strNewValue = 1
	End If
	
	If bolUsersEditGranted Then
		strEditLink = "<a href=""UserEdit.asp?lngUserID=" & (rstUsers.Fields.Item("UserID").Value) & """>"
		strEndEditLink = "</a>&nbsp;"
	Else
		strEditLink = ""
		strEndEditLink = "&nbsp;"
	End If
%>           
                <table border="0" cellspacing="0" cellpadding="0" class="info">
                  <tr>
                    <th align="left"><h4>&nbsp;</h4></th>
                    <th align="left"><h4><strong>Display Name</strong></h4></th>
                    <th align="left"><h4><strong>Username</strong></h4></th>
                    <th align="left" nowrap="nowrap"><h4><strong>CAS Position</strong></h4></th>
                    <th align="left" nowrap="nowrap"><h4>Title</h4></th>
                    <th align="left"><h4>&nbsp;</h4></th>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td><%=strEditLink & (rstUsers.Fields.Item("FirstName").Value) & " " & (rstUsers.Fields.Item("LastName").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstUsers.Fields.Item("Username").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstUsers.Fields.Item("PositionTitle").Value) & strEndEditLink%></td>
                    <td nowrap="nowrap"><%=strEditLink & (rstUsers.Fields.Item("Title").Value) & strEndEditLink%></td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                  	<td>&nbsp;</td>
                  	<td>&nbsp;</td>
                  	<td>
<%
	If bolUsersEditGranted Then
%>                   
                    <form id="frmDeactivate" name="frmDeactivate" method="post" action="">
                  	  <input type="submit" name="btnDeactivate" id="btnDeactivate" value="<%=strActiveButtonLabel%>" />
                    	      <input name="htbxUserID" type="hidden" id="htbxUserID" value="<%=(rstUsers.Fields.Item("UserID").Value)%>" />
               	          <input type="hidden" name="MM_delete" id="MM_delete" value="frmDeactivate" />
               	          <input type="hidden" name="htbxNewValue" id="htbxNewValue" value="<%=strNewValue%>" />
              	    </form>
<%
	Else
		Response.Write("&nbsp;")
	End If
%>	                    
                    </td>
                  	<td>
<%
	If bolUsersDeleteGranted Then
%>                   
					<form id="frmDelete" name="frmDelete" method="post" action="">
                  	  <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
                   	  <input type="hidden" name="htbxUserID" id="htbxUserID" value="<%=(rstUsers.Fields.Item("UserID").Value)%>" />
               	      <input type="hidden" name="MM_delete" id="MM_delete" value="frmDelete" />
              	    </form>
<%
	Else
		Response.Write("&nbsp;")
	End If
%>	                    
                    </td>
                  	<td><form id="frmResetPassword" name="frmResetPassword" method="get" action="">
               	      <input type="submit" name="btnResetPassword" id="btnResetPassword" value="Reset Password" />
                  	  <input type="hidden" name="htbxUserID" id="htbxUserID" value="<%=(rstUsers.Fields.Item("UserID").Value)%>" />
               	      <input type="hidden" name="MM_delete" id="MM_delete" value="frmResetPassword" />
                  	</form></td>
                    <td>&nbsp;</td>
                  </tr>
                </table>
				<table border="0" cellspacing="0" cellpadding="0" class="info">
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
                <table border="0" cellspacing="0" cellpadding="0" class="info">
                  <tr>
                    <th colspan="5"><h3>User Log</h3></th>
                  </tr>
                  <tr>
                    <td width="10">&nbsp;</td>
                    <td><h4>Page</h4></td>
                    <td><h4>Date Time</h4></td>
                    <td><h4>Granted</h4></td>
                    <td><h4>Type</h4></td>
                  </tr>
                  <tr>
                    <td colspan="5" class="line"><hr /></td>
                  </tr>

<%
	If rstUserLog.EOF Then
%>
                  <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td colspan="3">This User has not accessed the CAS</td>
                  </tr>

<%
	Else
		Do While Not rstUserLog.EOF
%>                 
                  <tr>
                    <td colspan="2"><a href="UserLogInformation.asp?lngUserLogID=<%=(rstUserLog.Fields.Item("UserLogID").Value)%>" class="info_link"><%=(rstUserLog.Fields.Item("SIDPage").Value)%></a></td>
                    <td><%=(rstUserLog.Fields.Item("AccessDateTime").Value)%></td>
                    <td><%=(rstUserLog.Fields.Item("AccessGranted").Value)%></td>
                    <td><%=(rstUserLog.Fields.Item("AccessTypeName").Value)%></td>
                  </tr>

<%
			rstUserLog.MoveNext
		Loop
	End If
%>              
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
rstUsers.Close()
Set rstUsers = Nothing
%>
<%
rstPages.Close()
Set rstPages = Nothing
%>
<%
rstUserLog.Close()
Set rstUserLog = Nothing
%>

