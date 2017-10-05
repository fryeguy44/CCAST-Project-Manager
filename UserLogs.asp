<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%
dteFromDate = Request.QueryString("tbxFromDate")
dteToDate = Request.QueryString("dteToDate")
If dteFromDate = "" Then
	If bolTestMachine = "Yes" Then
		dteFromDate = "6/1/2010"
	Else
		dteFromDate = Date() - 30
	End If
End If
If dteToDate = "" Then
	If bolTestMachine = "Yes" Then
		dteToDate = "6/30/2010"
	Else
		dteToDate = Date() + 1
	End If
End If

strUserName = Request.QueryString("strUserName")
If strUserName <> "" Then
	strWhere = " AND UserLogs.UserName = N'" & strUserName & "'"
End If

strPageAddress = Request.QueryString("strPageAddress")
If strPageAddress <> "" Then
	strWhere = strWhere & " AND UserLogs.SIDPage = N'" & strPageAddress & "'"
End If
lngShowAccessTypeID = Request.QueryString("lngShowAccessTypeID")
If lngShowAccessTypeID <> "" Then
	strWhere = strWhere & " AND UserLogs.AccessTypeID = " & lngShowAccessTypeID
End If
%>
<%
Dim rstUserLogs__dteFromDate
rstUserLogs__dteFromDate = "6/1/2010"
If (dteFromDate <> "") Then 
  rstUserLogs__dteFromDate = dteFromDate
End If
%>
<%
Dim rstUserLogs__dteToDate
rstUserLogs__dteToDate = "6/15/2010"
If (dteToDate <> "") Then 
  rstUserLogs__dteToDate = dteToDate
End If
%>
<%
Dim rstUserLogs
Dim rstUserLogs_cmd
Dim rstUserLogs_numRows

Set rstUserLogs_cmd = Server.CreateObject ("ADODB.Command")
rstUserLogs_cmd.ActiveConnection = MM_OBA_STRING
rstUserLogs_cmd.CommandText = "SELECT TOP 1500 UserLogs.UserLogID, UserLogs.UserName, AccessTypes.AccessTypeName, UserLogs.AccessTypeID, UserLogs.SIDPage,  UserLogs.AccessDateTime, UserLogs.AccessGranted, UserLogs.PageCompleted FROM UserLogs INNER JOIN AccessTypes ON UserLogs.AccessTypeID = AccessTypes.AccessTypeID INNER JOIN Users ON UserLogs.UserName = Users.UserName COLLATE SQL_Latin1_General_Pref_CP1_CI_AS WHERE (UserLogs.AccessDateTime > ? AND UserLogs.AccessDateTime < ?)" & strWhere & " ORDER BY UserLogs.AccessDateTime DESC" 
rstUserLogs_cmd.Prepared = true
rstUserLogs_cmd.Parameters.Append rstUserLogs_cmd.CreateParameter("param1", 135, 1, -1, rstUserLogs__dteFromDate) ' adDBTimeStamp
rstUserLogs_cmd.Parameters.Append rstUserLogs_cmd.CreateParameter("param2", 135, 1, -1, rstUserLogs__dteToDate) ' adDBTimeStamp

Set rstUserLogs = rstUserLogs_cmd.Execute
rstUserLogs_numRows = 0
%>
<%
Dim rstUsers
Dim rstUsers_cmd
Dim rstUsers_numRows

Set rstUsers_cmd = Server.CreateObject ("ADODB.Command")
rstUsers_cmd.ActiveConnection = MM_OBA_STRING
rstUsers_cmd.CommandText = "SELECT UserName FROM Users ORDER BY UserName" 
rstUsers_cmd.Prepared = true

Set rstUsers = rstUsers_cmd.Execute
rstUsers_numRows = 0
%>
<%
Dim rstPages
Dim rstPages_cmd
Dim rstPages_numRows

Set rstPages_cmd = Server.CreateObject ("ADODB.Command")
rstPages_cmd.ActiveConnection = MM_OBA_STRING
rstPages_cmd.CommandText = "SELECT PageTitle, PageAddress FROM dbo.Pages ORDER BY PageTitle" 
rstPages_cmd.Prepared = true

Set rstPages = rstPages_cmd.Execute
rstPages_numRows = 0
%>
<%
Dim rstAccessTypes
Dim rstAccessTypes_cmd
Dim rstAccessTypes_numRows

Set rstAccessTypes_cmd = Server.CreateObject ("ADODB.Command")
rstAccessTypes_cmd.ActiveConnection = MM_OBA_STRING
rstAccessTypes_cmd.CommandText = "SELECT * FROM dbo.AccessTypes" 
rstAccessTypes_cmd.Prepared = true

Set rstAccessTypes = rstAccessTypes_cmd.Execute
rstAccessTypes_numRows = 0
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
<!-- jQuery UI -->
<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/themes/base/jquery-ui.css" />
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.0/jquery-ui.min.js"></script>
<!-- Datepicker -->
<script type="text/javascript" charset="utf-16">
$(function() {
	$("#dteFromDate").datepicker({ maxDate: 0 });
});
$(function() {
	$("#tbxToDate").datepicker({ maxDate: 0 });
});
</script>

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

	<table border="0" cellspacing="0" cellpadding="0" class="fluid">
<%
If bolSecurityViewGranted OR bolDeveloperViewGranted Then
%>    
<form action="UserLogs.asp" method="get" name="frmDetailDates" id="frmDetailDates">
              <tr class="date_select">
                <th colspan="6">
				<p>Period From 
                    <input name="tbxFromDate" type="text" id="tbxFromDate" value="<%=dteFromDate%>" size="6" />
                  to 
                  <input name="tbxToDate" type="text" id="tbxToDate" value="<%=dteToDate%>" size="6" />
                  <select name="strUserName" id="strUserName">
                    <option value="" <%If (Not isNull(strUserName)) Then If ("" = CStr(strUserName)) Then Response.Write("selected=""selected""") : Response.Write("")%>>All Users</option>
                    <%
While (NOT rstUsers.EOF)
%><option value="<%=(rstUsers.Fields.Item("UserName").Value)%>" <%If (Not isNull(strUserName)) Then If (CStr(rstUsers.Fields.Item("UserName").Value) = CStr(strUserName)) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstUsers.Fields.Item("UserName").Value)%></option>
                    <%
  rstUsers.MoveNext()
Wend
If (rstUsers.CursorType > 0) Then
  rstUsers.MoveFirst
Else
  rstUsers.Requery
End If
%>
                  </select>
                  <select name="strPageAddress" id="strPageAddress">
                    <option value="" <%If (Not isNull(strPageAddress)) Then If ("" = CStr(strPageAddress)) Then Response.Write("selected=""selected""") : Response.Write("")%>>All Pages</option>
                    <%
While (NOT rstPages.EOF)
%><option value="<%=(rstPages.Fields.Item("PageAddress").Value)%>" <%If (Not isNull(strPageAddress)) Then If (CStr(rstPages.Fields.Item("PageAddress").Value) = CStr(strPageAddress)) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstPages.Fields.Item("PageTitle").Value)%></option>
                    <%
  rstPages.MoveNext()
Wend
If (rstPages.CursorType > 0) Then
  rstPages.MoveFirst
Else
  rstPages.Requery
End If
%>
                  </select>
                  <select name="lngShowAccessTypeID" id="lngShowAccessTypeID">
                    <option value="" <%If (Not isNull(lngShowAccessTypeID)) Then If ("" = CStr(lngShowAccessTypeID)) Then Response.Write("selected=""selected""") : Response.Write("")%>>All Access Types</option>
                    <%
While (NOT rstAccessTypes.EOF)
%><option value="<%=(rstAccessTypes.Fields.Item("AccessTypeID").Value)%>" <%If (Not isNull(lngShowAccessTypeID)) Then If (CStr(rstAccessTypes.Fields.Item("AccessTypeID").Value) = CStr(lngShowAccessTypeID)) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstAccessTypes.Fields.Item("AccessTypeName").Value)%></option>
                    <%
  rstAccessTypes.MoveNext()
Wend
If (rstAccessTypes.CursorType > 0) Then
  rstAccessTypes.MoveFirst
Else
  rstAccessTypes.Requery
End If
%>
                  </select>
                  <input type="submit" name="btnGo" id="btnGo" value="Go" />
                </p>				</th>
<%
If Not rstUserLogs.EOF Then
%>                
                <th colspan="2"><a href="UserLogsExcel.asp?<%=Server.HTMLEncode(Request.QueryString)%>" target="_blank" class="excel"></a></th>
<%
Else
	Response.Write("<th>&nbsp;</th>")     
End If
%>          
              </tr>
      </form>	  
      <tr>
		<td width="10">&nbsp;</td>
		<td><h4>Date/Time Requested </h4></td>
		<td align="left"><h4>Completed </h4></th>
		<td align="left" nowrap="nowrap"><h4>Load Time</h4></th>
		<td><h4>User</h4></td>
		<td><h4>Page</h4></td>
		<td><h4>Type Access</h4></td>
		<td><h4>Granted</h4></td>
	  </tr>
	  <tr class="line">
		<td colspan="8"><hr /></td>
	  </tr>
<%
Do While Not rstUserLogs.EOF
%>      
	  <tr class="tr_hover">
		<td colspan="2"><a href="UserLogInformation.asp?lngUserLogID=<%=(rstUserLogs.Fields.Item("UserLogID").Value)%>" class="info_link"><%=(rstUserLogs.Fields.Item("AccessDateTime").Value)%></a></td>
		<td><%=(rstUserLogs.Fields.Item("PageCompleted").Value)%></td>
		<td><%=DateDiff("s", rstUserLogs.Fields.Item("AccessDateTime").Value, rstUserLogs.Fields.Item("PageCompleted").Value)%></td>
		<td><%=(rstUserLogs.Fields.Item("UserName").Value)%></td>
		<td><%=(rstUserLogs.Fields.Item("SIDPage").Value)%></td>
		<td><%=(rstUserLogs.Fields.Item("AccessTypeName").Value)%></td>
		<td><%=(rstUserLogs.Fields.Item("AccessGranted").Value)%></td>
	  </tr>
<%
		rstUserLogs.MoveNext
	Loop
Else
%>      
      <tr>
		<td colspan="8">This page requires either "Developer" or "User" security grants to view</td>
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
rstUserLogs.Close()
Set rstUserLogs = Nothing
%>
<%
rstUsers.Close()
Set rstUsers = Nothing
%>
<%
rstPages.Close()
Set rstPages = Nothing
%>
<%
rstAccessTypes.Close()
Set rstAccessTypes = Nothing
%>
