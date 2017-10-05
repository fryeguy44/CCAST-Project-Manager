<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%
lngUserLogID = Request.QueryString("lngUserLogID")
%>
<%
Dim rstUserLog__lngUserLogID
rstUserLog__lngUserLogID = "1"
If (lngUserLogID <> "") Then 
  rstUserLog__lngUserLogID = lngUserLogID
End If
%>
<%
Dim rstUserLog
Dim rstUserLog_cmd
Dim rstUserLog_numRows

Set rstUserLog_cmd = Server.CreateObject ("ADODB.Command")
rstUserLog_cmd.ActiveConnection = MM_OBA_STRING
rstUserLog_cmd.CommandText = "SELECT UserLogs.UserLogID, UserLogs.UserName, UserLogs.AccessTypeID, UserLogs.SIDPage, UserLogs.AccessDateTime, UserLogs.AccessGranted,  AccessTypes.AccessTypeName, Pages.PageTitle FROM UserLogs INNER JOIN AccessTypes ON UserLogs.AccessTypeID = AccessTypes.AccessTypeID INNER JOIN Pages ON UserLogs.SIDPage = Pages.PageAddress COLLATE SQL_Latin1_General_Pref_CP1_CI_AS INNER JOIN Users ON UserLogs.UserName = Users.UserName COLLATE SQL_Latin1_General_Pref_CP1_CI_AS WHERE UserLogID = ?" 
rstUserLog_cmd.Prepared = true
rstUserLog_cmd.Parameters.Append rstUserLog_cmd.CreateParameter("param1", 5, 1, -1, rstUserLog__lngUserLogID) ' adDouble

Set rstUserLog = rstUserLog_cmd.Execute
rstUserLog_numRows = 0
%>
<%
Dim rstLogParameters__lngUserLogID
rstLogParameters__lngUserLogID = "1"
If (lngUserLogID <> "") Then 
  rstLogParameters__lngUserLogID = lngUserLogID
End If
%>
<%
Dim rstLogParameters
Dim rstLogParameters_cmd
Dim rstLogParameters_numRows

Set rstLogParameters_cmd = Server.CreateObject ("ADODB.Command")
rstLogParameters_cmd.ActiveConnection = MM_OBA_STRING
rstLogParameters_cmd.CommandText = "SELECT * FROM UserLogPageParameters WHERE UserLogID = ?" 
rstLogParameters_cmd.Prepared = true
rstLogParameters_cmd.Parameters.Append rstLogParameters_cmd.CreateParameter("param1", 5, 1, -1, rstLogParameters__lngUserLogID) ' adDouble

Set rstLogParameters = rstLogParameters_cmd.Execute
rstLogParameters_numRows = 0
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
		
	<h1>User Log Information</h1>

<%
If bolSecurityViewGranted Then
'If True Then
%>           
	<table border="0" cellspacing="0" cellpadding="0" class="box">
	  <tr>
		<td><strong>User</strong></td>
		<td><%=(rstUserLog.Fields.Item("UserName").Value)%></td>
	  </tr>
	  <tr>
		<td><strong>Access Type</strong></td>
		<td><%=(rstUserLog.Fields.Item("AccessTypeName").Value)%></td>
	  </tr>
	  <tr>
	    <td><strong>Date/Time</strong></td>
	    <td><%=(rstUserLog.Fields.Item("AccessDateTime").Value)%></td>
      </tr>
	  <tr>
	    <td><strong>Page Address</strong></td>
	    <td><%=(rstUserLog.Fields.Item("SIDPage").Value)%></td>
      </tr>
	  <tr>
	    <td><strong>Page Title</strong></td>
	    <td><%=(rstUserLog.Fields.Item("PageTitle").Value)%></td>
      </tr>
	  
	  <tr>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
      </tr>
    </table>
    <table border="1" cellpadding="0" cellspacing="0" class="info">
	  <tr>
	    <th colspan="2"><h3>Parameters</h3></th>
      </tr>
	  <tr>
	    <td width="30%"><h4>Parameter</h4></td>
	    <td><h4>Value</h4></td>
      </tr>
	  <tr class="line">
	    <td height="1" colspan="2"><hr /></td>
      </tr>
<%
	intRep = 0
	Do While Not rstLogParameters.EOF
		intRep = intRep + 1
		If intRep = 1 Then
			strParameters = "?" 
		Else
			strParameters = strParameters & "&" 
		End If
		strParameters = strParameters &  (rstLogParameters.Fields.Item("InputField").Value) & "=" & (rstLogParameters.Fields.Item("InputValue").Value)
%>  
	  <tr>
	    <td width="30%"><%=(rstLogParameters.Fields.Item("InputField").Value)%></td>
	    <td><%=(rstLogParameters.Fields.Item("InputValue").Value)%></td>
      </tr>
<%
		rstLogParameters.MoveNext
	Loop
	If intRep = 0 Then
		strLink = "&nbsp;"
	Else
		strLink = "<a href=""" & rstUserLog.Fields.Item("SIDPage").Value & strParameters & """ class=""info_link"" target=""_blank"">View visited page</a>"
	End If

%>     
	  <tr>
		<td width="30%">&nbsp;</td>
		<td><%=strLink%></td>
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
rstUserLog.Close()
Set rstUserLog = Nothing
%>
<%
rstLogParameters.Close()
Set rstLogParameters = Nothing
%>
