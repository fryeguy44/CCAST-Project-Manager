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
    MM_editCmd.CommandText = "INSERT INTO dbo.Users (FirstName, LastName, UserName, PositionID, Title, EmailAddress, LandingPageID, VendorID, ClientID) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxFirstName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("tbxLastName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("tbxUsername")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("cbxPositionID"), Request.Form("cbxPositionID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 50, Request.Form("tbxTitle")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 50, Request.Form("tbxEmailAddress")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("cbxLandingPageID"), Request.Form("cbxLandingPageID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("cbxVendorID"), Request.Form("cbxVendorID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("cbxClientID"), Request.Form("cbxClientID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
strOrder = Request.QueryString("strOrder")

	Select Case strOrder
		Case "Name"
			strOrderClause = " ORDER BY LastName"
			strNameSortImage = "_Down"
		Case "UserName"
			strOrderClause = " ORDER BY UserName"
			strUserNameSortImage = "_Down"
		Case Else
			strOrderClause = " ORDER BY PositionName, UserName"
			strPositionSortImage = "_Down"
	End Select

%>
<%
Dim rstUsers
Dim rstUsers_cmd
Dim rstUsers_numRows

Set rstUsers_cmd = Server.CreateObject ("ADODB.Command")
rstUsers_cmd.ActiveConnection = MM_OBA_STRING
rstUsers_cmd.CommandText = "SELECT Users.UserID, Users.UserName, Positions.PositionName, Users.FirstName, Users.LastName, Users.Title, Users.Active, Users.EmailAddress, Vendors.VendorName, Pages.PageTitle AS LandingPage, Clients.ClientName FROM Users INNER JOIN Positions ON Users.PositionID = Positions.PositionID INNER JOIN Pages ON Users.LandingPageID = Pages.PageID LEFT OUTER JOIN Clients ON Users.ClientID = Clients.ClientID LEFT OUTER JOIN Vendors ON Users.VendorID = Vendors.VendorID" & strOrderClause
rstUsers_cmd.Prepared = true

Set rstUsers = rstUsers_cmd.Execute
rstUsers_numRows = 0
%>
<%
Dim rstPositions
Dim rstPositions_cmd
Dim rstPositions_numRows

Set rstPositions_cmd = Server.CreateObject ("ADODB.Command")
rstPositions_cmd.ActiveConnection = MM_OBA_STRING
rstPositions_cmd.CommandText = "SELECT PositionID, PositionName FROM Positions ORDER BY PositionName" 
rstPositions_cmd.Prepared = true

Set rstPositions = rstPositions_cmd.Execute
rstPositions_numRows = 0
%>
<%
Dim rstVendors
Dim rstVendors_cmd
Dim rstVendors_numRows

Set rstVendors_cmd = Server.CreateObject ("ADODB.Command")
rstVendors_cmd.ActiveConnection = MM_OBA_STRING
rstVendors_cmd.CommandText = "SELECT VendorID, VendorName FROM Vendors ORDER BY VendorName" 
rstVendors_cmd.Prepared = true

Set rstVendors = rstVendors_cmd.Execute
rstVendors_numRows = 0
%>
<%
Dim rstPages
Dim rstPages_cmd
Dim rstPages_numRows

Set rstPages_cmd = Server.CreateObject ("ADODB.Command")
rstPages_cmd.ActiveConnection = MM_OBA_STRING
rstPages_cmd.CommandText = "SELECT PageID, PageTitle FROM Pages WHERE NavigationPage = 1 ORDER BY PageTitle" 
rstPages_cmd.Prepared = true

Set rstPages = rstPages_cmd.Execute
rstPages_numRows = 0
%>
<%
Dim rstClients
Dim rstClients_cmd
Dim rstClients_numRows

Set rstClients_cmd = Server.CreateObject ("ADODB.Command")
rstClients_cmd.ActiveConnection = MM_OBA_STRING
rstClients_cmd.CommandText = "SELECT ClientID, ClientName FROM Clients ORDER BY ClientName" 
rstClients_cmd.Prepared = true

Set rstClients = rstClients_cmd.Execute
rstClients_numRows = 0
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

<!-- InstanceBeginEditable name="Head" --><!-- InstanceEndEditable -->

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


<%
If bolDeveloperViewGranted Then
		

%>
            <table border="0" cellspacing="0" cellpadding="0" class="box">
              <tr>
                <th colspan="12" align="center"><h2>Users</h2></th>
              </tr>
              <tr>
                <th>&nbsp;</th>
                <th colspan="2" align="left"><a href="Users.asp?strOrder=Name"><h4>Display Name&nbsp;<img border="0" src="global/images/arrow<%=strNameSortImage%>.gif" width="8" height="10" align="absmiddle" alt=""></h4></a></th>
                <th align="left"><a href="Users.asp?strOrder=UserName"><h4>Username&nbsp;<img border="0" src="global/images/arrow<%=strUserNameSortImage%>.gif" width="8" height="10" align="absmiddle" alt=""></h4></a></th>
                <th align="left"><a href="Users.asp?strOrder=Position"><h4>Position&nbsp;<img border="0" src="global/images/arrow<%=strPositionSortImage%>.gif" width="8" height="10" align="absmiddle" alt=""></h4></a></th>
                <th align="left"><h4>Title</h4></th>
                <th align="left"><h4>Email</h4></th>
                <th align="left"><h4>Landing Page</h4></th>
                <th align="left"><h4>Vendor</h4></th>
                <th align="center"><h4>Client</h4></th>
                <th align="center"><h4>Active</h4></th>
                <th align="left">&nbsp;</th>
              </tr>
<%
	If bolDeveloperAddGranted Then
%>              
		<form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
              <tr class="tr_hover">
                <td>&nbsp;</td>
                <td nowrap="nowrap"><input name="tbxFirstName" type="text" id="tbxFirstName" placeholder="User First Name" tabindex="1" size="11" /></td>
                <td nowrap="nowrap"><input name="tbxLastName" type="text" id="tbxLastName" placeholder="User Last Name" tabindex="2" size="11" /></td>
                <td><input type="text" name="tbxUsername" id="tbxUsername" tabindex="3" placeholder="UserName" /></td>
                <td><select name="cbxPositionID" id="cbxPositionID" tabindex="4">
                  <%
While (NOT rstPositions.EOF)
%>
                  <option value="<%=(rstPositions.Fields.Item("PositionID").Value)%>"><%=(rstPositions.Fields.Item("PositionName").Value)%></option>
                  <%
  rstPositions.MoveNext()
Wend
If (rstPositions.CursorType > 0) Then
  rstPositions.MoveFirst
Else
  rstPositions.Requery
End If
%>
                </select></td>
                <td><input name="tbxTitle" type="text" id="tbxTitle" placeholder="User Job Title" tabindex="5" size="15" /></td>
                <td><input name="tbxEmailAddress" type="text" id="tbxEmailAddress" placeholder="Email Address" tabindex="6" /></td>
                <td><select name="cbxLandingPageID" id="cbxLandingPageID">
                  <%
While (NOT rstPages.EOF)
%>
                  <option value="<%=(rstPages.Fields.Item("PageID").Value)%>" <%If (Not isNull("22")) Then If (CStr(rstPages.Fields.Item("PageID").Value) = CStr("22")) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstPages.Fields.Item("PageTitle").Value)%></option>
                  <%
  rstPages.MoveNext()
Wend
If (rstPages.CursorType > 0) Then
  rstPages.MoveFirst
Else
  rstPages.Requery
End If
%>
                </select></td>
                <td><select name="cbxVendorID" id="cbxVendorID">
                  <option value="0">Not A Vendor</option>
                  <%
While (NOT rstVendors.EOF)
%>
                  <option value="<%=(rstVendors.Fields.Item("VendorID").Value)%>"><%=(rstVendors.Fields.Item("VendorName").Value)%></option>
                  <%
  rstVendors.MoveNext()
Wend
If (rstVendors.CursorType > 0) Then
  rstVendors.MoveFirst
Else
  rstVendors.Requery
End If
%>
                </select></td>
                <td align="center"><select name="cbxClientID" id="cbxClientID">
                  <option value="0">Not A Client</option>
                  <%
While (NOT rstClients.EOF)
%>
                  <option value="<%=(rstClients.Fields.Item("ClientID").Value)%>"><%=(rstClients.Fields.Item("ClientName").Value)%></option>
<%
  rstClients.MoveNext()
Wend
If (rstClients.CursorType > 0) Then
  rstClients.MoveFirst
Else
  rstClients.Requery
End If
%>
                </select></td>
                <td align="center"><input type="submit" name="btnAdd" id="btnAdd" value="Add" tabindex="7" /></td>
                <td>&nbsp;</td>
              </tr>
              <input type="hidden" name="MM_insert" value="frmAdd" />
        </form>	
              <tr class="tr_hover">
                <td colspan="12"><hr /></td>
              </tr>
        	  		
<%
	End If
	
	intCurrentUserID = 0
	Do While Not rstUsers.EOF  
		If bolDeveloperEditGranted Then
			strEditLink = "<a href=""UserEdit.asp?lngUserID=" & (rstUsers.Fields.Item("UserID").Value) & """>"
			strEndEditLink = "</a>&nbsp;"
		Else
			strEditLink = ""
			strEndEditLink = "&nbsp;"
		End If

		intFlushCount = intFlushCount + 1

		If intFlushCount > 100 Then
			intFlushCount = 0
			Response.Flush()
		End If
%>            
              <tr class="tr_hover">
                <td><a href="UserInformation.asp?lngUserID=<%=(rstUsers.Fields.Item("UserID").Value)%>" class="row_info"></a></td>
                <td colspan="2" nowrap="nowrap"><%=strEditLink & (rstUsers.Fields.Item("FirstName").Value) & " " & (rstUsers.Fields.Item("LastName").Value) & strEndEditLink%></td>
                <td><%=strEditLink & (rstUsers.Fields.Item("UserName").Value) & strEndEditLink%></td>
                <td><%=strEditLink & (rstUsers.Fields.Item("PositionName").Value) & strEndEditLink%></td>
                <td align="Left" nowrap="nowrap"><%=strEditLink & rstUsers.Fields.Item("Title").Value & strEndEditLink%></td>
                <td nowrap="nowrap"><%=strEditLink & rstUsers.Fields.Item("EmailAddress").Value & strEndEditLink%></td>
                <td nowrap="nowrap"><%=strEditLink & rstUsers.Fields.Item("LandingPage").Value & strEndEditLink%></td>
                <td nowrap="nowrap"><%=strEditLink & rstUsers.Fields.Item("VendorName").Value & strEndEditLink%></td>
                <td nowrap="nowrap"><%=strEditLink & rstUsers.Fields.Item("ClientName").Value & strEndEditLink%></td>
                <td align="center"<%=strActiveAlert%>><%=strEditLink & (rstUsers.Fields.Item("Active").Value) & strEndEditLink%></td>
                <td nowrap="nowrap">&nbsp;</td>
              </tr>
<%
		rstUsers.MoveNext
	Loop

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
rstVendors.Close()
Set rstVendors = Nothing
%>
<%
rstPages.Close()
Set rstPages = Nothing
%>
<%
rstClients.Close()
Set rstClients = Nothing
%>
<%
rstUsers.Close()
Set rstUsers = Nothing
%>
<%
rstPositions.Close()
Set rstPositions = Nothing
%>
