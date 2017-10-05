<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim lngUserID
Dim strReturnPath

lngUserID = Request.QueryString("lngUserID")
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
    MM_editCmd.CommandText = "UPDATE dbo.Users SET UserName = ?, PositionID = ?, FirstName = ?, LastName = ?, Title = ?, EmailAddress = ?, LandingPageID = ?, VendorID = ?, ClientID = ?, Active = ? WHERE UserID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("tbxUserName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxPositionID"), Request.Form("cbxPositionID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("tbxFirstName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("tbxLastName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 50, Request.Form("tbxTitle")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 50, Request.Form("tbxEmailAddress")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("cbxLandingPageID"), Request.Form("cbxLandingPageID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("cbxVendorID"), Request.Form("cbxVendorID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("cbxClientID"), Request.Form("cbxClientID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 5, 1, -1, MM_IIF(Request.Form("chkActive"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
    MM_editCmd.CommandText = "DELETE FROM dbo.Users WHERE UserID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If

End If
%>
<%
Dim rstUsers__lngUserID
rstUsers__lngUserID = "1"
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
rstUsers_cmd.CommandText = "SELECT * FROM Users WHERE UserID = ?" 
rstUsers_cmd.Prepared = true
rstUsers_cmd.Parameters.Append rstUsers_cmd.CreateParameter("param1", 5, 1, -1, rstUsers__lngUserID) ' adDouble

Set rstUsers = rstUsers_cmd.Execute
rstUsers_numRows = 0
%>
<%
Dim rstPositions
Dim rstPositions_cmd
Dim rstPositions_numRows

Set rstPositions_cmd = Server.CreateObject ("ADODB.Command")
rstPositions_cmd.ActiveConnection = MM_OBA_STRING
rstPositions_cmd.CommandText = "SELECT * FROM Positions ORDER BY PositionName" 
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
If (CStr(Request("MM_update")) = "frmEdit") Then
	Response.Redirect(Request.Form("htbxReturnPath"))
End If
If (CStr(Request("MM_delete")) = "frmDelete" And CStr(Request("MM_recordId")) <> "") Then
	Response.Redirect("Users.asp")
End If

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
		
	<table border="0" cellspacing="0" cellpadding="0" class="box">
        <tr>
            <th colspan="4"><h2><%=strPageTitle & " " & strSubTitle%></h2></th>
      </tr>
<%
If bolDeveloperEditGranted Then
	If rstUsers.EOF Then
%>  
        <tr>
            <td colspan="4"><a href="Users.asp">The User you are attempting to edit has been deleted. Click here to return to the User List page</a></td>
        </tr>
<%
	Else
%>     
    	<form id="frmEdit" name="frmEdit" method="POST" action="<%=MM_editAction%>">
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>User Name</strong></td>
          <td><input name="tbxUserName" type="text" id="tbxUserName" tabindex="0" value="<%=(rstUsers.Fields.Item("UserName").Value)%>" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Position Title</strong></td>
          <td><select name="cbxPositionID" id="cbxPositionID">
            <%
While (NOT rstPositions.EOF)
%>
            <option value="<%=(rstPositions.Fields.Item("PositionID").Value)%>" <%If (Not isNull((rstUsers.Fields.Item("PositionID").Value))) Then If (CStr(rstPositions.Fields.Item("PositionID").Value) = CStr((rstUsers.Fields.Item("PositionID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstPositions.Fields.Item("PositionName").Value)%></option>
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
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>First Name</strong></td>
          <td><input name="tbxFirstName" type="text" id="tbxFirstName" tabindex="2" value="<%=(rstUsers.Fields.Item("FirstName").Value)%>" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Last Name</strong></td>
          <td><input name="tbxLastName" type="text" id="tbxLastName" tabindex="3" value="<%=(rstUsers.Fields.Item("LastName").Value)%>" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Title</strong></td>
          <td><input name="tbxTitle" type="text" id="tbxTitle" tabindex="4" value="<%=(rstUsers.Fields.Item("Title").Value)%>" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Email</strong></td>
          <td><input name="tbxEmailAddress" type="text" id="tbxEmailAddress" tabindex="5" value="<%=(rstUsers.Fields.Item("EmailAddress").Value)%>" size="30" /></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Landing Page</strong></td>
          <td><select name="cbxLandingPageID" id="cbxLandingPageID">
            <%
While (NOT rstPages.EOF)
%>
            <option value="<%=(rstPages.Fields.Item("PageID").Value)%>" <%If (Not isNull((rstUsers.Fields.Item("LandingPageID").Value))) Then If (CStr(rstPages.Fields.Item("PageID").Value) = CStr((rstUsers.Fields.Item("LandingPageID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstPages.Fields.Item("PageTitle").Value)%></option>
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
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Vendor</strong></td>
          <td><select name="cbxVendorID" id="cbxVendorID">
            <option value="0" <%If (Not isNull((rstUsers.Fields.Item("VendorID").Value))) Then If ("0" = CStr((rstUsers.Fields.Item("VendorID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>Not a Vendor</option>
            <%
While (NOT rstVendors.EOF)
%>
<option value="<%=(rstVendors.Fields.Item("VendorID").Value)%>" <%If (Not isNull((rstUsers.Fields.Item("VendorID").Value))) Then If (CStr(rstVendors.Fields.Item("VendorID").Value) = CStr((rstUsers.Fields.Item("VendorID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstVendors.Fields.Item("VendorName").Value)%></option>
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
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="right"><strong>Client</strong></td>
          <td><select name="cbxClientID" id="cbxClientID">
            <option value="0" <%If (Not isNull((rstUsers.Fields.Item("ClientID").Value))) Then If ("0" = CStr((rstUsers.Fields.Item("ClientID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>Not A Client</option>
            <%
While (NOT rstClients.EOF)
%>
<option value="<%=(rstClients.Fields.Item("ClientID").Value)%>" <%If (Not isNull((rstUsers.Fields.Item("ClientID").Value))) Then If (CStr(rstClients.Fields.Item("ClientID").Value) = CStr((rstUsers.Fields.Item("ClientID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rstClients.Fields.Item("ClientName").Value)%></option>
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
          <td>&nbsp;</td>
        </tr>
        <tr>
            <td width="10">&nbsp;</td>
            <td align="right"><strong>Active</strong></td>
            <td><input <%If (CStr((rstUsers.Fields.Item("Active").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="chkActive" type="checkbox" id="chkActive" tabindex="6" value="True" /></td>
		<td>&nbsp;</td>
		</tr>
        <tr>
            <td width="10">&nbsp;</td>
            <td>&nbsp;</td>
            <td><input type="submit" name="btnEdit" id="btnEdit" value="Update" tabindex="7" /></td>
            <td>&nbsp;</td>
      </tr>
        <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
        <input type="hidden" name="MM_update" value="frmEdit" />
        <input type="hidden" name="MM_recordId" value="<%= rstUsers.Fields.Item("UserID").Value %>" />
        </form>
<%
		If bolDeveloperDeleteGranted Then
%>                
      <tr>
        <td width="10">&nbsp;</td>
            <td><form id="frmDelete" name="frmDelete" method="POST" action="<%=MM_editAction%>">
              <input type="submit" name="btnDelete" id="btnDelete" value="Delete" />
              <input name="htbxReturnPath" type="hidden" id="htbxReturnPath" value="<%=strReturnPath%>" />
              <input type="hidden" name="MM_delete" value="frmDelete" />
              <input type="hidden" name="MM_recordId" value="<%= rstUsers.Fields.Item("UserID").Value %>" />
            </form>            </td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
      </tr>
<%
		End If
	End If
Else
%>
        <tr>
            <td colspan="4">Certain &quot;Developer&quot; permissions are required to perform this task.</td>
        </tr>

<%

End If
%>
        <tr>
            <td width="10">&nbsp;</td>
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
rstUsers.Close()
Set rstUsers = Nothing
%>
<%
rstPositions.Close()
Set rstPositions = Nothing
%>
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
