<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

<%

Dim lngElementID

lngElementID = Request.QueryString("lngElementID")
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
If (CStr(Request("MM_insert")) = "frmAdd") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Grants (UserID, GrantLevelID, ElementID) VALUES (?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("cbxUserID"), Request.Form("cbxUserID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("cbxGrantLevelID"), Request.Form("cbxGrantLevelID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("htbxElementID"), Request.Form("htbxElementID"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
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
rstElements_cmd.CommandText = "SELECT ElementName, ElementID FROM Elements WHERE ElementID = ? ORDER BY Elements.ElementName" 
rstElements_cmd.Prepared = true
rstElements_cmd.Parameters.Append rstElements_cmd.CreateParameter("param1", 5, 1, -1, rstElements__lngElementID) ' adDouble

Set rstElements = rstElements_cmd.Execute
rstElements_numRows = 0
%>
<%
Dim rstGrants__lngElementID
rstGrants__lngElementID = "1"
If (lngElementID <> "") Then 
  rstGrants__lngElementID = lngElementID
End If
%>
<%
Dim rstGrants
Dim rstGrants_cmd
Dim rstGrants_numRows

Set rstGrants_cmd = Server.CreateObject ("ADODB.Command")
rstGrants_cmd.ActiveConnection = MM_OBA_STRING
rstGrants_cmd.CommandText = "SELECT Grants.ElementID, Grants.GrantID, GrantLevels.LevelName, Grants.GrantLevelID, Positions.PositionName FROM Grants INNER JOIN GrantLevels ON Grants.GrantLevelID = GrantLevels.GrantLevelID INNER JOIN Positions ON Positions.PositionID = Grants.PositionID WHERE ElementID = ? ORDER BY Positions.PositionName" 
rstGrants_cmd.Prepared = true
rstGrants_cmd.Parameters.Append rstGrants_cmd.CreateParameter("param1", 5, 1, -1, rstGrants__lngElementID) ' adDouble

Set rstGrants = rstGrants_cmd.Execute
rstGrants_numRows = 0
%>
<%
Dim rstElementPages__lngElementID
rstElementPages__lngElementID = "1"
If (lngElementID <> "") Then 
  rstElementPages__lngElementID = lngElementID
End If
%>
<%
Dim rstElementPages
Dim rstElementPages_cmd
Dim rstElementPages_numRows

Set rstElementPages_cmd = Server.CreateObject ("ADODB.Command")
rstElementPages_cmd.ActiveConnection = MM_OBA_STRING
rstElementPages_cmd.CommandText = "SELECT Pages.PageTitle, Pages.NavigationPage, PageElements.PageID, PageElements.PageElementID, PageElements.ElementID FROM Pages INNER JOIN PageElements ON Pages.PageID = PageElements.PageID WHERE ElementID = ? ORDER BY Pages.PageTitle" 
rstElementPages_cmd.Prepared = true
rstElementPages_cmd.Parameters.Append rstElementPages_cmd.CreateParameter("param1", 5, 1, -1, rstElementPages__lngElementID) ' adDouble

Set rstElementPages = rstElementPages_cmd.Execute
rstElementPages_numRows = 0
%>
<%
Dim rstPositions
Dim rstPositions_cmd
Dim rstPositions_numRows

Set rstPositions_cmd = Server.CreateObject ("ADODB.Command")
rstPositions_cmd.ActiveConnection = MM_OBA_STRING
rstPositions_cmd.CommandText = "SELECT  PositionID, PositionName FROM Positions ORDER BY PositionName" 
rstPositions_cmd.Prepared = true

Set rstPositions = rstPositions_cmd.Execute
rstPositions_numRows = 0
%>
<%
Dim rstGrantLevels
Dim rstGrantLevels_cmd
Dim rstGrantLevels_numRows

Set rstGrantLevels_cmd = Server.CreateObject ("ADODB.Command")
rstGrantLevels_cmd.ActiveConnection = MM_OBA_STRING
rstGrantLevels_cmd.CommandText = "SELECT     GrantLevelID, LevelName FROM         GrantLevels ORDER BY GrantLevelID" 
rstGrantLevels_cmd.Prepared = true

Set rstGrantLevels = rstGrantLevels_cmd.Execute
rstGrantLevels_numRows = 0
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

<h1>Element Information for </h1>

			<table border="0" cellspacing="0" cellpadding="0" class="box">
<%
If bolDeveloperViewGranted Then
	If bolDeveloperEditGranted Then
		strEditLink = "<a href=""ElementEdit.asp?lngElementID=" & (rstElements.Fields.Item("ElementID").Value) & """>"
		strEndEditLink = "</a>&nbsp;"
	Else
		strEditLink = ""
		strEndEditLink = "&nbsp;"
	End If
%>            
              <tr>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	</tr>
              <tr>
              	<td>&nbsp;</td>
                <td align="center"><h1><strong><%=strEditLink & (rstElements.Fields.Item("ElementName").Value) & strEndEditLink%></strong></h1></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	<td>&nbsp;</td>
              	</tr>
			  
              
<%
	If bolSecurityViewGranted Then
%>              
			 </table>
			<table border="0" cellspacing="0" cellpadding="0" class="info">
              <tr>
                <th colspan="3"><h2><a href="Grants.asp?lngElementID=<%=(rstElements.Fields.Item("ElementID").Value)%>">Security</a></h2></th>
              </tr>
<%
		If rstGrants.EOF Then
%>
              <tr>
                <td colspan="3"><a href="Grants.asp?lngElementID=<%=(rstElements.Fields.Item("ElementID").Value)%>">This Element has no Security.</a></td>
              </tr>

<%
		Else
%>               
                  <tr class="column_titles">
                	<td width="3%">&nbsp;</td>
                    <td width="53%"><h4>Position</h4></td>
                    <td width="44%"><h4>Security Level</h4></td>
                  </tr>
                  <tr class="line">
                    <td colspan="3"><hr /></td>
                  </tr>
<%
			Do While Not rstGrants.EOF
				If bolSecurityEditGranted Then
					strEditLink = "<a href=""GrantEdit.asp?lngGrantID=" & (rstGrants.Fields.Item("GrantID").Value) & """>"
					strEndEditLink = "</a>&nbsp;"
				Else
					strEditLink = ""
					strEndEditLink = "&nbsp;"
				End If
%>
                  <tr class="tr_hover">
                	<td><a href="PositionGrants.asp?strPositionTitle=<%=(rstGrants.Fields.Item("POSITION_TITLE").Value)%>" class="row_info"></a></td>
                    <td><%=strEditLink & (rstGrants.Fields.Item("POSITION_TITLE").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstGrants.Fields.Item("LevelName").Value) & strEndEditLink%></td>
                  </tr>
<%
				rstGrants.MoveNext
			Loop
%>                  
<form id="frmAdd" name="frmAdd" method="POST" action="<%=MM_editAction%>">
              <tr>
                <td>&nbsp;</td>
                <td><select name="cbxUserID" id="cbxUserID">
                  <%
While (NOT rstPositions.EOF)
%>
                  <option value="<%=(rstPositions.Fields.Item("MasterPositionID").Value)%>"><%=(rstPositions.Fields.Item("PositionName").Value)%></option>
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
                <td><select name="cbxGrantLevelID" id="cbxGrantLevelID">
                  <%
While (NOT rstGrantLevels.EOF)
%>
                  <option value="<%=(rstGrantLevels.Fields.Item("GrantLevelID").Value)%>"><%=(rstGrantLevels.Fields.Item("LevelName").Value)%></option>
                  <%
  rstGrantLevels.MoveNext()
Wend
If (rstGrantLevels.CursorType > 0) Then
  rstGrantLevels.MoveFirst
Else
  rstGrantLevels.Requery
End If
%>
                </select></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td><input type="submit" name="btnAdd" id="btnAdd" value="Add" />
                <input name="htbxElementID" type="hidden" id="htbxElementID" value="<%=lngElementID%>" /></td>
              </tr>
            <input type="hidden" name="MM_insert" value="frmAdd" />
            </form>                  
                  
<%
		End If
	End If
%>              
            </table>
			<table border="0" cellspacing="0" cellpadding="0" class="info">
              <tr>
                <th colspan="2"><h2>Pages</h2></th>
              </tr>
<%
	If rstElementPages.EOF Then
%>
              <tr>
                <td colspan="2"><a href="PageElements.asp?lngElementID=<%=(rstElements.Fields.Item("ElementID").Value)%>">This Element is not displayed on any Page.</a></td>
              </tr>

<%
	Else
%>               
              <tr>
                <td colspan="2"><table border="0" cellspacing="0" cellpadding="0">
                  <tr class="column_titles">
                	<td width="3%"></td>
                    <td width="45%"><h4>Page</h4></td>
                    <td width="52%"><h4>Navigation</h4></td>
                  </tr>
                  <tr class="line">
                    <td colspan="3"><hr /></td>
                  </tr>
<%
		Do While Not rstElementPages.EOF
			If bolDeveloperEditGranted Then
				strEditLink = "<a href=""PageEdit.asp?lngPageID=" & (rstElementPages.Fields.Item("PageID").Value) & """>"
				strEndEditLink = "</a>&nbsp;"
			Else
				strEditLink = ""
				strEndEditLink = "&nbsp;"
			End If
%>
                  <tr>
                	<td><a href="PageElements.asp?lngPageElementID=<%=(rstElementPages.Fields.Item("PageElementID").Value)%>" class="row_info"></a></td>
                    <td><%=strEditLink & (rstElementPages.Fields.Item("PageTitle").Value) & strEndEditLink%></td>
                    <td><%=strEditLink & (rstElementPages.Fields.Item("NavigationPage").Value) & strEndEditLink%></td>
                  </tr>
<%
			rstElementPages.MoveNext
		Loop
%>                  
                  
                  
                </table></td>
          </tr>
<%
	End If
%>              
              <tr>
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
<%
rstGrants.Close()
Set rstGrants = Nothing
%>
<%
rstElementPages.Close()
Set rstElementPages = Nothing
%>
<%
rstPositions.Close()
Set rstPositions = Nothing
%>
<%
rstGrantLevels.Close()
Set rstGrantLevels = Nothing
%>
