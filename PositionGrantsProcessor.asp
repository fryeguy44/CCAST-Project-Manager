<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/OBA.asp" -->
<%
Dim strPositionTitle
Dim strReturnPath

strPositionTitle = Request.Form("htbxPositionTitle")
lngPositionID = Request.Form("htbxPositionID")

strReturnPath = Request.Form("htbxReturnPath")
%>
<%
Dim rstElements__lngPositionID
rstElements__lngPositionID = "1"
If (lngPositionID <> "") Then 
  rstElements__lngPositionID = lngPositionID
End If
%>
<%
Dim rstElements
Dim rstElements_cmd
Dim rstElements_numRows

Set rstElements_cmd = Server.CreateObject ("ADODB.Command")
rstElements_cmd.ActiveConnection = MM_OBA_STRING
rstElements_cmd.CommandText = "SELECT Elements.ElementName, GrantLevels.LevelName, ISNULL(GrantLevels.GrantLevelID, 0) AS GrantLevelID, Elements.ElementID FROM (SELECT  GrantLevelID, ElementID FROM Grants WHERE (Grants.PositionID = ?)) AS A INNER JOIN GrantLevels ON A.GrantLevelID = GrantLevels.GrantLevelID RIGHT OUTER JOIN Elements ON A.ElementID = Elements.ElementID ORDER BY Elements.ElementName" 
rstElements_cmd.Prepared = true
rstElements_cmd.Parameters.Append rstElements_cmd.CreateParameter("param1", 202, 1, 255, rstElements__lngPositionID) ' adDouble

Set rstElements = rstElements_cmd.Execute
rstElements_numRows = 0
%>
<%
    Set DeleteCMD = Server.CreateObject ("ADODB.Command")
    DeleteCMD.ActiveConnection = MM_OBA_STRING
	Do While Not rstElements.EOF
		Response.Write("cbxGrantLevelID = " & Request.Form("cbxGrantLevelID"  & (rstElements.Fields.Item("ElementID").Value)))
		Response.Write("<br />ElementGrantLevelID = " & rstElements.Fields.Item("GrantLevelID").Value)
		Response.Write("<br />strPositionTitle = " & strPositionTitle)
		Response.Write("<br />ElementID = " & (rstElements.Fields.Item("ElementID").Value))
		Response.Write("<br />strReturnPath = " & strReturnPath)
		
		If CInt(Request.Form("cbxGrantLevelID"  & (rstElements.Fields.Item("ElementID").Value))) <> CInt(rstElements.Fields.Item("GrantLevelID").Value) Then
			If Request.Form("cbxGrantLevelID"  & (rstElements.Fields.Item("ElementID").Value)) = "0" Then
				DeleteCMD.CommandText = "DELETE FROM Grants WHERE (ElementID = " & (rstElements.Fields.Item("ElementID").Value) & ") AND (PositionID = '" & lngPositionID & "')"
				DeleteCMD.Execute
				
			Else
				If (rstElements.Fields.Item("GrantLevelID").Value) = 0 Then
					DeleteCMD.CommandText = "INSERT INTO Grants (ElementID, PositionID, GrantLevelID) VALUES (" & (rstElements.Fields.Item("ElementID").Value) & ", '" & lngPositionID & "', " & Request.Form("cbxGrantLevelID"  & (rstElements.Fields.Item("ElementID").Value)) & ")"
					DeleteCMD.Execute
					
				Else
						DeleteCMD.CommandText = "UPDATE Grants SET GrantLevelID = " & Request.Form("cbxGrantLevelID"  & (rstElements.Fields.Item("ElementID").Value)) & " WHERE (ElementID = " & (rstElements.Fields.Item("ElementID").Value) & ") AND (PositionID = '" & lngPositionID & "')"
						DeleteCMD.Execute
		
						
				End If
			End If
		End If
	
	
		rstElements.MoveNext
	Loop
    DeleteCMD.ActiveConnection.Close
%>
<%
	Response.Redirect(strReturnPath)
%>
<%
rstElements.Close()
Set rstElements = Nothing
%>
