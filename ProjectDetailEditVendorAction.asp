<!--#include file="Connections/OBA.asp" -->
<!--#include file="_upload.asp" -->

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
Dim DestinationPath, DestinationPath_temp, id
DestinationPath = Server.MapPath("/UploadFiles")

Dim Form: Set Form = New ASPForm 
	
Form.State
	
If (CStr(Form("MM_update")) = "frmEdit") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd
	
	If Form.State = 0 AND Form("fileExt") <> "" Then 'Completted
		Set MM_editCmd = Server.CreateObject ("ADODB.Command")
		MM_editCmd.ActiveConnection = MM_OBA_STRING
		
		
		'MM_editCmd.CommandText = "select count(*)  as recordsCount from UploadFiles WHERE ProjectDetailID = " &  request.querystring("lngProjectDetailID")
		'SET RS = MM_editCmd.Execute
		
		'IF RS("recordsCount") = 0 Then
			MM_editCmd.CommandText = "INSERT INTO UploadFiles (ProjectDetailID ,UploadFileExtension) VALUES (? ,?);" 
		'Else	
			'MM_editCmd.CommandText = "UPDATE UploadFiles SET ProjectDetailID = ?,UploadFileExtension = ? WHERE ProjectDetailID = " &  request.querystring("lngProjectDetailID")
		'End IF
		
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, request.querystring("lngProjectDetailID")) ' adDouble
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 1000, Form("fileExt")) ' adDouble
		MM_editCmd.Execute
		
		
		MM_editCmd.CommandText = "select top 1 uploadFileID from UploadFiles order by uploadFileID desc" 
		SET RS = MM_editCmd.Execute
		
		lastInsertedId = RS("uploadFileID")
		
		MM_editCmd.ActiveConnection.Close
		
		
		Dim DestFileName
		DestFileName = lastInsertedId & "." & Form("fileExt")
		Form.Files.Save DestinationPath,DestFileName
	End If
	
	Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_OBA_STRING
    MM_editCmd.CommandText = "UPDATE dbo.ProjectDetails SET ProjectStageID = ?, DeveloperNotes = ? WHERE ProjectDetailID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Form("cbxProjectStageID"), Form("cbxProjectStageID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 1000, Form("tbxDeveloperNotes")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Form("MM_recordId"), Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
	
	response.redirect("/ProjectDetailEditVendor.asp?lngProjectDetailID=" & request.querystring("lngProjectDetailID"))
	
  End If
End If

%>