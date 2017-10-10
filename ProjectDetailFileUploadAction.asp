<!--#include file="Connections/OBA.asp" -->
<!--#include file="_upload.asp" -->


<%
Dim DestinationPath, DestinationPath_temp, id
DestinationPath = Server.MapPath("/UploadFiles")

Dim Form: Set Form = New ASPForm 
	
response.write(Form.State)
	
    ' execute the update
    Dim MM_editCmd
	
	If Form.State = 0 AND Form("fileExt") <> "" Then 'Completted
	
		Set MM_editCmd = Server.CreateObject ("ADODB.Command")
		MM_editCmd.ActiveConnection = MM_OBA_STRING
		
		MM_editCmd.CommandText = "INSERT INTO UploadFiles (ProjectDetailID ,UploadFileExtension, title) VALUES (? ,?, ?);" 
		
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Form("lngProjectDetailID")) ' adDouble
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 1000, Form("fileExt")) ' adDouble
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 1000, Form("title")) ' adDouble
		MM_editCmd.Execute
		
		
		MM_editCmd.CommandText = "select top 1 uploadFileID from UploadFiles order by uploadFileID desc" 
		SET RS = MM_editCmd.Execute
		
		lastInsertedId = RS("uploadFileID")
		
		MM_editCmd.ActiveConnection.Close
		
		

		DestFileName = lastInsertedId & "." & Form("fileExt")
		Form.Files.Save DestinationPath,DestFileName
	End If
	
	
%>