<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
  Dim FSO, ProjectDetailID, FileType, MyFileName
  ProjectDetailID = request.querystring("ProjectDetailID")
  FileType = request.querystring("FileType")
  
  
  
  
  'GIF
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".gif"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  
  'JPG
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".jpg"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  
  'JPEG
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".jpeg"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  
  'BMP
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".bmp"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  
  'TIF
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".tif"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  
  'PNG
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".png"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  
  'DOC
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".doc"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  
  'DOCX
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".docx"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  'XLSX
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".xls"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  'XLSX
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".xlsx"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  'PDF
  MyFileName = Server.MapPath("/UploadFiles/") & "/" & ProjectDetailID & "-" & FileType & ".pdf"
  Set FSO = Server.Createobject("Scripting.FileSystemObject")
  If FSO.fileExists(MyFileName) Then Call FSO.deletefile(MyFileName)
  
  ' Execute The Update
  Dim uploadFileCmd
  
  Set uploadFileCmd = Server.CreateObject ("ADODB.Command")
  uploadFileCmd.ActiveConnection = MM_OBA_STRING
  
  If FileType = "devNotes" Then
    uploadFileCmd.CommandText = "UPDATE dbo.ProjectDetails SET DeveloperNotesFile = ? WHERE ProjectDetailID = ?" 
  Else
    uploadFileCmd.CommandText = "UPDATE dbo.ProjectDetails SET OwnerNotesFile = ? WHERE ProjectDetailID = ?" 
  End If
  
  uploadFileCmd.Prepared = true
  uploadFileCmd.Parameters.Append uploadFileCmd.CreateParameter("param1", 202, 1, 50, "") ' adVarWChar
  uploadFileCmd.Parameters.Append uploadFileCmd.CreateParameter("param2", 5, 1, -1, ProjectDetailID) ' adDouble
  uploadFileCmd.Execute
  uploadFileCmd.ActiveConnection.Close
%>

<script type="text/javascript">
  opener.location.reload();
  opener.focus();
  self.close();
</script>