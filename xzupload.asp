<!--#include file="clsUpload.asp" -->
<%
Dim objUpload 
Dim strFile, strPath
' Instantiate Upload Class '
Set objUpload = New clsUpload
strFile = objUpload.Fields("file").FileName
strPath = server.mappath("/UploadFiles") & "/" & strFile
' Save the binary data to the file system '
objUpload("file").SaveAs strPath
Set objUpload = Nothing

%>