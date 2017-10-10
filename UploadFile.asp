<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
  Dim GotFile, FileType, ProjectDetailID, FileError
  ProjectDetailID = request.querystring("ProjectDetailID")
  GotFile = request.querystring("GotFile")
  FileType = request.querystring("FileType")
  FileError = False
  
  
  
  
  Dim FileDescription, Path
  
  If FileType = "devNotes" Then
    ImageDescription = "Developer Notes File"
  ElseIf FileType = "ownerNotes" Then
    ImageDescription = "Owner Notes File"
  End If
  
  If GotFile <> "False" Then
    Set Upload = Server.CreateObject("Persits.Upload")
    
    Upload.Save
    ProjectDetailID = Upload.form("ProjectDetailID")
    FileType = Upload.form("FileType")
    Path = server.mappath("/UploadFiles/")
    
    
    
    
    Dim File, Extension
    Set File = Upload.Files(1)
    
    If lcase(right(File.Filename, 4)) = ".jpg" OR lcase(right(File.Filename, 4)) = ".bmp" OR lcase(right(File.Filename, 4)) = ".gif" OR lcase(right(File.Filename, 4)) = ".tif" OR lcase(right(File.Filename, 4)) = ".png" OR lcase(right(File.Filename, 4)) = ".doc" OR lcase(right(File.Filename, 4)) = ".pdf" OR lcase(right(File.Filename, 4)) = ".xls" Then
      Extension = lcase(right(File.Filename, 4))
    ElseIf lcase(right(File.Filename, 5)) = ".jpeg" Then
      Extension = lcase(right(File.Filename, 5))
    ElseIf lcase(right(File.Filename, 5)) = ".docx" Then
      Extension = lcase(right(File.Filename, 5))
    ElseIf lcase(right(File.Filename, 5)) = ".xlsx" Then
      Extension = lcase(right(File.Filename, 5))
    End If
    
    
    
    
    'error checking
    Dim ErrOut
    
    If FileType = "devNotes" or FileType = "ownerNotes" Then
      ErrOut = True
      
      If Extension = ".jpg" Then
        ErrOut = False
      ElseIf Extension = ".jpeg" Then
        ErrOut = False
      ElseIf Extension = ".bmp" Then
        ErrOut = False
      ElseIf extension = ".gif" Then
        ErrOut = False
      ElseIf extension = ".tif" Then
        errOut = False
      ElseIf extension = ".png" Then
        ErrOut = False
      ElseIf extension = ".doc" Then
        ErrOut = False
      ElseIf extension = ".docx" Then
        ErrOut = False
      ElseIf extension = ".xls" Then
        ErrOut = False
      ElseIf extension = ".xlsx" Then
        ErrOut = False
      ElseIf extension = ".pdf" Then
        ErrOut = False
      End If
    Else
      ErrOut = False
    End If
    
    
    
    
    dim MyDevNotesFileName, MyOwnerNotesFileName
    
    If Not ErrOut Then
      If FileType = "devNotes" Then
        MyDevNotesFileName = ProjectDetailID & "-" & FileType & Extension
        'Response.Write Path & "\" & MyDevNotesFileName
        File.SaveAs Path & "\" & MyDevNotesFileName
      Else
        MyOwnerNotesFileName = ProjectDetailID & "-" & FileType & Extension
        'Response.Write Path & "\" & MyOwnerNotesFileName
        File.SaveAs Path & "\" & MyOwnerNotesFileName
      End If
      
      
      
      
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
      
      If FileType = "devNotes" Then
        uploadFileCmd.Parameters.Append uploadFileCmd.CreateParameter("param1", 202, 1, 50, MyDevNotesFileName) ' adVarWChar
      Else
        uploadFileCmd.Parameters.Append uploadFileCmd.CreateParameter("param1", 202, 1, 50, MyOwnerNotesFileName) ' adVarWChar
      End If
      
      uploadFileCmd.Parameters.Append uploadFileCmd.CreateParameter("param2", 5, 1, -1, ProjectDetailID) ' adDouble
      uploadFileCmd.Execute
      uploadFileCmd.ActiveConnection.Close
%>
<script type="text/javascript">
  opener.location.reload();
  opener.focus();
  self.close();
</script>
<%  Else %>
<script type="text/javascript">
  self.alert("You must select file of type .jpg, .jpeg, .gif, .bmp, .tif, .doc, docx, .xls, .xlsx, .pdf or .png!");
  opener.location.reload();
  opener.focus();
  self.close();
</script>
<%  End If
  Else %>
<script type="text/javascript">
  function verify(frm, bttn) {
    var errOut = false;
    
    if (frm.MyFile.value == '') {
      self.alert("Please choose a file to upload!");
      errOut = true;
    }
    
    if (!errOut) {
      bttn.disabled = true;
      frm.submit();
    }
  }
</script>
<p class="title">Upload <%=FileDescription%></p>

<form action="UploadFile.asp" method="POST" enctype="multipart/form-data">
  <input type="file" name="MyFile" size="30" /><br />
  <br />
  
  <div style="text-align: center;">
    <input type="Button" value="Upload" onClick="verify(this.form, this);" />
  </div>
  
  <input type="hidden" name="GotFile" value="True" />
  <input type="hidden" name="ProjectDetailID" value="<%=ProjectDetailID%>" />
  <input type="hidden" name="FileType" value="<%=FileType%>" />
</form>
<%  End If %>