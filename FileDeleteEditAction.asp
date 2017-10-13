<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->
<%
     dim actionCase , ID ,titlename
	  actionCase = request.querystring("Action")
	  ID = request.querystring("uniqueID")
	  titlename = request.querystring("titlename")
	  
	  
	  
	  
	 ' response.end
	  
   Select Case actionCase
       Case "deleteuploadFile"
          deleteuploadFile(ID)
	    Case "edituploadedFileTitle"
			Call edituploadedFileTitle(ID,titlename)
   End Select
		
  function edituploadedFileTitle(ID,titlename)
		
		 Set MM_editCmd = Server.CreateObject ("ADODB.Command")
		 MM_editCmd.ActiveConnection = MM_OBA_STRING  
		 MM_editCmd.CommandText = "update UploadFiles set title = '"&titlename& "' where UploadFileID ="&ID
				 MM_editCmd.Execute
				' MM_editCmd.ActiveConnection.Close
  end function
				 
				 
		function deleteuploadFile(ID)
		  
			   Set MM_editCmd = Server.CreateObject ("ADODB.Command")
				MM_editCmd.ActiveConnection = MM_OBA_STRING
								
				MM_editCmd.CommandText = "select * from UploadFiles where UploadFileID ="&ID
				SET RS = MM_editCmd.Execute
				fileExtension = RS.Fields.Item("UploadFileExtension").value
				DestinationPath = Server.MapPath("/UploadFiles")				
		      Set obj = CreateObject("Scripting.FileSystemObject") 'Calls the File System Object
            obj.DeleteFile(DestinationPath&"\"&ID&"."&fileExtension) 'Deletes the file throught the DeleteFile function
			   MM_editCmd.CommandText = "delete from UploadFiles where UploadFileID ="&ID
			   MM_editCmd.Execute
			   MM_editCmd.ActiveConnection.Close 
				
				
	  end function
 %>