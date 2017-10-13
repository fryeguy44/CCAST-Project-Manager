<!--#include file="Connections/OBA.asp" -->
<div class="filesList">
<h2>Files</h2>

<fieldset>
	<form id="frmFile" name="frmFile" method="POST" enctype="multipart/form-data" action="ProjectDetailFileUploadAction.asp">
		<table>
			<tr>
			  <td align="right"><strong>File title:</strong> </td>
			  <td><INPUT TYPE="text" NAME="title" id="filetitle" pattern=".{1,50}" required title="5 to 10 characters"></td>
			</tr>
			<tr>
			  <td colspan="2"><br /></td>
			</tr>
         
         
			<tr id="fileUpload">
			  <td align="right"><strong>Upload file:</strong> </td>
			  <td><INPUT TYPE="file" NAME="file" id="fileinput" style="display:inline-block;"  > <input type="submit" name="btnEdit" id="btnEdit" value="Upload" /></td>
			</tr>
         
         <tr id="editfiletitle" style="display:none">
			  <td align="right"></td>
			  <td><input type="button" name="updatetitle" id="updatetitle" value="Update" />
           <input type="hidden" name="idno" id="fileid"  />
           </td>
			</tr>
         
		</table>
		<input type="hidden" name="filename" id="filename" value=""  />
		<input type="hidden" name="fileExt" id="fileExt" value=""  />
		<input type="hidden" name="lngProjectDetailID" id="lngProjectDetailID" value="<%=request.querystring("lngProjectDetailID") %>"  />
	</form>
</fieldset>

<ul style="padding: 10px;margin-left: 14px;">
<%
Set MM_editCmd = Server.CreateObject ("ADODB.Command")
MM_editCmd.ActiveConnection = MM_OBA_STRING
MM_editCmd.CommandText = "select * from UploadFiles WHERE 1=1 AND ProjectDetailID = " &  request.querystring("lngProjectDetailID")
SET RS = MM_editCmd.Execute
IF Not RS.EOF THEN
Do While Not RS.EOF
%>
	<li style="list-style-type: inherit;">
   
   
	 <div style="width:200px;float:left;"><a style="text-decration:underline; color:#000;" href="/UploadFiles/<%=(RS.Fields.Item("uploadFileID").Value) & "." & (RS.Fields.Item("UploadFileExtension").Value) %>" target="_blank"><% IF (RS.Fields.Item("title").Value) <> "" THEN Response.Write(RS.Fields.Item("title").Value) ELSE Response.Write(RS.Fields.Item("uploadFileID").Value & "." & RS.Fields.Item("UploadFileExtension").Value) END IF %>
		</a> </div>
       <div style="width:200px;float:left;">
      &nbsp;
      <a class="editfileuploadedtitle" data-id="<% IF (RS.Fields.Item("title").Value) <> "" THEN Response.Write(RS.Fields.Item("title").Value) ELSE Response.Write(RS.Fields.Item("uploadFileID").Value & "." & RS.Fields.Item("UploadFileExtension").Value) END IF %>" id="<%=(RS.Fields.Item("uploadFileID").Value)%>" style="text-decration:underline;color:#000;" target="_blank">Edit</a>
     &nbsp;&nbsp;&nbsp;
      <a class="editDeleteAction" data-id="deleteuploadFile" id="<%=(RS.Fields.Item("uploadFileID").Value)%>" style="text-decration:underline; color:red;" target="_blank">Delete</a>
      </div>
<%
RS.MoveNext
Loop 
Else
%>
No files yet.
<% END IF %>
</ul>


</div>

<div class="filesList" style="display:none;text-align:center;">
	<img src="/img/loading.gif" />
</div>