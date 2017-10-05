<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/OBA.asp" -->

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
Function Nz(o)
  If IsNull(o) Then Nz = 0 Else Nz = o
End Function
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr"><!-- InstanceBegin template="/Templates/master.dwt" codeOutsideHTMLIsLocked="false" -->
<%
If lngAccessTypeID = "" Then
	lngAccessTypeID = 1
End If
%>
<!--#include file="Templates/incMasterSecurity.asp" --> 
<style>
.top_heading h1
{
	font-size: 26px;
    margin: 15px 0 12px;
    text-transform: uppercase;
    text-decoration: none;
    text-align: center;
    color: #091436;
    font-weight: bold;
}
.top_heading th h4
{
	margin: 5px 4px;
	font-size: 10px;
	font-weight: bold;
	
}


</style>
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
              <div align="right" id="ccast">
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
	
	<div id="content"  class="top_heading">

<!-- #BeginEditable "content" -->
<%
If bolVendorOnlyViewGranted Then
	strWhereClause = " AND ProjectDetails.VendorID = " & Session("VendorID")
Else
	If bolClientOnlyViewGranted Then
		strWhereClause = " AND Projects.ClientID = " & Session("ClientID")
	Else
		strWhereClause = ""
	End If
End If
%>
<%
Dim rstProjects
Dim rstProjects_cmd
Dim rstProjects_numRows




Set rstProjects_cmd = Server.CreateObject ("ADODB.Command")
rstProjects_cmd.ActiveConnection = MM_OBA_STRING
rstProjects_cmd.CommandText = "SELECT Projects.ProjectID, Projects.ProjectDescription, Projects.ProjectPriority, ProjectDetails.ProjectDetailID, ProjectDetails.DetailDescription, ProjectDetails.DeveloperNotes, ProjectDetails.OwnerNotes, ProjectDetails.Priority, ProjectDetails.BilledToClient, ProjectDetails.StartTime, ProjectStages.StageName, ProjectStages.SortOrder, ProjectDetails.Hours, Projects.ProjectRate, Clients.ClientName, Vendors.VendorName,  COALESCE(UploadFiles.NumFiles, 0) AS NumFiles FROM Vendors INNER JOIN ProjectStages INNER JOIN ProjectDetails ON ProjectStages.ProjectStageID = ProjectDetails.ProjectStageID ON Vendors.VendorID = ProjectDetails.VendorID LEFT OUTER JOIN (SELECT ProjectDetailID, Count(*) AS NumFiles FROM UploadFiles GROUP BY ProjectDetailID) AS UploadFiles ON ProjectDetails.ProjectDetailID = UploadFiles.ProjectDetailID RIGHT OUTER JOIN Clients INNER JOIN Projects ON Clients.ClientID = Projects.ClientID ON ProjectDetails.ProjectID = Projects.ProjectID WHERE (ProjectStages.StageName <> N'Complete') AND (ProjectStages.StageName <> 'On Hold'" & strWhereClause & ") ORDER BY Clients.ClientName, Projects.ProjectPriority, ProjectDetails.Priority, ProjectStages.SortOrder DESC"


'"SELECT  Projects.ProjectID, Projects.ProjectDescription, Projects.ProjectPriority, ProjectDetails.ProjectDetailID, ProjectDetails.DetailDescription, ProjectDetails.DeveloperNotes, ProjectDetails.OwnerNotes, ProjectDetails.Priority,  ProjectDetails.BilledToClient, ProjectDetails.StartTime, ProjectStages.StageName, ProjectStages.SortOrder, ProjectDetails.Hours, Projects.ProjectRate, Clients.ClientName, Vendors.VendorName FROM Vendors INNER JOIN ProjectStages INNER JOIN ProjectDetails ON ProjectStages.ProjectStageID = ProjectDetails.ProjectStageID ON Vendors.VendorID = ProjectDetails.VendorID RIGHT OUTER JOIN Clients INNER JOIN Projects ON Clients.ClientID = Projects.ClientID ON ProjectDetails.ProjectID = Projects.ProjectID WHERE StageName <> N'Complete' AND StageName <> 'On Hold'" & strWhereClause & " ORDER BY Clients.ClientName, Projects.ProjectPriority, ProjectDetails.Priority, ProjectStages.SortOrder DESC"


rstProjects_cmd.Prepared = true

Set rstProjects = rstProjects_cmd.Execute
rstProjects_numRows = 0
%>
	<h1>Ongoing Projects</h1>	
	<table border="0" cellspacing="0" cellpadding="0" class="fluid">
	  <tr>
	    <th align="left">&nbsp;</th>
	    <th align="left"><h4>Client</h4></th>
	    <th align="left"><h4>Vendor</h4></th>
	    <th align="left"><h4>Description</h4></th>
	    <th align="left"><h4>Detail</h4></th>
	    <th align="left"><h4>Dev Notes</h4></th>
	    <th align="left"><h4>Owner Notes</h4></th>
		<th align="left"><h4>Files</h4></th>
	    <th style="text-align:center;"><h4>Project Priority</h4></th>
	    <th align="center"><h4>Priority</h4></th>
	    <th style="text-align:center;"><h4>Project Stage</h4></th>
	    <th align="left"><h4>&nbsp;</h4></th>
      </tr>
<%
curCurrentCompletionValue = 0

If bolProjectsViewGranted OR bolVendorOnlyViewGranted OR bolClientOnlyViewGranted Then
	Do While Not rstProjects.EOF
		strEdit = ""
		strEditEnd = ""
		strEditDetails = ""
		strEditDetailsEnd = ""
		strWorking = ""
		strVendorName = ""
		
		If bolProjectsEditGranted Then
			strEdit = "<a href=""ProjectEdit.asp?lngProjectID=" & (rstProjects.Fields.Item("ProjectID").Value) & """>"
			strEditDetails = "<a href=""ProjectDetailEdit.asp?lngProjectDetailID=" & (rstProjects.Fields.Item("ProjectDetailID").Value) & """>"
			strEditEnd = "</a>"
			strEditDetailsEnd = "</a>"
			strVendorName = rstProjects.Fields.Item("VendorName").Value
			
			If Not IsNull(rstProjects.Fields.Item("StartTime").Value) Or rstProjects.Fields.Item("StartTime").Value <> "" Then
				strWorking = "W"
			End If
			
		End If 
		
		If bolVendorOnlyViewGranted Then
			strEditDetails = "<a href=""ProjectDetailEditVendor.asp?lngProjectDetailID=" & (rstProjects.Fields.Item("ProjectDetailID").Value) & """>"
			strEditDetailsEnd = "</a>"
			strVendorName = rstProjects.Fields.Item("VendorName").Value
		End If 
		
		If bolClientOnlyViewGranted Then
			strEditDetails = "<a href=""ProjectDetailEditClient.asp?lngProjectDetailID=" & (rstProjects.Fields.Item("ProjectDetailID").Value) & """>"
			strEditDetailsEnd = "</a>"
			strVendorName = "N/A"
		End If 
		
		If IsNull(rstProjects.Fields.Item("DetailDescription").Value) Then
			strDetailDesc = "A current detailed description is not available"
		Else
			strDetailDesc = rstProjects.Fields.Item("DetailDescription").Value
		End If
%>      
	  <tr class="tr_hover">
        <td><a href="ProjectInformation.asp?lngProjectID=<%=(rstProjects.Fields.Item("ProjectID").Value)%>" class="row_info"></a></td>
        <td nowrap="nowrap"><%=(rstProjects.Fields.Item("ClientName").Value)%></td>
        <td nowrap="nowrap"><%=strVendorName%></td>
		<td nowrap="nowrap"><%=strEdit & (rstProjects.Fields.Item("ProjectDescription").Value) & strEditEnd%></td>
		<td nowrap="nowrap"><%=strEditDetails & strDetailDesc & strEditDetailsEnd%></td>
		<td><%=strEditDetails & (rstProjects.Fields.Item("DeveloperNotes").Value) & strEditDetailsEnd%></td>
		<td><%=strEditDetails & (rstProjects.Fields.Item("OwnerNotes").Value) & strEditDetailsEnd%></td>
		<td>
			<a href="javascript:oprenFilesPopUp(<%=(rstProjects.Fields.Item("ProjectDetailID").Value) %>);"><%=(rstProjects.Fields.Item("NumFiles").Value) %></a>
		</td>
		<td align="center"><%=strEdit & (rstProjects.Fields.Item("ProjectPriority").Value) & strEditEnd%></td>
		<td align="center"><%=strEditDetails & (rstProjects.Fields.Item("Priority").Value) & strEditDetailsEnd%></td>
		<td align="center"><%=strEditDetails & (rstProjects.Fields.Item("StageName").Value) & strEditDetailsEnd%></td>
		<td><%=strWorking%></td>
	  </tr>
<%
        If (rstProjects.Fields.Item("StageName").Value) = "Complete" Then
			curCurrentCompletionValue = curCurrentCompletionValue + Nz( (rstProjects.Fields.Item("Hours").Value) * (rstProjects.Fields.Item("ProjectRate").Value))
		End If
		
		rstProjects.MoveNext
    Loop
	If bolDeveloperViewGranted Then
%>
	  <tr>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td align="right">Completion Value:</td>
		<td><%=FormatCurrency(curCurrentCompletionValue)%></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	  </tr>

<%	
	End If 
Else
%>  
        <tr>
            <td colspan="11">Viewing this list requires certain &quot;Projects&quot; permissions</td>
        </tr>

<%
End If
%>
	  <tr>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
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

<script type="text/javascript">
	function oprenFilesPopUp(projecDetailID) {
		$.get("/ProjectDetailFiles.asp", {lngProjectDetailID:projecDetailID}, function(data){
			$("#filesmodal .modal-body").html(data);
			document.getElementById('fileinput').addEventListener('change', readSingleFile, false);
			
			$("form#frmFile").submit(function(){
				var formData = new FormData($(this)[0]);
				console.log(JSON.stringify(formData));
				//console.log(formData);
				$(".filesList").toggle();
				$.ajax({
					  url: $("#frmFile").attr('action'),
					  type: "POST",
					  data: formData,
					  processData: false,
					  contentType: false,
					  success: function (data) {
						$(".filesList").toggle();
						oprenFilesPopUp($("#lngProjectDetailID").val());
					  }
				});
				
				return false;
			});
			$("#filesmodal").modal(true);
		});
	}

	function readSingleFile(evt) {
	
		//Retrieve the first (and only!) File from the FileList object
		var f = evt.target.files[0]; 

		if (f) {
			$('#filename').val(f.name);
			$('#fileExt').val(f.name.split(".").pop());
		}
	  }
	  
	
</script>
<!-- Bootstrap -->
<link href="css/bootstrap.css?v=1.1" rel="stylesheet">

<!-- Custom -->
<link href="style.css" rel="stylesheet">

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) --> 
<script src="js/jquery-1.11.2.min.js"></script>

<!-- Include all compiled plugins (below), or include individual files as needed --> 
<script src="js/bootstrap.js"></script>

		<!--  Modal content for the mixer image example -->
		  <div class="modal fade pop-up-5" id="filesmodal" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel-5" aria-hidden="true">
			<div class="modal-dialog modal-lg">
			  <div class="modal-content">

				<div class="modal-header">
				  <button type="button" class="close" data-dismiss="modal" aria-hidden="true">×</button>
				  <h4 class="modal-title" id="myLargeModalLabel-2"></h4>
				</div>
				<div class="modal-body" style="background:#fff;">
				
				<h1>Test</h1>
				
				</div>
			  </div><!-- /.modal-content -->
			</div><!-- /.modal-dialog -->
		  </div><!-- /.modal mixer image -->
		


</body>

<!-- InstanceEnd --></html>
<%
rstProjects.Close()
Set rstProjects = Nothing
%>
