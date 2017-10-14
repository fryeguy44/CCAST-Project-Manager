$(document).ready(function(e) 
 {
  $("body").delegate(".editDeleteAction","click", function()
   {
		   var uniqueID = $(this).attr('id'); 
			var performAction = $(this).data('id');
			
	  if(performAction == "deleteuploadFile")
		{
		    if(confirm("Are you sure to delete this file") == false)
			 {
				 return false;
			 }
	
			 $.ajax({
					  url: 'FileDeleteEditAction.asp',
					  type: "get",
					  data: {'Action':performAction,'uniqueID':uniqueID},
					  success: function (data) {
						$(".filesList").toggle(); 
						oprenFilesPopUp($("#lngProjectDetailID").val());
						//$(".filesList").toggle();
						//oprenFilesPopUp($("#lngProjectDetailID").val());
					  }
				}); 
				
		}	
		 	
	return false;
   });
	
 $("body").delegate(".editfileuploadedtitle","click", function()
   {
      titlename = $(this).data('id');	
		id = $(this).attr('id');
		$('#filetitle').val(titlename);
		$('#fileid').val(id);
      $('#fileUpload').hide();
		$('#editfiletitle').show();
			
	});
	
$("body").delegate("#updatetitle","click", function()
   {
        var titlename = $('#filetitle').val();
		  var uniqueID = $('#fileid').val();		  
		   $.ajax({
					  url: 'FileDeleteEditAction.asp',
					  type: "get",
					  data: {'Action':'edituploadedFileTitle','uniqueID':uniqueID,'titlename':titlename},
					  success: function (data) {
						$(".filesList").toggle(); 
						  oprenFilesPopUp($("#lngProjectDetailID").val());
					  }
				}); 
 	   });
  });
function openCommonFilesPopUp(projecDetailID)
{	
$.get("/ProjectDetailEditVendorAjaxAction.asp", {lngProjectDetailID:projecDetailID}, function(data){
		
		   //alert($("#frmEdit").attr('action'));
		    $("#filesmodal .modal-body").html(data);
			//document.getElementById('fileinput').addEventListener('change', readSingleFile, false);
			$("form#frmEdit_action").submit(function(){
				var formData = new FormData($(this)[0]);
				console.log(JSON.stringify(formData));
				//console.log(formData);
				$(".filesList").toggle(); 
				$.ajax({
					  url:$("form#frmEdit_action").attr('action'),
					  type: "POST",
					  data: formData,
					  processData: false,
					  contentType: false,
					  success: function (data) {
						//  alert(data);
						$(".filesList").toggle(); 
						openCommonFilesPopUp(projecDetailID);
						
					  }
				});
				return false;
			});
			$("#filesmodal").modal(true);
		});
	}