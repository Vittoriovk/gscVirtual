
<script>
function modalGenericaStart(dataIn)
{
   var vp=$("#ModalGenericaVirtualPath").val(); 
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/modalGenerica.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	    $("#ModalGenerica").html(msg); 
		$('#confirmModalGenerica').modal('toggle');
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}

</script>
<%
      localVirtualPath = virtualPath
	  if right(localVirtualPath,1)="\" or right(localVirtualPath,1)="/" then 
	     localVirtualPath=mid(localVirtualPath,1,len(localVirtualPath)-1)
	  end if  
%>
<input type="hidden" name="ModalGenericaVirtualPath" id="ModalGenericaVirtualPath" value="<%=VirtualPath%>">

