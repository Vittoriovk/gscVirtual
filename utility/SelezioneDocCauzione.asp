<script>
function popolaDocCauzione(IdTipoCauzione,IdCauzione)
{
   var dataIn="IdTipoCauzione=" + IdTipoCauzione + "&IdCauzione=" + IdCauzione;
   var vp=$("#localVirtualPath").val(); 
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/PopolaDocCauzione.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	    $("#CauzioneBodyDett").html(msg); 
		$('#CauzioneModal').modal('toggle');
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}

</script>

<input type="hidden" name="CauzioneCampoDest" id="CauzioneCampoDest" value=""> 
<div class="modal fade" id="CauzioneModal"  aria-hidden="true" role="dialog">
  <div class="modal-dialog modal-lg">
    <div class="modal-content">
      <div class="modal-header">

        <h2>documenti Cauzione</h2> 
        <button type="button" class="close" id="btnCauzioneClose" data-dismiss="modal">
          <span aria-hidden="true">×</span><span class="sr-only">Chiudi</span>
        </button>
      </div>
      <div class="modal-body"> 
		<div id="CauzioneBodyDett">	  
		</div>		  
      </div> 
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
      </div>
    </div>
  </div>
</div>
