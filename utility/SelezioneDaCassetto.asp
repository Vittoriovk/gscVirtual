<script>
function popolaCassetto(IdAccount,IdDocumento,tipoRife,idRife)
{
   var dataIn="IdAccount=" + IdAccount + "&IdDocumento=" + IdDocumento + "&tipoRife=" + tipoRife + "&idRife=" + idRife;
   var vp=$("#hiddenVirtualPath").val(); 
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/PopolaCassetto.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	    $("#CassettoBodyDett").html(msg); 
		$('#CassettoModal').modal('toggle');
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}

function localCassettoSel()
{
    var s=$('#cassettopieno').val();
	if  (s=='0')
	{
	     $('#btnCassettoClose').click();
		 return true;
    } 
	var s=$('input[name="CassettoCampoNome"]:checked').val();
	$('#btnCassettoClose').click();
	callerCassettoSel(s);
}
</script>

<input type="hidden" name="CassettoCampoDest" id="CassettoCampoDest" value=""> 
<div class="modal fade" id="CassettoModal"  aria-hidden="true" role="dialog">
  <div class="modal-dialog modal-lg">
    <div class="modal-content">
      <div class="modal-header">

        <h2>Selezionare documento Cassetto</h2> 
        <button type="button" class="close" id="btnCassettoClose" data-dismiss="modal">
          <span aria-hidden="true">×</span><span class="sr-only">Chiudi</span>
        </button>
      </div>
      <div class="modal-body"> 
		<div id="CassettoBodyDett">	  
		</div>		  
      </div> 
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
        <button type="button" class="btn btn-primary" onclick="localCassettoSel();";>Seleziona</button>
      </div>
    </div>
  </div>
</div>
