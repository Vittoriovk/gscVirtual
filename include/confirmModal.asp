<script>
function attivaFormModal(motivo,oper,param1,param2)
{
    xx=$('#titoloConfirm').html(motivo);
	xx=$('#operConfirmModal').val(oper);
	xx=$('#operConfirmParm1').val(param1);
	xx=$('#operConfirmParm2').val(param2);
	xx=$('#confirmModal').modal('toggle');
}
function submitFormModal()
{
    yy=$('#operConfirmModal').val();
	xx=ImpostaValoreDi("Oper",yy);
	document.Fdati.submit();
}
</script>

<div class="modal fade" id="confirmModal"  aria-hidden="true" role="dialog">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">
        <input type="hidden" name="operConfirmModal" id="operConfirmModal" value="">
		<input type="hidden" name="operConfirmParm1" id="operConfirmParm1" value="">
		<input type="hidden" name="operConfirmParm2" id="operConfirmParm2" value="">
        <h2 Id="titoloConfirm">TitoloConfirm</h2>
        <button type="button" class="close" data-dismiss="modal">
          <span aria-hidden="true">×</span><span class="sr-only">Chiudi</span>
        </button>
      </div>
      <div class="modal-header">
		<h4>Ci indichi il motivo della sua scelta</h4>
      </div>
  
      <div class="modal-body"> 
		  <div class="form-group">
		  <textarea class="form-control" name="motivoModal" id="motivoModal"></textarea>
		  </div>
      </div> 

      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
        <button type="button" class="btn btn-primary" onclick="submitFormModal();";>Conferma</button>
      </div>
    </div>
  </div>
</div>