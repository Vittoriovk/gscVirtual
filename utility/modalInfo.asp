
<script>
function myAlert(titolo,info)
{
    $("#myAlertTitolo").html('&nbsp;&nbsp;' + titolo); 
    $("#myAlertInfo").html(info);
	xx=$('#myAlertModal').modal('toggle');
}

</script>

<div class="modal fade" id="myAlertModal"  aria-hidden="true" role="dialog">
  <!-- a pieno schermo -->
  <div class="modal-dialog">
    <div class="modal-content">
        <div class="row bg-warning">
			<div class="col-10 "><h3 class="bg-warning " id="myAlertTitolo"></h3>
			</div>
			<div class="col-2">
                 <button type="button" Id="dismissModalGenerica" class="close" data-dismiss="modal">
                 <span aria-hidden="true">×&nbsp;&nbsp;</span><span class="sr-only">Chiudi</span>
                 </button>
			</div>
			<div class="col-1"></div>
		</div>
        <div class="row bg-light">
		   <div class="col-1"></div>
		   <div class="col-11">
		   <h4 id="myAlertInfo"></h4>  
		   </div>
		</div>

      <div class="row bg-light text-center">
         <button type="button" class="btn btn-default" data-dismiss="modal">Chiudi</button>
      </div>

    </div>
  </div>
</div>