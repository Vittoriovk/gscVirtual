
<script>
function myConfirm(titolo,info,action)
{
    $("#myConfirmTitolo").html('&nbsp;&nbsp;' + titolo); 
    $("#myConfirmInfo").html(info);
	$("#myConfirmAction").val(action);
	xx=$('#myConfirmModal').modal('toggle');
}
function myConfirmInfo(titolo,info,action,addInfo)
{
    $("#myConfirmTitoloInfo").html('&nbsp;&nbsp;' + titolo); 
    $("#myConfirmInfoInfo").html(info);
	$("#myConfirmAction").val(action);
	xx=$('#myConfirmModalInfo').modal('toggle');
}
function myConfirmYes()
  {
    var act = $("#myConfirmAction").val();
    ImpostaValoreDi("Oper",act);
    document.Fdati.submit();
}
function myConfirmInfoYes()
  {
    var act = $("#myConfirmAction").val();
	var inf = $("#myConfirmAddinfo").val();
    ImpostaValoreDi("Oper",act);
	ImpostaValoreDi("ItemToModify",inf);
    document.Fdati.submit();
}
</script>
</script>
<input type="hidden" name="myConfirmAction"  id="myConfirmAction"  value="">

<div class="modal fade" id="myConfirmModal"  aria-hidden="true" role="dialog">
  <!-- a pieno schermo -->
  <div class="modal-dialog">
    <div class="modal-content">
        <div class="row bg-warning">
			<div class="col-10 "><h3 class="bg-warning " id="myConfirmTitolo"></h3>
			</div>
			<div class="col-2">
			</div>
			<div class="col-1"></div>
		</div>
        <div class="row bg-light">
		   <div class="col-1"></div>
		   <div class="col-11">
		   <h4 id="myConfirmInfo"></h4>  
		   </div>
		</div>
      <div class="row bg-light text-center">
	     <div class="col-6">
         <button type="button" onclick="myConfirmYes();" class="btn btn-success" data-dismiss="modal">Conferma</button>
		 </div>
         <div class="col-6">
         <button type="button" onclick="myConfirmNo();" class="btn btn-danger" data-dismiss="modal">Annulla</button>
		 </div>
      </div>

    </div>
  </div>
</div>

<div class="modal fade" id="myConfirmModalInfo"  aria-hidden="true" role="dialog">
  <!-- a pieno schermo -->
  <div class="modal-dialog">
    <div class="modal-content">
        <div class="row bg-warning">
			<div class="col-10 "><h3 class="bg-warning " id="myConfirmTitoloInfo"></h3>
			</div>
			<div class="col-2">
			</div>
			<div class="col-1"></div>
		</div>
        <div class="row bg-light">
		   <div class="col-1"></div>
		   <div class="col-11">
		   <h4 id="myConfirmInfoInfo"></h4>  
		   </div>
		</div>
		
        <div class="row  bg-light">
		   <div class="col-1"></div>
		   <div class="col-11">
              <div class="form-group ">
			     <%xx=ShowLabel("informazioni aggiuntive")%>
					 <input type="text"  name="myConfirmAddinfo" id="myConfirmAddinfo" class="form-control">
                  </div>		
		   </div>
      </div>  
        <div class="row  bg-light">
		   <div class="col-12"></div>
        </div>
      <div class="row bg-light text-center">
	     <div class="col-6">
         <button type="button" onclick="myConfirmInfoYes();" class="btn btn-success" data-dismiss="modal">Conferma</button>
		 </div>
         <div class="col-6">
         <button type="button" onclick="myConfirmInfoNo();" class="btn btn-danger" data-dismiss="modal">Annulla</button>
		 </div>
      </div>

    </div>
  </div>
</div>