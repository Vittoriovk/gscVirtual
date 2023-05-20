	<%
	err.clear 
	'si aspetta in input il nome della struttura 
	if NomeStruttura="" then 
	   NomeStruttura="SEDE_GENERICA"
	end if 
	if DescStruttura="" then 
	   DescStruttura="SEDE"
	end if 
	
	%>
	
  <a class="btn btn-info" data-toggle="collapse" href="#collapse<%=NomeStruttura%>" role="button" 
     onclick="LoadSedi('<%=NomeStruttura%>')"
     aria-expanded="false" aria-controls="collapse<%=NomeStruttura%>">
	 <span Id="<%=NomeStruttura%>_Plus"><i class="fa fa-1x fa-plus-circle"></i></span>
	 <span Id="<%=NomeStruttura%>_Minus" style= "display:none"><i class="fa fa-1x fa-minus-circle"></i></span>
	 <input type="hidden" id="<%=NomeStruttura%>_plusMinus" value = "+">
	 </a>
	 <B> <%=DescStruttura%> </B>
  </p> 
  
	<div class="row">
	  <div class="col">
		<div class="collapse" id="collapse<%=NomeStruttura%>">
			<div class="table-responsive" id="div_<%=NomeStruttura%>"></div>
		</div>
	  </div>
	</div>
	
<%
campiStruttura=""
campiStruttura=campiStruttura + "IdAccountSede,,;"
campiStruttura=campiStruttura + "IdTipoSede,,;"
campiStruttura=campiStruttura + "DescTipoSede,Tipo Sede,;"
campiStruttura=campiStruttura + "IdStato,,IT;"
campiStruttura=campiStruttura + "DescStato,Stato,;"
campiStruttura=campiStruttura + "Indirizzo,Indirizzo,;"
campiStruttura=campiStruttura + "Civico,,;"
campiStruttura=campiStruttura + "Cap,Cap,;"
campiStruttura=campiStruttura + "Comune,Comune,;"
campiStruttura=campiStruttura + "Provincia,Provincia,;"
campiStruttura=campiStruttura + "ProvinciaIT,,;"
campiStruttura=campiStruttura + "Azioni,Azioni,;"
%>
<input type="hidden" name="SEDE_NOME_STRUTTURA" id="SEDE_NOME_STRUTTURA"   value="<%=NomeStruttura%>">

<input type="hidden" name="<%=NomeStruttura%>_Entita"    id="<%=NomeStruttura%>_Entita"   value="SEDE">
<input type="hidden" name="<%=NomeStruttura%>_Account"   id="<%=NomeStruttura%>_Account"  value="<%=IdAccount%>">

<input type="hidden" name="<%=NomeStruttura%>_Header"    id="<%=NomeStruttura%>_Header"   value="<%=campiStruttura%>">
 <!-- prefisso dei campi del form  -->
<input type="hidden" name="<%=NomeStruttura%>_Prefix"    id="<%=NomeStruttura%>_Prefix"   value="Sede_">
 <!-- postfix dei campi del form  -->
<input type="hidden" name="<%=NomeStruttura%>_Postfix"   id="<%=NomeStruttura%>_Postfix"  value="0">

<input type="hidden" name="<%=NomeStruttura%>_Oper"      id="<%=NomeStruttura%>_Oper"     value="<%=flagOperStruttura%>">
<input type="hidden" name="<%=NomeStruttura%>_OperCall"  id="<%=NomeStruttura%>_OperCall" value="attivaFormAddress">


<!--#include virtual="/gscVirtual/configurazioni/sedi/FormSede.asp"--> 

<script>
function LoadSedi(NomeStruttura)
{
   evaluateMinusPlus(NomeStruttura);
   var act=$("#" + NomeStruttura + "_plusMinus").val();
   	if (act=='+') {
		return true;
	}
   var dataIn="sendData=" + $("#sendDataForUpd").val(); 
   OperAmmesse   = $("#" + NomeStruttura + "_Oper").val();
   campiStruttura= $("#" + NomeStruttura + "_Header").val();
   IdAccount     = $("#" + NomeStruttura + "_Account").val();   
   dataIn = dataIn + "&OperAmmesse=" + OperAmmesse;
   dataIn = dataIn + "&NomeStruttura=" + NomeStruttura;
   dataIn = dataIn + "&campiStruttura=" + campiStruttura;
   dataIn = dataIn + "&IdAccount=" + IdAccount;
   //prompt("Copy to clipboard: Ctrl+C, Enter", dataIn);
   var vp=$("#localVirtualPath").val();   
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/configurazioni/sedi/SediListaForm.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	    $("#div_" + NomeStruttura).html(msg); 
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}
</script>
<script>
// idx e' l'indice della riga 
function attivaFormSede(tabName,idx)
{
	$("#" + tabName + "_Idx0").val(idx);
	xx=resetFormSede(tabName);
	xx=$('#' + tabName + 'Modal').modal('toggle');
}

function resetFormSede(tabName)
{
	// legge la struttura e carica i dati sul form
	populateFormByRow(tabName);

	var idx = $("#" + tabName + "_Idx0").val();
	// abilito pulsante cancellazione se è una riga valida
	if (idx==0)
		$('#divformSedeDelete').hide();
	else
		$('#divformSedeDelete').show();
	// qui effettuo solo operazioni ad hoc 
	localSedeChangeStato(0);
	/* imposto la provincia se stato = IT */
	var idS = $("#Sede_IdStato0").val();
    
	if (idS=="IT") {
		var pro = $("#Sede_Provincia0").val();
		var arr = Array.from(document.querySelector("#Sede_ProvinciaIT0").options);
		var val = "";

		arr.forEach(function(option_element) {
			var option_text  = option_element.text;
			var option_value = option_element.value;	
			if (option_text == pro)
				val = option_value;
		});
		if (val.length>0)
			$("#Sede_ProvinciaIT0").val(val);
	}
}

</script>

<script>
function localSedeChangeStato(id)
{
	var v,d;
	v = $("#Sede_IdStato" + id).val();
	d = "#divProv";
	if (v=="IT") {
	   $(d + "IT" + id).show();	
       $(d + id).hide();   
	}
	else {
	   $(d + "IT" + id).hide();	
       $(d + id).show();
 	}
	/* recupero descrizione della lista e la metto nel campo desiderato*/
	v = $("#Sede_IdStato" + id+" option:selected").text();
	$("#Sede_DescStato" + id).val(v);
	

}
function localSedeChangeTipoSede(id)
{
	var v;
	/* recupero descrizione della lista e la metto nel campo desiderato*/
	v = $("#Sede_IdTipoSede" + id + " option:selected").text();
	$("#Sede_DescTipoSede" + id).val(v);
}
function localSedeChangeProvincia(id)
{
	var v;
	v = $("#Sede_ProvinciaIT" + id + " option:selected").text();
	$("#Sede_Provincia" + id).val(v); 
	
	// cambiato lo stato quindi cambiato datalist 
	var idS=$("#Sede_ProvinciaIT" + id).val();
	var dataList = $('#Sede_Comune0').attr('list');
	if (!(dataList=="dataList_ComuneSede" + idS)) {
	
	   // data list inesistente lo creo 
	   if ($('#dataList_ComuneSede' + idS).length==0) 
	   {
	       setDataListInHtml("COMUNE_BYSIGLAPROV_IT","dataList_ComuneSede" + idS,idS,"dataList_ComuneSede" + idS);
	   }
	   $('#Sede_Comune0').attr('list','dataList_ComuneSede' + idS);	
	}
}

function localSedeSubmit(tabName)
{
var xx,yy,oldN,oldD,vS;
    var id=0;
    xx=false;
	yy=ImpostaColoreFocus("Sede_ProvinciaIT" + id,"","white");
	oldN=ValoreDi("NameLoaded");
	oldD=ValoreDi("DescLoaded");
	xx=ImpostaValoreDi("NameLoaded","Sede_IdTipoSede,TE;Sede_Indirizzo,TE;Sede_Cap,TE;Sede_Comune,TE;Sede_Provincia,TE");
	xx=ImpostaValoreDi("DescLoaded",id);
    xx=ElaboraControlli();
	yy=ImpostaValoreDi("NameLoaded",oldN);
	yy=ImpostaValoreDi("DescLoaded",oldD);
 	if (xx==false) {
	   yy=ControllaCampo("Sede_ProvinciaIT" + id,"TE");
	   return false;
	} 
	
	if (ValoreDi("Sede_IdStato0")=="IT"){
	   xx=ControllaCampo("Sede_ProvinciaIT" + id,"TE");
	} 
	
	if (xx==true) {
		xx=localVerificaSede(id);
		if (xx==false)
			bootbox.alert("Sede esistente in archivio");
	}
 	if (xx==false) 
	   return false;
	
	/* recupero l'indice : se = 0 è nuovo altrimenti è una modifica */
	var idx = $("#" + tabName + "_Idx0").val();
	/* se nuovo metto azione in NEW */
	if (idx==0)
		yy=ImpostaValoreDi("Sede_Azioni0","NEW");
	else 
		yy=ImpostaValoreDi("Sede_Azioni0","MOD");
	var riga = createStructureRow(tabName);
	$("#btnS_" + tabName).prop("disabled", true);
	updateSede(tabName,riga);
	$('#' + tabName + 'Modal').modal('hide');
	LoadSedi(tabName);
	LoadSedi(tabName);
	$("#btnS_" + tabName).prop("disabled", false);
	
}

function localSedeRemove(tabName)
{
	yy=ImpostaValoreDi("Sede_Azioni0","DEL");
	var riga = createStructureRow(tabName);
	updateSede(tabName,riga);
	$('#' + tabName + 'Modal').modal('hide');
	LoadSedi(tabName);
	LoadSedi(tabName);
}

function updateSede(NomeStruttura,riga)
{
   var dataIn="sendData=" + $("#SendDataForCall").val(); 
   campiStruttura= $("#" + NomeStruttura + "_Header").val();
   IdAccount     = $("#" + NomeStruttura + "_Account").val();   
   dataIn = dataIn + "&NomeStruttura=" + NomeStruttura;
   dataIn = dataIn + "&campiStruttura=" + campiStruttura;
   dataIn = dataIn + "&rowData=" + riga;
   dataIn = dataIn + "&IdAccount=" + IdAccount;
   dataIn = encodeURI(dataIn);

   var vp=$("#localVirtualPath").val();  
   $.ajax({
      type: "POST",
      async: false,
      url: vp + "/configurazioni/Sedi/SediUpdate.asp",
      data: dataIn,
      success: function(msg)
      {
	    return true;
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}

function localVerificaSede(id)
{
   var esito=true;
   return esito;
   
   /*deve controllare che ci sia una sola sede Legale e Residenza per stato */
   
   var idA = $("#Sede_IdAccountSede").val();
   var tpS = $("#Sede_IdTipoSede" + id).val();
   var idS = $("#Sede_IdStato" + id).val();
   var ind = $("#Sede_Indirizzo" + id).val();
   var civ = $("#Sede_Civico" + id).val();
   var cap = $("#Sede_Cap" + id).val();   
   var com = $("#Sede_Comune" + id).val();
   var pro = $("#Sede_Provincia" + id).val();

	return esito;
}
</script>