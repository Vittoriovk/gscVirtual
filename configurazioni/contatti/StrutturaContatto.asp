	<%
	err.clear 
	'si aspetta in input il nome della struttura 
	if NomeStruttura="" then 
	   NomeStruttura="CONTATTO_GENERICO"
	end if 
	if DescStruttura="" then 
	   DescStruttura="CONTATTI"
	end if 
	
	%>


  <a class="btn btn-info" data-toggle="collapse" href="#collapse<%=NomeStruttura%>" role="button" 
     onclick="LoadContatti('<%=NomeStruttura%>')"
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
'la struttura è del tipo nomeCampo,Descrizione,default su tabella;NomeCampo ....
'se Descrizione su tabella è vuota non viene valorizzato
campiStruttura=""
campiStruttura=campiStruttura + "IdAccount,," & IdAccount & ";"
campiStruttura=campiStruttura + "IdAccountContatto,,;"
campiStruttura=campiStruttura + "IdTipoContatto,,;"
campiStruttura=campiStruttura + "DescTipoContatto,Tipo Contatto,;"
campiStruttura=campiStruttura + "DescContatto,Descrizione,;"
campiStruttura=campiStruttura + "NoteContatto,Note,;"
campiStruttura=campiStruttura + "ValFlagPrincipale,Principale,NO;"
campiStruttura=campiStruttura + "Azioni,Azioni,;"
%>

<input type="hidden" name="CONTATTO_NOME_STRUTTURA" id="CONTATTO_NOME_STRUTTURA"   value="<%=NomeStruttura%>">
 <!-- entita' da gestire   -->
<input type="hidden" name="<%=NomeStruttura%>_Entita"    id="<%=NomeStruttura%>_Entita"   value="CONTATTO">
<input type="hidden" name="<%=NomeStruttura%>_Account"   id="<%=NomeStruttura%>_Account"  value="<%=IdAccount%>">

<input type="hidden" name="<%=NomeStruttura%>_Header"    id="<%=NomeStruttura%>_Header"   value="<%=campiStruttura%>">
 <!-- prefisso dei campi del form  -->
<input type="hidden" name="<%=NomeStruttura%>_Prefix"    id="<%=NomeStruttura%>_Prefix"   value="Contatto_">
 <!-- postfix dei campi del form  -->
<input type="hidden" name="<%=NomeStruttura%>_Postfix"   id="<%=NomeStruttura%>_Postfix"  value="0">

<input type="hidden" name="<%=NomeStruttura%>_Oper"      id="<%=NomeStruttura%>_Oper"     value="<%=flagOperStruttura%>">
<input type="hidden" name="<%=NomeStruttura%>_OperCall"  id="<%=NomeStruttura%>_OperCall" value="attivaFormContatto">

<!--#include virtual="/gscVirtual/configurazioni/contatti/FormContatto.asp"--> 

<script>
function LoadContatti(NomeStruttura)
{
   evaluateMinusPlus(NomeStruttura);
   var act=$("#" + NomeStruttura + "_plusMinus").val();
   	if (act=='+') {
		return true;
	}
   var dataIn="sendData=" + $("#sendDataProd").val(); 
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
      url: vp + "/configurazioni/contatti/ContattiListaForm.asp",
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
function attivaFormContatto(tabName,idx)
{
	$("#" + tabName + "_Idx0").val(idx);
	xx=resetFormContatto(tabName);
	xx=$('#' + tabName + 'Modal').modal('toggle');
}

function resetFormContatto(tabName)
{
	// legge la struttura e carica i dati sul form
	populateFormByRow(tabName);

	var idx = $("#" + tabName + "_Idx0").val();
	// abilito pulsante cancellazione se è una riga valida
	if (idx==0)
		$('#divformContattoDelete').hide();
	else
		$('#divformContattoDelete').show();

	// qui effettuo solo operazioni ad hoc 
	
	// valorizzo DescTipoContatto
	var pro = $("#Contatto_DescTipoContatto0").val();
	var arr = Array.from(document.querySelector("#Contatto_IdTipoContatto0").options);
	var val = "";

	arr.forEach(function(option_element) {
		var option_text  = option_element.text;
		var option_value = option_element.value;	
		if (option_text == pro)
				val = option_value;
		});
	if (val.length>0)
		$("#Contatto_IdTipoContatto0").val(val);
		
	// valorizzo il flagprincipale
	var fla = $("#Contatto_ValFlagPrincipale0").val();
	$("#Contatto_FlagPrincipale" + fla + "0").prop( "checked", true );
	$("#Contatto_FlagPrincipale" + fla + "0").click();
	
}

function localContattoChangeTipo(id)
{
	var v;
	v = $("#Contatto_IdTipoContatto" + id + " option:selected").text();
	$("#Contatto_DescTipoContatto" + id).val(v); 
}



function localContattoSubmit(tabName)
{
var xx,yy,oldN,oldD;
    var id=0;
    xx=false;
	oldN=ValoreDi("NameLoaded");
	oldD=ValoreDi("DescLoaded");
	yy=ImpostaValoreDi("NameLoaded","Contatto_IdTipoContatto,TE;Contatto_DescContatto,TE");
	yy=ImpostaValoreDi("DescLoaded",id);
    xx=ElaboraControlli();
	yy=ImpostaValoreDi("NameLoaded",oldN);
	yy=ImpostaValoreDi("DescLoaded",oldD);
	if (xx=true) {
		xx=localVerificaContatto(id);
		if (xx==false)
			bootbox.alert("Contatto esistente in archivio");
	}
 	if (xx==false) 
	   return false;
   
	/* recupero l'indice : se = 0 è nuovo altrimenti è una modifica */
	var idx = $("#" + tabName + "_Idx0").val();
	/* se nuovo metto azione in NEW */
	if (idx==0)
		yy=ImpostaValoreDi("Contatto_Azioni0","NEW");
	else 
		yy=ImpostaValoreDi("Contatto_Azioni0","MOD");
	var riga = createStructureRow(tabName);
	
	$("#btnS_" + tabName).prop("disabled", true);
	updateContatto(tabName,riga);
	$('#' + tabName + 'Modal').modal('hide');
	LoadContatti(tabName);
	LoadContatti(tabName);
	$("#btnS_" + tabName).prop("disabled", false);
}

function localContattoRemove(tabName)
{
	yy=ImpostaValoreDi("Contatto_Azioni0","DEL");
	var riga = createStructureRow(tabName);
	updateContatto(tabName,riga);
	$('#' + tabName + 'Modal').modal('hide');
	LoadContatti(tabName);
	LoadContatti(tabName);
}

function updateContatto(NomeStruttura,riga)
{
   var dataIn="sendData=" + $("#sendDataContatto").val(); 
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
      url: vp + "/configurazioni/contatti/ContattiUpdate.asp",
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

function localContattoSetFlag(id,val)
{
	$("#Contatto_FlagPrincipale" + id ).val(val);
	$("#Contatto_ValFlagPrincipale" + id ).val(val);
}


function localVerificaContatto(id)
{
   var esito=true;
   return esito;
   
   var idA = $("#Contatto_IdAccountContatto").val();
   var tpC=ValoreDi("Contatto_IdTipoContatto" + id);
   var deC=ValoreDi("Contatto_DescContatto" + id );
   var dataIn="TipoVerifica=CONTATTO&IdAccountContatto="  + id + "&IdAccount=" + idA + "&IdTipoContatto=" + tpC + "&DescContatto=" + deC;
   
   $.ajax({
      type: "POST",
	  async: false,
      url: "/gscVirtual/utility/VerificaDuplicato.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
        if (msg=="")
			esito = true;
		else {
			esito = false;
			}
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });
	
	return esito;
}
</script>
