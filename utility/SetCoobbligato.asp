<!--#include virtual="/gscVirtual/include/includeStd.asp"-->

<%
session("EsitoCallSelIndirizzo") = "" 
IdCauzione = session("params_IdCauzione")
azione     = session("params_Azione")
'lista dei campi da valorizzare tipo 
'  Campo=nome_delcampo;
ListaCampi= Session("params_ListaCampi")
ArCampi = split(ListaCampi,";")
c_Coobbligato = ""
c_NumCoobbligato = ""
for J=lbound(ArCampi) to ubound(ArCampi)-1
    str=ArCampi(j)
	ArDett = split(str,":")
	if ubound(ArDett)=1 then 
	   nome=ucase(trim(ArDett(0)))
	   valo=trim(ArDett(1))
	   if Nome="COOBBLIGATI" then 
	      c_Coobbligato=valo
	   end if 
	   if Nome="NUMCOOB" then 
	      c_NumCoobbligato=valo
	   end if 
	   
   end if 
next 

prefix = "params_"

%>

<script>

function coobbligato_registra(id,action)
{
    if (action=='delete') {
	   if (!confirm('Si desidera cancellare la riga selezionata ?'))
           return false;
	}
	else {
	   var oldN=ValoreDi("NameLoaded");
	   xx=ImpostaValoreDi("NameLoaded",ValoreDi("NameLoadedCoob"));
	   xx=ImpostaValoreDi("DescLoaded",id);
	   yy=ElaboraControlli();
	   xx=ImpostaValoreDi("NameLoaded",oldN);
	
 	   if (yy==false)
	     return false;
	}
	idCauzione=ValoreDi("CoobIdCauzione");
	var dataIn = "";
	dataIn = dataIn + "IdCauzione=" + idCauzione + "&IdCauzioneCoobbligato=" + id 
	dataIn = dataIn + "&action="    + action;
	dataIn = dataIn + "&RagSoc="    + encodeURI(ValoreDi("params_RagSoc"    + id));
	dataIn = dataIn + "&PI="        + encodeURI(ValoreDi("params_PI"        + id));
	dataIn = dataIn + "&CF="        + encodeURI(ValoreDi("params_CF"        + id));
	dataIn = dataIn + "&Indirizzo=" + encodeURI(ValoreDi("params_Indirizzo" + id));
	dataIn = dataIn + "&Comune="    + encodeURI(ValoreDi("params_Comune"    + id));
	dataIn = dataIn + "&Cap="       + encodeURI(ValoreDi("params_Cap"       + id));
	dataIn = dataIn + "&Provincia=" + encodeURI(ValoreDi("params_Provincia" + id));
	xx=coobbligato_update(dataIn);
}

function confirmModalProcedi()
{

	var s=$('input[name="sel"]:checked').val();

	var coob = $('#params_Provincia' + s).val();

	d=$('#c_Provincia').val();
	$('#'+d).val(coob);	
	$('#dismissModalGenerica').click();
}
</script>

<script>
function coobbligato_update(dataIn)
{
   var vp=$("#ModalGenericaVirtualPath").val(); 
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/updateCoobbligato.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	   coobbligato_reload();
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}

function coobbligato_reload()
{
   var idCauzione=ValoreDi("CoobIdCauzione");
   var azione=ValoreDi("CoobAzione");
   var dataIn="params_IdCauzione=" + idCauzione + "&params_Azione=" + azione;
   var vp=$("#ModalGenericaVirtualPath").val(); 
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/setCoobbligatoDati.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	    $("#coobb_dati").html(msg); 
		var el=ValoreDi("elenco_coobbligati");
		var de=ValoreDi("c_Coobbligato");
		$("#"+de).val(el);
		var nc=ValoreDi("num_elenco_coobbligati");
		var dc=ValoreDi("c_NumCoobbligato");
		$("#"+dc).val(nc);
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}

</script>
      <%
	  prefix = "params_"
	  NameLoadedCoob=""
	  NameLoadedCoob= NameLoadedCoob & prefix & "RagSoc,TE;"
	  NameLoadedCoob= NameLoadedCoob & prefix & "Indirizzo,TE;"
	  NameLoadedCoob= NameLoadedCoob & prefix & "Comune,TE;"
	  NameLoadedCoob= NameLoadedCoob & prefix & "Provincia,TE;"
	  'response.write "dddd:" & c_Coobbligato
	  %>

      <input type="hidden" name="c_Coobbligato"    id="c_Coobbligato"    value="<%=c_Coobbligato%>">
	  <input type="hidden" name="c_NumCoobbligato" id="c_NumCoobbligato" value="<%=c_NumCoobbligato%>">
	  <input type="hidden" name="CoobIdCauzione"   id="CoobIdCauzione"   value="<%=idCauzione%>">
	  <input type="hidden" name="CoobAzione"       id="CoobAzione"       value="<%=azione%>">
	  <input type="hidden" name="NameLoadedCoob"   id="NameLoadedCoob"   value="<%=NameLoadedCoob%>">
	  <div Id="coobb_dati">
      </div>
<script language=javascript>
   coobbligato_reload();
</script>
