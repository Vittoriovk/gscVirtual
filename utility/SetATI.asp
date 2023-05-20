<!--#include virtual="/gscVirtual/include/includeStd.asp"-->

<%
session("EsitoCallSelIndirizzo") = "" 
IdCauzione = session("params_IdCauzione")
azione     = session("params_Azione")
'lista dei campi da valorizzare tipo 
'  Campo=nome_delcampo;
ListaCampi= Session("params_ListaCampi")
ArCampi = split(ListaCampi,";")
c_ati    = ""
c_Numati = ""
for J=lbound(ArCampi) to ubound(ArCampi)-1
    str=ArCampi(j)
	ArDett = split(str,":")
	if ubound(ArDett)=1 then 
	   nome=ucase(trim(ArDett(0)))
	   valo=trim(ArDett(1))
	   if Nome="ATI" then 
	      c_ATI   =valo
	   end if 
	   if Nome="NUMATI" then 
	      c_NumATI=valo
	   end if 	   
 
   end if 
next 

prefix = "params_"

%>

<script>

function ati_registra(id,action)
{
    if (action=='delete') {
	   if (!confirm('Si desidera cancellare la riga selezionata ?'))
           return false;
	}
	else {
	   var oldN=ValoreDi("NameLoaded");
	   xx=ImpostaValoreDi("NameLoaded",ValoreDi("NameLoadedATI"));
	   xx=ImpostaValoreDi("DescLoaded",id);
	   yy=ElaboraControlli();
	   xx=ImpostaValoreDi("NameLoaded",oldN);
	
 	   if (yy==false)
	     return false;
	}
	idCauzione=ValoreDi("ATIIdCauzione");
	var dataIn = "";
	dataIn = dataIn + "IdCauzione=" + idCauzione + "&IdCauzioneATI=" + id 
	dataIn = dataIn + "&action="    + action;
	xx=ati_update(dataIn);
}

function confirmModalProcedi()
{

	var s=$('input[name="sel"]:checked').val();

	var ATI = $('#params_Provincia' + s).val();

	d=$('#c_Provincia').val();
	$('#'+d).val(ATI);	
	$('#dismissModalGenerica').click();
}
</script>

<script>
function ati_update(dataIn)
{
   var vp=$("#ModalGenericaVirtualPath").val(); 
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/updateATI.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	   ati_reload();
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}

function ati_reload()
{
   var idCauzione=ValoreDi("ATIIdCauzione");
   var azione=ValoreDi("ATIAzione");
   var dataIn="params_IdCauzione=" + idCauzione + "&params_Azione=" + azione;
   var vp=$("#ModalGenericaVirtualPath").val(); 
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/setAtiDati.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	    $("#ATI_dati").html(msg); 
		var el=ValoreDi("elenco_ATI");
		var de=ValoreDi("c_ATI");
		$("#"+de).val(el);
		var nc=ValoreDi("num_elenco_ATI");
		var dc=ValoreDi("c_NumATI");
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
	  NameLoadedATI=""
	  %>

      <input type="hidden" name="c_ATI"           id="c_ATI"           value="<%=c_ati%>">
	  <input type="hidden" name="c_NumATI"        id="c_NumATI"        value="<%=c_NumATI%>">
	  <input type="hidden" name="ATIIdCauzione"   id="ATIIdCauzione"   value="<%=idCauzione%>">
	  <input type="hidden" name="ATIAzione"       id="ATIAzione"       value="<%=azione%>">
	  <input type="hidden" name="NameLoadedATI"   id="NameLoadedATI"   value="<%=NameLoadedATI%>">
	  <div Id="ATI_dati">
      </div>
<script language=javascript>
   ati_reload();
</script>
