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
	   if Nome="CIG" then 
	      c_CIG   =valo
	   end if 
	   if Nome="NUMCIG" then 
	      c_NumCIG=valo
	   end if 	   
 
   end if 
next 

prefix = "params_"

%>

<script>

function cig_registra(id,action)
{
    if (action=='delete') {
	   if (!confirm('Si desidera cancellare la riga selezionata ?'))
           return false;
	}
	else {
	   var oldN=ValoreDi("NameLoaded");
	   xx=ImpostaValoreDi("NameLoaded",ValoreDi("NameLoadedCIG"));
	   xx=ImpostaValoreDi("DescLoaded",id);
	   yy=ElaboraControlli();
	   xx=ImpostaValoreDi("NameLoaded",oldN);
	
 	   if (yy==false)
	     return false;
	}
	idCauzione=ValoreDi("CIGIdCauzione");
	var dataIn = "";
	dataIn = dataIn + "IdCauzione=" + idCauzione + "&IdCauzioneCIG=" + id 
	dataIn = dataIn + "&action="    + action;
	dataIn = dataIn + "&CIG="       + encodeURI(ValoreDi("params_CIG"     + id));
	dataIn = dataIn + "&DescCIG="   + encodeURI(ValoreDi("params_DescCIG" + id));
	xx=cig_update(dataIn);
}

function confirmModalProcedi()
{
	$('#dismissModalGenerica').click();
}
</script>

<script>
function cig_update(dataIn)
{
   var vp=$("#ModalGenericaVirtualPath").val(); 
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/updateCIG.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	   cig_reload();
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}

function cig_reload()
{
   var idCauzione=ValoreDi("CIGIdCauzione");
   var azione=ValoreDi("CIGAzione");
   var dataIn="params_IdCauzione=" + idCauzione + "&params_Azione=" + azione;
   var vp=$("#ModalGenericaVirtualPath").val(); 
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/setCigDati.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	    $("#CIG_dati").html(msg); 
		var el=ValoreDi("elenco_CIG");
		var de=ValoreDi("c_CIG");
		$("#"+de).val(el);
		var nc=ValoreDi("num_elenco_CIG");
		var dc=ValoreDi("c_NumCIG");
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
	  NameLoadedCIG=""
	  NameLoadedCIG= NameLoadedCIG & prefix & "CIG,TE;"
	  NameLoadedCIG= NameLoadedCIG & prefix & "DescCIG,TE;"
	  %>

      <input type="hidden" name="c_CIG"           id="c_CIG"           value="<%=c_CIG%>">
	  <input type="hidden" name="c_NumCIG"        id="c_NumCIG"        value="<%=c_NumCIG%>">
	  <input type="hidden" name="CIGIdCauzione"   id="CIGIdCauzione"   value="<%=idCauzione%>">
	  <input type="hidden" name="CIGAzione"       id="CIGAzione"       value="<%=azione%>">
	  <input type="hidden" name="NameLoadedCIG"   id="NameLoadedCIG"   value="<%=NameLoadedCIG%>">
	  <div Id="CIG_dati">
      </div>
<script language=javascript>
   cig_reload();
</script>
