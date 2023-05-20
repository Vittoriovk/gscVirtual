<script language="JavaScript">
function CheckDatiTecnici(Id)
{
	xx=ImpostaColoreFocus("IdElenco" + Id,"","white");
	locName=ValoreDi("DescLoaded");
	yy=ImpostaValoreDi("DescLoaded",Id);

	xx=ElaboraControlli();
	
	var te = ValoreDi("IdTipoCampo" + Id);
	var el = ValoreDi("IdElenco" + Id);
	
	te = te.toUpperCase();
	var fl = (te=="MENUDI" || te=="SCELTA" || te=="SPUNTA" );
	
	if (xx==true && fl==true ) {
		xx=ControllaCampo("IdElenco" + Id,"LI");
		if (xx==false)
			bootbox.alert("Elenco richiesto");
	}
	if (xx==true && fl==false && trim(el)!="-1" ) {
		xx=ImpostaColoreFocus("IdElenco" + Id,"","yellow");
		bootbox.alert("Elenco NON richiesto");
		xx=false;
	}
	
 	if (xx==false)
	   return false;

   ImpostaValoreDi("Oper","update");
   document.Fdati.submit(); 
   
}

function CheckProdotto(Id)
{
	xx=ImpostaColoreFocus("Prezzo" + Id,"","white");
	locName=ValoreDi("DescLoaded");
	yy=ImpostaValoreDi("DescLoaded",Id);

	xx=ElaboraControlli();
	
	var te = ValoreDi("FlagPrezzoFisso" + Id);
	var el = ValoreDi("Prezzo" + Id);
	var va = parseFloat(el.replace(",","."));
	te = te.toUpperCase();
	
	if (xx==true && te=='NO' && va>0 ) {
		xx=ImpostaColoreFocus("Prezzo" + Id,"","yellow");
		bootbox.alert("Prezzo non richiesto");
		xx=false;
	}	
	if (xx==true && te=='SI' && va==0 ) {
		xx=ImpostaColoreFocus("Prezzo" + Id,"","yellow");
		bootbox.alert("Prezzo richiesto");
		xx=false;
	}
	
 	if (xx==false)
	   return false;

   ImpostaValoreDi("Oper","update");
   document.Fdati.submit(); 
   
}
</script>