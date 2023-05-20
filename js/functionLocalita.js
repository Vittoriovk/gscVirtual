<script language="JavaScript">

/*
riceve un item stato e provincia.
Se lo stato e' ITALIA crea la lista dati per le provincie 
*/
function absoluteChangeStato(objStato,objProvincia)
{
   const prefix = "absoluteProvincia";
   var st;
   st = $("#" + objStato).val().trim().toUpperCase();
   
   if (!(st=="IT" || st=="ITA" || st=="ITALIA")) {
	   $('#' + objProvincia).attr('list',"");
       return false;
   }
   var pr;
   //recupero provincia 
   pr = "IT";
   // recupero la data list associata al comune 
   var dataList = $('#' + objProvincia).attr('list');
   if (typeof dataList == 'undefined')
		dataList = "";
	var dataListR= prefix + pr;
	//alert(dataListR);
	// non e' la stessa bisogna cambiare 
	if (!(dataList==dataListR)) {
	
	   // data list inesistente lo creo 
	   if ($('#' + dataListR).length==0) 
	   {
	       setDataListInHtml("PROVINCIA_IT",dataListR,pr,dataListR);
	   }
	   $('#' + objProvincia).attr('list',dataListR);	
	   $('#' + objProvincia).val('');
	}   
}


function absoluteChangeProvincia(objStato,objProvincia,objComune)
{
   const prefix = "absoluteComune";
   var st;
   st = $("#" + objStato).val().trim().toUpperCase();
   if (!(st=="IT" || st=="ITA" || st=="ITALIA"))
       return false;
   var pr;
   //recupero provincia 
   pr = $("#" + objProvincia).val();
   pr = getSiglaProvinciaDaProvincia(pr);
   if (pr.length==0){
       $('#' + objComune).attr('list',"");
       return false;
   }
	// recupero la data list associata al comune 
	var dataList = $('#' + objComune).attr('list');
	if (typeof dataList == 'undefined')
		dataList = "";
	var dataListR= prefix + pr;
	// non e' la stessa bisogna cambiare 
	if (!(dataList==dataListR)) {
	
	   // data list inesistente lo creo 
	   if ($('#' + dataListR).length==0) 
	   {
	       setDataListInHtml("COMUNE_BYPROVINCIA_IT",dataListR,pr,dataListR);
	   }
	   $('#' + objComune).attr('list',dataListR);	
	   $('#' + objComune).val('');
	}   
}

</script>