<script>
// crea lìil data.list e lo appende all'html 
function setDataListInHtml(action,id,attr,refId)
{
   var dataIn="action=" + action + "&id=" + id + "&attr=" + attr;
   var esito = callFunctionDataList(dataIn);
   //alert(dataIn + " " + esito);
   if (esito.length>0) {
      // id inesistente lo appendo all'html 
      if ($('#' + refId).length==0) {
	     var frm = $('form[name="Fdati"]');
		 $(frm).append(esito);
      }
   }
   return true ;
}

function callFunctionDataList(dataIn)
{
   var vp=$("#hiddenVirtualPath").val(); 
   var esito="";
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/loadFunctionDataList.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	   esito = msg;
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
	return esito;
  
}
</script>
