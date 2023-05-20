<script language="JavaScript">

/* richiama il processo di update */
function callRowUpdate(tabName,idx)
{
   var dataStru = $("#"+tabName + "_Header").val();
   
   var iRow = tabName + "_Row_" + idx;
   var dataRows = $("#" + iRow).val();
   
   var vp=$("#localVirtualPath").val();
   
   var dataIn="ns=" + tabName + "&ss=" + dataStru "&vs=" + dataRows ;  
   alert(dataIn);
   return true;
   $.ajax({
      type: "POST",
      async: false,
      url: vp + "utility/updateRows.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
		alert(msg);
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
}
</script>