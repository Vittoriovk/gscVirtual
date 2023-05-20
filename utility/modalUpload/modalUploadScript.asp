
<script type="text/javascript">
function mostraUpload(id)
{
   // pulisco file 
   var vf=$('#modalUploadNewName' + id).val();
   var vd=$('#modalUploadNewDesc' + id).val();
   $('#modalUploadFile').val("");
   $('#DescDocumentoFile').val(vd);
   $('#modalUploadFileOld').val(vf);
   $('#confirmModalUpload').modal('toggle');
}

function eseguiUpload(idModUpload)
{
   var vf = $('#modalUploadFile').val();
   var of = $('#modalUploadNewName' + idModUpload).val();
   var vd = $('#DescDocumentoFile').val();
   if (vf=="" && of=="") {
       alert("selezionare il documento da caricare");
       return false;
   }
   if (vd=="") {
       alert("descrivere il documento");
       return false;
   } 
   if (vf!="") {   
          
      esito = eseguiUploadFile(idModUpload);
	  descErr=esito.errore;
	  if (descErr=="ok") {
         linkDocumento = esito.modalUploadNewPath; 
         descDocumento = esito.modalUploadNewDesc;
         nomeDocumento = esito.modalUploadNewName;
	     }
      }	  
   else {
      $("#modalUploadNewDesc" + idModUpload).val($('#DescDocumentoFile').val());
      linkDocumento = $("#modalUploadNewPath" + idModUpload).val();
      descDocumento = $("#modalUploadNewDesc" + idModUpload).val();
      nomeDocumento = $("#modalUploadNewName" + idModUpload).val(); 
   } 
   idTableUpload = $("#modalUploadIdUploa" + idModUpload).val();
   xx=loadLinkDocumento(idModUpload,linkDocumento,descDocumento,nomeDocumento,idTableUpload);
   $('#confirmModalUpload').modal('toggle');
}

function loadLinkDocumento(idModUpload,linkDocumento,descDocumento,nomeDocumento,idTableUpload)
{
   var dataIn="";
   dataIn = dataIn + "idModUpload="    + encodeURI(idModUpload);
   dataIn = dataIn + "&linkDocumento=" + encodeURI(linkDocumento);
   dataIn = dataIn + "&nomeDocumento=" + encodeURI(nomeDocumento);
   dataIn = dataIn + "&descDocumento=" + encodeURI(descDocumento);
   dataIn = dataIn + "&idTableUpload=" + encodeURI(idTableUpload);
   
   var vp=$("#localVirtualPath").val(); 
   var ur="/utility/modalUpload/loadShowUploadDoc.asp";
   //$("#nameFileUploaded").val(vp + ur + "?" + dataIn); 
   //alert(vp + ur + "?" + dataIn);
   $.ajax({
      type: "POST",
      async: false,
      url: vp + ur,
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
        $("#row" + idModUpload).html(msg); 
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
}

</script>


<script>
function eseguiUploadFile(idModUpload)
{
    alert('sono qui');
    var form = $('#formUploadModal')[0];
    var data = new FormData(form);

	var map1 = {
          errore: 'ok',
          modalUploadFileMod: '',
		  modalUploadNewName: '', 
		  modalUploadNewPath: '',
		  modalUploadNewDesc: '' 
          };

    // aggiungo il suffisso dei nomi 
   data.append("idUp", idModUpload);
   $("#btnConfirmUpload").prop("disabled", true);
   var vp=$("#localVirtualPath").val();
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "utility/modalUpload/ModalUploadExecute.asp",
	  data: data,
      processData: false,
      contentType: false,
      cache: false,
      timeout: 800000,
      success: function(msg)
      {
		nf = document.getElementById("modalUploadFile").files[0].name;
		map1['modalUploadNewName'] = nf;
        map1['modalUploadFileMod'] = 'S';
        map1['modalUploadNewPath'] = msg;
	    map1['modalUploadNewDesc'] = $('#DescDocumentoFile').val();
        $("#btnConfirmUpload").prop("disabled", false);		
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
	    var descE = xhr.status + ":Chiamata Upload fallita, si prega di riprovare..." + thrownError;
        alert(descE);
		map1['errore'] = descE;
		$("#btnConfirmUpload").prop("disabled", false);
      }
    });   
    return map1; 
}
</script>


