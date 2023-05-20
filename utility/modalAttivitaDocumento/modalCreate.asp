<!--#include virtual="/gscVirtual/include/includeStd.asp"-->

<script>
function MAD_Check()
{
    dl=ValoreDi("DescLoaded");
	nl=ValoreDi("NameLoaded");
	yy=ImpostaValoreDi("NameLoaded",ValoreDi("MAD_NameLoaded"));
	yy=ImpostaValoreDi("DescLoaded","0");
	xx=ElaboraControlli();
	yy=ImpostaValoreDi("NameLoaded",nl);
	yy=ImpostaValoreDi("DescLoaded",dl);
 	if (xx==false)
	   return false;
	$("#MAD_btnConfirm").prop("disabled", true);
	
    var esito="";
	var pathFile=$("#MAD_FileToUpload0").val();
	
	if (pathFile=="")
	   xx="OK:";
	else   
	   xx = MAD_eseguiUploadFile();
	 
	if (xx.substring(0,2)=="OK") {
       var pathFile = xx.substring(3);
	   esito = MAD_eseguiUpdateFile(pathFile);
	     
	}
	$("#MAD_btnConfirm").prop("disabled", false);
	if (esito=="")
       xx=MAD_Reload();

}

function MAD_eseguiUploadFile()
{
   var form = $('#formModalAttivitaDocumento')[0];
   var data = new FormData(form);
   var retVal="";
   
   var vp=$("#hiddenVirtualPath").val();
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "utility/modalAttivitaDocumento/UploadExecute.asp",
	  data: data,
      processData: false,
      contentType: false,
      cache: false,
      timeout: 800000,
      success: function(msg)
      {
        retVal = msg;
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
	    var descE = xhr.status + ":Chiamata esegui Upload fallita, si prega di riprovare..." + thrownError;
        alert(descE);
		retVal = "ERR:Chiamata Fallita"
      }
    });   
	return retVal;
}
function MAD_eseguiUpdateFile(fileLoaded)
{
   var form = $('#formModalAttivitaDocumento')[0];
   var data = new FormData(form);
   data.append("MAD_PathFile", fileLoaded);
   var retVal="";
   
   var vp=$("#hiddenVirtualPath").val();
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "utility/modalAttivitaDocumento/UpdateExecute.asp",
	  data: data,
      processData: false,
      contentType: false,
      cache: false,
      timeout: 800000,
      success: function(msg)
      {
        retVal = msg;
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
	    var descE = xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError;
        alert(descE);
		retVal = "ERR:Chiamata Fallita"
      }
    });   
	return retVal;
}

</script>

<%
IdAttivitaDocumento=cdbl("0" & Trim(Request("IdAttivitaDocumento")))

If IdAttivitaDocumento=0 then 
   response.end
end if 

'leggo i campi del db 
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.CursorLocation = 3 
Rs.Open "select * from AttivitaDocumento where IdAttivitaDocumento=" & IdAttivitaDocumento , ConnMsde
if err.number=0 then 
   if Rs.EOF then 
      response.end 
   end if 
   IdAttivita        = Rs("IdAttivita")
   IdNumAttivita     = Rs("IdNumAttivita")
   IdDocumento       = Rs("IdDocumento")
   DescDocumento     = Rs("DescDocumento")
   FlagObbligatorio  = Rs("FlagObbligatorio")
   FlagDataScadenza  = Rs("FlagDataScadenza")
   IdUpload          = Rs("IdUpload")
   DataValidazione   = Rs("DataValidazione")
   NoteValidazione   = Rs("NoteValidazione")
   IdTipoValidazione = Rs("IdTipoValidazione")
   Rs.close 
end if 
DocumentoDaCaricare = LeggiCampo("Select * from Documento Where IdDocumento = " & IdDocumento,"DescDocumento")
if DescDocumento="" then 
   DescDocumento = DocumentoDaCaricare
end if 
if Cdbl(IdUpload)>0 then 
   Rs.Open "select * from Upload where IdUpload=" & IdUpload , ConnMsde
   if err.number=0 then 
      if Rs.EOF = false then 
         IdTabella          = Rs("IdTabella")
         IdTabellaKeyInt    = Rs("IdTabellaKeyInt")
         IdTabellaKeyString = Rs("IdTabellaKeyString")
         DataUpload         = Rs("DataUpload")
         TimeUpload         = Rs("TimeUpload")
         IdTipoDocumento    = Rs("IdTipoDocumento")
         DescBreve          = Rs("DescBreve")
         DescEstesa         = Rs("DescEstesa")
         NomeDocumento      = Rs("NomeDocumento")
         PathDocumento      = Rs("PathDocumento")
         ValidoDal          = Rs("ValidoDal")
         ValidoAl           = Rs("ValidoAl")
      else 
         ValidoDal          = Dtos()
         ValidoAl           = 99991231
      end if 
      Rs.close 
   end if 
end if 
if Ucase(IdAttivita)="FORMAZIONE" then 
   Titolo   ="Documenti per formazione"
   subTitolo=LeggiCampo("select * from formazione where IdFormazione=" & IdNumAttivita,"DescFormazione")
end if 
if Cdbl(FlagDataScadenza)=1 then 
   ShowValidoDal=True
   ShowValidoAl =True
else 
   ShowValidoDal=false
   ShowValidoAl =false 
end if 

prefix = "MAD_"
id="0"
NameLoaded=""
%>


<div class="modal fade" id="confirmModalAttivitaDocumento"  aria-hidden="true" role="dialog">
  <!-- a pieno schermo -->
  <div class="modal-dialog modal-lg">
    <div class="modal-content">
        <div class="row">
			<div class="col-10"><h4>&nbsp;&nbsp;<%=Titolo%></h4></div>
			<div class="col-2">
                 <button type="button" Id="dismissModalAttivitaDocumento" class="close" data-dismiss="modal">
                 <span aria-hidden="true">×&nbsp;&nbsp;</span><span class="sr-only">Chiudi</span>
                 </button>
			</div>
			<div class="col-1"></div>
		</div>

	  <%if SubTitolo<>"" then %>
        <div class="row">
			<div class="col-10"><h6>&nbsp;&nbsp;&nbsp;<%=SubTitolo%></h6></div>
			<div class="col-2">&nbsp;&nbsp;</div>
			<div class="col-1"></div>
		</div>
	  
	  <%end if  %>
      <div class="modal-body"> 
	  <form method="POST" enctype="multipart/form-data" id="formModalAttivitaDocumento">
		<div class="row">
		   <div class="col-5">
                <div class="form-group ">
			     <%xx=ShowLabel("Documento")%>
				 <input type="text" class="form-control" readonly value="<%=DocumentoDaCaricare%>" >
                </div>		
		   </div>
		   <div class="col-5">
                <div class="form-group ">
			     <%xx=ShowLabel("Descrizione Documento")
				 nn=prefix & "DescDocumento"
				 NameLoaded = NameLoaded & nn & ",TE"
				 nn=nn & id
				 %>
				 <input type="text" class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=DescDocumento%>" >
                </div>		
		   </div>				   
          </div>
		  
          <div class="row">
             <div class="col-5">
                <%xx=ShowLabel("Documento da caricare")
				nn=prefix & "FileToUpload"
				'obbligatorio se non gia' caricato 
				if NomeDocumento="" then
				   NameLoaded = NameLoaded & ";" & nn & ",TE"
				end if 
				nn=nn & id
				%>
                <input class="form-control" style="height:2.5em;" type="file" name="<%=nn%>" id="<%=nn%>">
             </div>
			 <%if NomeDocumento<>"" then %>
             <div class="col-5">
                 <div class="form-group ">
                     <%xx=ShowLabel("Documento gia' caricato")%>
                     <input type="text" readonly class="form-control" value="<%=NomeDocumento%>" />
                 </div>        
             </div> 
			 <%end if %>
          </div>
		  <% if ShowValidoDal=True or ShowValidoAl =True then %>
		  <div class="row">
		     <div class="col-3">
                 <div class="form-group ">
                     <%xx=ShowLabel("Valido Dal")
					 nn=prefix & "ValidoDal"
					 NameLoaded = NameLoaded & ";" & nn & ",DTO"
					 nn=nn & id
					 if ValidoDal=0 then 
					    ValidoDal=DtoS()
					 end if 
					 %>
             	      <input type="text" <%=readonly%> name="<%=nn%>" id="<%=nn%>" value="<%=Stod(ValidoDal)%>" 
                      class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >					 
                 </div>        
             </div> 
		     <div class="col-3">
                 <div class="form-group ">
                     <%xx=ShowLabel("Valido Al")
					 nn=prefix & "ValidoAl"
					 NameLoaded = NameLoaded & ";" & nn & ",DTO"
					 nn=nn & id
					 ValidoAl="20501231"
					 %>
             	      <input type="text" <%=readonly%> name="<%=nn%>" id="<%=nn%>" value="<%=Stod(ValidoAl)%>" 
                      class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >					 
                 </div>        
             </div> 

		  </div>
		  <%end if %>
		  
		  <input type="hidden" name="MAD_NameLoaded"          id="MAD_NameLoaded"          value="<%=NameLoaded%>">
		  <input type="hidden" name="MAD_IdAttivitaDocumento" id="MAD_IdAttivitaDocumento" value="<%=IdAttivitaDocumento%>">
		  <input type="hidden" name="MAD_IdUpload"            id="MAD_IdUpload"            value="<%=IdUpload%>">
		  <input type="hidden" name="MAD_IdTipoDocumento"     id="MAD_IdTipoDocumento"     value="<%=IdDocumento%>">
		  
		 </form>		  
      </div> 
	  <%
		 if session("EsitoCallSelIndirizzo")="" then 
      %>
      <div class="modal-footer">
         <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
		 <!-- viene scritta dalla call precedente -->
         <button type="button" class="btn btn-primary" Id="MAD_btnConfirm" onclick="MAD_Check();";>Procedi</button>
      </div>
	  <%
	     end if 
      %>
    </div>
  </div>
</div>