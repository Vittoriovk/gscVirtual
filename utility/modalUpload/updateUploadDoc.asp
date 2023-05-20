

<!--
parametri di ingresso necessari 
IdTabella          : testo      - Tabella a cui relazionare l'upload 
IdTabellaKeyInt    : true/false - chiave numerica
IdTabellaKeyString : true/false - chiave stringa 
-->

<%
if UpdateUploadDoc="S" then 
   'leggo i dati salvati e recupero i valori 
   modalUploadFilesId=Request("modalUploadFilesId")
   ArDati=split(modalUploadFilesId,";")
 
   for J=lbound(ArDati) to Ubound(ArDati)
      IdModUpload = trim(ArDati(j))
	  
      if IdModUpload<>"" then 
         IsFileUploadMd = request("modalUploadFileMod" & IdModUpload ) 
         'il record è stato modificato
         if IsFileUploadMd = "S" then 
		    
            'recupero l'id dell'upload 
            IdTableUpload   = Cdbl("0" & Request("modalUploadIdUploa" & IdModUpload))
			IdTipoDocumento = Request("modalUploadTipoDoc" & IdModUpload)
            NomeFileUpload  = Request("modalUploadNewName" & IdModUpload)
            Linkdocumento   = Request("modalUploadNewPath" & IdModUpload)
            DescFileUpload  = Request("modalUploadNewDesc" & IdModUpload)
			
			
			
			Set DizDatabase = CreateObject("Scripting.Dictionary")
			xx = InizializeUpload(DizDatabase,IdTableUpload)

            xx=SetDiz(DizDatabase,"IdTabella"         ,IdTabella)
            xx=SetDiz(DizDatabase,"IdTabellaKeyInt"   ,IdTabellaKeyInt)
            xx=SetDiz(DizDatabase,"IdTabellaKeyString",IdTabellaKeyString)
            xx=SetDiz(DizDatabase,"IdTipoDocumento"   ,IdTipoDocumento)
            xx=SetDiz(DizDatabase,"DescBreve"         ,DescFileUpload)
            xx=SetDiz(DizDatabase,"DescEstesa"        ,DescFileUpload)
            xx=SetDiz(DizDatabase,"NomeDocumento"     ,NomeFileUpload)
            xx=SetDiz(DizDatabase,"PathDocumento"     ,Linkdocumento)
   
            xx = UpdateUpload(DizDatabase)
response.write "eccomi:" & err.description

            end if 
      end if 
   next 
end if 
modalUploadFilesId=""
%>


