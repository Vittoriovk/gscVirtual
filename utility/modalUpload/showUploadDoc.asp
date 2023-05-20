<!--
parametri di ingresso necessari 
IdModUpload    : testo      - id della riga (deve essere assoluto nella pagina)
InsertRow      : true/false - utilizzato per creare la riga contenitore
canUpload      : true/false - possibilita' di modificare i dati (upload e descrizione)
IsFileUploadMd : S/N        - indica se la riga è stata modificata
NomeFileUpload : testo      - nome del file
Linkdocumento  : path file  - path del file            
DescFileUpload : testo      - descrizione del file
IdTableUpload  : numerico   - id della tabella Upload 
-->


<%if InsertRow then %>
<div class="row" id="row<%=IdModUpload%>">
<%end if %>
<div class="col-9">
   <%xx=ShowLabel("Allegato")%>
   <input type="text" readonly class="form-control" name="nameFileUploaded" id="nameFileUploaded" value="<%=NomeFileUpload%>" >
</div> 

<div class="col-3">
<% xx=ShowLabel("Azioni") %>
<br>
<!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->	
<%
'abilito cancellazione o upload 
if canUpload then 
%>
  <input type="hidden" name="modalUploadFileMod<%=IdModUpload%>" id="modalUploadFileMod<%=IdModUpload%>" value="<%=IsFileUploadMd%>">
  <input type="hidden" name="modalUploadIdUploa<%=IdModUpload%>" id="modalUploadIdUploa<%=IdModUpload%>" value="<%=IdTableUpload%>">
  <input type="hidden" name="modalUploadOldName<%=IdModUpload%>" id="modalUploadOldName<%=IdModUpload%>" value="<%=NomeFileUpload%>">
  <input type="hidden" name="modalUploadNewName<%=IdModUpload%>" id="modalUploadNewName<%=IdModUpload%>" value="<%=NomeFileUpload%>">
  <input type="hidden" name="modalUploadNewPath<%=IdModUpload%>" id="modalUploadNewPath<%=IdModUpload%>" value="<%=Linkdocumento%>">
  <input type="hidden" name="modalUploadOldDesc<%=IdModUpload%>" id="modalUploadOldDesc<%=IdModUpload%>" value="<%=DescFileUpload%>">
  <input type="hidden" name="modalUploadNewDesc<%=IdModUpload%>" id="modalUploadNewDesc<%=IdModUpload%>" value="<%=DescFileUpload%>">


<%
   RiferimentoA=";#;;2;uplo;Aggiorna;;mostraUpload('" & IdModUpload & "');N"
   %>
   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->		
   <%
   if Linkdocumento<>"" then 
      RiferimentoA=";#;;2;dele;Rimuovi;;rimuoviUpload('" & IdModUpload & "');N"
   %>
   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
   <%	  
   end if 
end if 
%>
</div>
<%if InsertRow then %>
</div>
<%end if %>
