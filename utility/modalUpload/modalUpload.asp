<%
'si aspetta in input un IdModUpload
if Titolo="" then 
   Titolo="Upload documento"
end if 
%>
 
<div class="modal fade" id="confirmModalUpload"  aria-hidden="true" role="dialog">
  <!-- a pieno schermo -->
  <div class="modal-dialog modal-lg">
    <div class="modal-content">
        <div class="row">
            <div class="col-10"><h4>&nbsp;&nbsp;<%=Titolo%></h4></div>
            <div class="col-2">
                 <button type="button" Id="dismissModalUpload" class="close" data-dismiss="modal">
                 <span aria-hidden="true">×&nbsp;&nbsp;</span><span class="sr-only">Chiudi</span>
                 </button>
            </div>
            <div class="col-1"></div>
        </div>

      <%if SubTitolo<>"" then %>
        <div class="row">
            <div class="col-10"><h6>&nbsp;&nbsp;&nbsp;<%=SubTitolo%></h6></div>
            <div class="col-2">&nbsp;&nbsp</div>
            <div class="col-1"></div>
        </div>
      
      <%end if  %>
      <div class="modal-body"> 
      <form method="POST" enctype="multipart/form-data" id="formUploadModal">
          <div class="row">
             <div class="col-1">
             </div> 
             <div class="col-6">
                <%xx=ShowLabel("Documento da caricare")%>
                <input class="form-control" style="height:2.5em;" type="file" name="modalUploadFile" id="modalUploadFile">
             </div>
             <div class="col-3">
                 <div class="form-group ">
                     <%xx=ShowLabel("Documento gia' caricato")%>
                     <input type="text" readonly class="form-control" id="modalUploadFileOld" name="modalUploadFileOld"/>
                 </div>        
             </div> 
          </div>
          <div class="row">
             <div class="col-1">
             </div> 
            
             <div class="col-8">
                 <div class="form-group ">
                     <%xx=ShowLabel("Descrizione del documento")%>
                     <input type="text" class="form-control" id="DescDocumentoFile" name="DescDocumentoFile"/>
                 </div>        
             </div>    
          </div>
 
      </form>
        
      </div> 

      <div class="modal-footer">
         <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
         <button type="button" class="btn btn-primary" onclick="eseguiUpload('<%=IdModUpload%>');" id="btnConfirmUpload";>Carica</button>
      </div>
    </div>
  </div>
</div>

