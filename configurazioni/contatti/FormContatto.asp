<!-- Nome del div da mostrare -->
<div class="modal fade" id="<%=NomeStruttura%>Modal" role="dialog" aria-hidden="true">
  <div class="modal-dialog modal-lg">
    <div class="modal-content">
      <div class="modal-header">

	    <!-- Descrizione dell'oggetto -->
        <h2>Gestione Contatto</h2> 

      </div>

      <div class="modal-body"> 
   <!-- indice della riga in elaborazione : deve esserci sempre  -->
   <input type="hidden" name="<%=NomeStruttura%>_Idx0"          id="<%=NomeStruttura%>_Idx0"     value="">

   <!-- identificativo assoluto della riga sul db : se non esiste deve essere inserito  -->
   <input type="hidden" name="Contatto_IdAccount0"              id="Contatto_IdAccount0" value="<%=IdAccount%>">	 
   
   <!-- identificativo assoluto della riga sul db : se non esiste deve essere inserito  -->
   <input type="hidden" name="Contatto_IdAccountContatto0"      id="Contatto_IdAccountContatto0" value="">	  
   <!-- elenco dei campi della struttura legati a liste : necessari per l'elenco -->
   <!-- per ognuno mettere le opportune funzioni di popolamento al cambio della option -->
   <input type="hidden" name="Contatto_DescTipoContatto0"       id="Contatto_DescTipoContatto0" value="">
   <input type="hidden" name="Contatto_Azioni0"                 id="Contatto_Azioni0"           value="">  
   <%
   if l_id="" then 
      l_Id = "0"
   end if 
   if IdTipoContatto="" then 
      IdTipoContatto="ALTR"
   end if    
   if FlagPrincipale="" then 
      FlagPrincipale="NO"
   end if 
   
   ao_Eve = "localContattoChangeTipo('" & l_id & "')"
   ao_Att = "0"                       'indica se deve mettere vuoto 
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Tipo Contatto</p>
      </div>   
   
      <div class = "col-8">
	     <%
	     response.write ListaDbChangeCompleta("SELECT * From TipoContatto order By DescTipoContatto","Contatto_IdTipoContatto" & l_Id,IdTipoContatto  ,"IdTipoContatto","DescTipoContatto" ,ao_Att   ,ao_Eve,""   ,""       ,""   ,""    ,ao_cla)
	     %>
      </div>

      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div> 
   </div>

 
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Contatto</p>
      </div> 
	  <div class="col-8">
	  <input type="text" name="Contatto_DescContatto<%=l_id%>" id="Contatto_DescContatto<%=l_id%>" class="form-control" value="<%=DescContatto%>" >
	  </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>

   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Note</p>
      </div> 
	  <div class="col-8">
	  <input type="text" name="Contatto_NoteContatto<%=l_id%>" id="Contatto_NoteContatto<%=l_id%>" class="form-control" value="<%=NoteContatto%>" >
	  </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>

   
   <%
   CheckSI=""
   CheckNO=""
   if FlagPrincipale="SI" then 
      CheckSI = " checked "
   else
      CheckNO = " checked "
   end if 
   
   %>
   
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Preferito/Principale</p>
      </div> 
	  <div class="col-2">
         <div class="form-check-inline">
              <input name="Contatto_FlagPrincipale<%=l_Id%>" value="SI" type="radio" id="Contatto_FlagPrincipaleSI<%=l_Id%>" 
			  <%=CheckSI%> onclick="localContattoSetFlag('<%=l_Id%>','SI')">
         </div>
		 <span class="font-weight-bold">SI</span>
	  </div>
  
	  <div class="col-2">
         <div class="form-check-inline">
              <input name="Contatto_FlagPrincipale<%=l_Id%>" value="NO" type="radio" id="Contatto_FlagPrincipaleNO<%=l_Id%>" 
			  <%=CheckNO%> onclick="localContattoSetFlag('<%=l_Id%>','NO')">
         </div>
         <span class="font-weight-bold">NO</span>		 
	  </div>	  
	  
      <div class="col-6">
         <p class="font-weight-bold"> </p>
      </div>
   </div>
   <input type="hidden" name="Contatto_ValFlagPrincipale<%=l_Id%>" id="Contatto_ValFlagPrincipale<%=l_Id%>" value="<%=FlagPrincipale%>">

   <%if SoloLettura=false then%>
	<div class="row">
		<div class="mx-auto button-color-click">
		   <div class="center">
		      <button id="btnS_<%=NomeStruttura%>" type="button" onclick="localContattoSubmit('<%=NomeStruttura%>')" class="btn btn-success">Registra</button>
           </div>
		</div>
		<div class="mx-auto button-color-click">
		   <div class="center">
		      <button id="btnR_<%=NomeStruttura%>" type="button" onclick="resetFormContatto('<%=NomeStruttura%>')" class="btn btn-warning">Pulisci</button>
           </div>		
		</div>
		<div id="divformContattoDelete" class="mx-auto button-color-click">
		   <div class="center">
		      <button id="btnD_<%=NomeStruttura%>" type="button" onclick="localContattoRemove('<%=NomeStruttura%>')" class="btn btn-danger">Rimuovi</button>
           </div>		
		</div>		

	</div>
	
   <%end if %>   
  
      </div> 

      <div class="modal-footer">
        <button id="<%NomeStruttura%>_FormCloseButton" type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
      </div>
    </div>
  </div>
</div>   
   