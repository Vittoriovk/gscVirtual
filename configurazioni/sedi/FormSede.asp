<!-- Nome del div da mostrare -->
<div class="modal fade" id="<%=NomeStruttura%>Modal" role="dialog" aria-hidden="true">
  <div class="modal-dialog modal-lg">
    <div class="modal-content">
      <div class="modal-header">

	    <!-- Descrizione dell'oggetto -->
        <h2>Gestione Indirizzo</h2> 
      </div>

      <div class="modal-body"> 
   <!-- indice della riga in elaborazione : deve esserci sempre  -->
   <input type="hidden" name="<%=NomeStruttura%>_Idx0"  id="<%=NomeStruttura%>_Idx0" value="">
   <!-- identificativo account  -->
   <input type="hidden" name="Sede_IdAccount0"          id="Sede_IdAccount0" value="<%=IdAccount%>">   
   <!-- identificativo assoluto della riga sul db : se non esiste deve essere inserito  -->
   <input type="hidden" name="Sede_IdAccountSede0"      id="Sede_IdAccountSede0" value="">	  
   <!-- elenco dei campi della struttura legati a liste : necessari per l'elenco -->
   <!-- per ognuno mettere le opportune funzioni di popolamento al cambio della option -->
   <input type="hidden" name="Sede_DescTipoSede0"       id="Sede_DescTipoSede0" value="">
   <input type="hidden" name="Sede_DescStato0"          id="Sede_DescStato0"    value="">
   <input type="hidden" name="Sede_Azioni0"             id="Sede_Azioni0"       value="">
   <%
   l_Id = "0"

   if IdTipoSede="" then 
      IdTipoSede="ALTR"
   end if    
   if IdStato="" then 
      IdStato="IT"
   end if 
   
   ao_Eve = ""                        'azzero evento
   ao_Att = "0"                       'indica se deve mettere vuoto 
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Tipo Sede</p>
      </div>   
   
      <div class = "col-3">
	     <%
		 qSel = ""
		 qSel = qSel &  " SELECT * From TipoSede "
		 if ProfiloAccount<>"" then 
		    qSel = qSel & " Where AddInfo not like '%NO_" & ProfiloAccount & "%' "
		 end if 
		 qSel = qSel & " order By DescTipoSede "
		 ao_Eve = "localSedeChangeTipoSede('" & l_id & "')"
	     response.write ListaDbChangeCompleta(qSel,"Sede_IdTipoSede" & l_Id,IdTipoSede  ,"IdTipoSede","DescTipoSede" ,ao_Att   ,ao_Eve,""   ,""       ,""   ,""    ,ao_cla)
	     %>
      </div>

      <div class="col-1">
         <p class="font-weight-bold">Stato</p>
      </div>
      <div class = "col-6">
	     <%
		 ao_Att = "0"
		 ao_Eve = "localSedeChangeStato('" & l_id & "')"
	     response.write ListaDbChangeCompleta("SELECT * From Stato order By DescStato","Sede_IdStato" & l_Id,IdStato  ,"IdStato","DescStato" ,ao_Att   ,ao_Eve,""   ,""       ,""   ,""    ,ao_cla)
	     %>
      </div>  
   </div>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Provincia</p>
      </div>
      <div class = "col-7">
	     <div id="divProvIT<%=l_id%>" style='display: block;'>
	     <%
         ao_Att = "0"                       'indica se deve mettere vuoto 
         ao_Cla = "class='form-control form-control-sm'"
         ao_Eve = "localSedeChangeProvincia('" & l_id & "')"		 
	     response.write ListaDbChangeCompleta("SELECT * From Provincia order By DescProvincia","Sede_ProvinciaIT" & l_Id,Provincia  ,"IdProvincia","DescProvincia" ,ao_Att   ,ao_Eve,""   ,""       ,""   ,""    ,ao_cla)
	     %>
	     </div>  
	  
	     <div id="divProv<%=l_id%>"   style='display: block;' >
            <input type="text" name="Sede_Provincia<%=l_id%>" id="Sede_Provincia<%=l_id%>" class="form-control" value="" >
		 </div>
      </div>
      <div class="col-1">
         <p class="font-weight-bold">Cap</p>
      </div> 
      <div class="col-2">
         <input type="text" name="Sede_Cap<%=l_id%>" id="Sede_Cap<%=l_id%>" class="form-control" value="" >
      </div> 	
  
   </div>	    
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Comune</p>
      </div> 
	  <div class="col-10">
	  <input type="text" list="dataList_ComuneSede" name="Sede_Comune<%=l_id%>" id="Sede_Comune<%=l_id%>" class="form-control" value="" >
	  </div>
   </div>
   
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Indirizzo</p>
      </div> 
	  <div class="col-7">
	  <input type="text" name="Sede_Indirizzo<%=l_id%>" id="Sede_Indirizzo<%=l_id%>" class="form-control" value="" >
	  </div>
      <div class="col-1">
         <p class="font-weight-bold">Civico</p>
      </div> 
      <div class="col-2">
         <input type="text" name="Sede_Civico<%=l_id%>" id="Sede_Civico<%=l_id%>" class="form-control" value="" >
      </div> 	  
   </div>

   
   <%if SoloLettura=false then%>
	<div class="row">
		<div class="mx-auto button-color-click">
		   <div class="center">
		      <button id="btnS_<%=NomeStruttura%>" type="button" onclick="localSedeSubmit('<%=NomeStruttura%>')" class="btn btn-success">Registra</button>
           </div>
		</div>
		<div class="mx-auto button-color-click">
		   <div class="center">
		      <button id="btnR_<%=NomeStruttura%>" type="button" onclick="resetFormSede('<%=NomeStruttura%>')" class="btn btn-warning">Pulisci</button>
           </div>		
		</div>
		<div id="divformSedeDelete" class="mx-auto button-color-click">
		   <div class="center">
		      <button id="btnD_<%=NomeStruttura%>" type="button" onclick="localSedeRemove('<%=NomeStruttura%>')" class="btn btn-danger">Rimuovi</button>
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