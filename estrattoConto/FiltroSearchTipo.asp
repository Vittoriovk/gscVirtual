            <%
			   IdTipoCredito = Request("ListaModPag")
               FlagBorsellino = 1
               FlagFido       = 1
               FlagEstratto   = 1
               FlagAction = true 

                   flagchecked   = IdtipoCredito
				   IdtipoCredito = ""
                   BorsSele=""
                   FidoSele=""
                   EstrSele=""
				   IdtipoCredito=""
				   if Flagchecked="BORS" and flagBorsellino=1 then
				      BorsSele= " checked "
					  IdtipoCredito= Flagchecked
				   end if 
				   if Flagchecked="FIDO" and flagFido=1 then
				      FidoSele= " checked "
					  IdtipoCredito= Flagchecked
				   end if 
				   if Flagchecked="ESTR" and flagEstratto=1 then
				      EstrSele= " checked "
					  IdtipoCredito= Flagchecked
				   end if 
                   if IdtipoCredito="" then 
					   if flagBorsellino=1 then
						  BorsSele= " checked "
						  IdtipoCredito= "BORS"
					   elseif flagFido=1 then
						  FidoSele= " checked "
						  IdtipoCredito= "FIDO"
					   else  
						  EstrSele= " checked "
						  IdtipoCredito= "ESTR"
					   end if 
				   end if 

			   %> 				

			      <div class="row">
				     <div class="col-2">
					 <p class="font-weight-bold">Tipo Estratto </p>
					 </div>
			      <% if  FlagBorsellino=1 then %>  
                     
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=BorsSele%> type="radio" 
                     	   name="ListaModPag" id="ListaModPagBORS" value="BORS"
						   onclick="localRicarica();" >				  
 					       Borsellino
					    </div>
                     </div>		
			      <% end if %>
			      <% if  FlagFido=1 then %>  
                     
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=FidoSele%> type="radio" 
                     	   name="ListaModPag" id="ListaModPagFIDO" value="FIDO"
						   onclick="localRicarica();">				  
 					       Fido
					    </div>
                     </div>		
			      <% end if %>				  
			      <% if  FlagEstratto=1 then %>  
                     
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=EstrSele%> type="radio" 
                     	   name="ListaModPag" id="ListaModPagESTR" value="ESTR"
						   onclick="localRicarica();">				  
 					       Estratto
					    </div>
                     </div>		
			      <% end if %>					  
			      </div>			

			<%
			AddRow=true
			dim CampoDb(10)
			ElencoOption = ";0;Cliente;1"
            CampoDB(1)   = "Denominazione"
						
			%>
		<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
    
