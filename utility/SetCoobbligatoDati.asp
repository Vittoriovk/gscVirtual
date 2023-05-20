<!--#include virtual="/gscVirtual/include/includeStd.asp"-->

<%
IdCauzione = request("params_IdCauzione")
azione     = request("params_Azione")
prefix     = "params_"
readonly   = ""

if azione="V" then 
   readonly=" readonly "
end if 
	  Set Rs = Server.CreateObject("ADODB.Recordset")
	  MyContQ = ""
	  MyContQ = MyContQ & " select * from CauzioneCoobbligato "
	  MyContQ = MyContQ & " Where IdCauzione = " & IdCauzione
      MyContQ = MyContQ & " order by 1"  
'response.write MyContQ
      Rs.CursorLocation = 3
      Rs.Open MyContQ, ConnMsde 
      LeggiContatti=true 
	  Conta=0
      If Err.number<>0 then	
       	 LeggiContatti=false
      elseIf Rs.EOF then	
         LeggiContatti=false
		 Rs.close 
      End if
	  Elenco=""
	  NumElenco=0
	  if LeggiContatti then 
	     
	     do while not Rs.eof 
		    conta=conta+1
			checked=""
			if conta=1 then 
			   checked=" checked "
			end if 
			id = Rs("IdCauzioneCoobbligato")
			Elenco    = Elenco & Rs("RagSoc") & " ; "
			Numelenco = NumElenco + 1 
		 %>
			<div class="row">
	
			   <div class="col-4">
                  <div class="form-group ">
				     <%xx=ShowLabel("Nominativo")
					   nn=prefix & "RagSoc" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("RagSoc")%>" >
                  </div>		
			   </div>
			   <div class="col-3">
                  <div class="form-group ">
				     <%xx=ShowLabel("Cod.fiscale")
					 nn=prefix & "CF" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("CF")%>" >
                  </div>		
			   </div>	
			   <div class="col-3">
                  <div class="form-group ">
				     <%xx=ShowLabel("Partita Iva")
					 nn=prefix & "PI" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("PI")%>" >
                  </div>		
			   </div>
			</div>
			<div class="row">
			   <div class="col-4">
                  <div class="form-group ">
				     <%xx=ShowLabel("Indirizzo")
					 nn=prefix & "Indirizzo" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("Indirizzo")%>" >
                  </div>		
			   </div>
			   <div class="col-1">
                  <div class="form-group ">
				     <%xx=ShowLabel("Cap")
					 nn=prefix & "Cap" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("Cap")%>" >
                  </div>		
			   </div>
			   
			   <div class="col-4">
                  <div class="form-group ">
				     <%xx=ShowLabel("Comune")
					 nn=prefix & "Comune" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("Comune")%>" >
                  </div>		
			   </div>
			   <div class="col-2">
                  <div class="form-group ">
				     <%xx=ShowLabel("Provincia")
					 nn=prefix & "Provincia" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("Provincia")%>" >
                  </div>		
			   </div>	
			   <%if azione<>"V" then %>
		     <div class="col-1">
                <div class="form-group ">
			    <%xx=ShowLabel("Azioni")%>
				
                <%RiferimentoA=";#;;2;upda;aggiorna;;coobbligato_registra(" & id & ",'');N"%>
                <!--#include virtual="/gscVirtual/include/Anchor.asp"-->					
                <%RiferimentoA=";#;;2;dele;cancella;;coobbligato_registra(" & id & ",'delete');N"%>
                <!--#include virtual="/gscVirtual/include/Anchor.asp"-->					
				
                </div>		
			 </div>
             <%end if %>			 
			</div>
	 
		 <%
		    Rs.moveNext 
		 loop  
		 Rs.close
	  end if 

	   'modifica ammesso metto rigo per insert 
	   If azione<>"V" then
          Id=0
	   %>

          <div class="row">
             <div class="col-4">
                <div class="form-group ">
                <%xx=ShowLabel("Ragione Sociale")
			      nn=prefix & "RagSoc" & id
			    %>
				<input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="" >
                </div>		
		     </div>
		     <div class="col-3">
                <div class="form-group ">
			    <%xx=ShowLabel("Cod.fiscale")
			      nn=prefix & "CF" & id
			    %>
				<input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="" >
                </div>		
			 </div>	
		     <div class="col-3">
                <div class="form-group ">
			    <%xx=ShowLabel("Part.IVA")
			      nn=prefix & "PI" & id
			    %>
				<input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="" >
                </div>		
			 </div>
          </div>
          <div class="row">		  
			   <div class="col-4">
                  <div class="form-group ">
				     <%xx=ShowLabel("Indirizzo")
					 nn=prefix & "Indirizzo" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="" >
                  </div>		
			   </div>
			   <div class="col-1">
                  <div class="form-group ">
				     <%xx=ShowLabel("Cap")
					 nn=prefix & "Cap" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="" >
                  </div>		
			   </div>
			   
			   <div class="col-4">
                  <div class="form-group ">
				     <%xx=ShowLabel("Comune")
					 nn=prefix & "Comune" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="" >
                  </div>		
			   </div>
			   <div class="col-2">
                  <div class="form-group ">
				     <%xx=ShowLabel("Provincia")
					 nn=prefix & "Provincia" & id
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="" >
                  </div>		
			   </div>	
		     <div class="col-1">
                <div class="form-group ">
			    <%xx=ShowLabel("Azioni")%>
				
                <%RiferimentoA=";#;;2;inse;Inserisci;;coobbligato_registra(0,'');N"%>
                <!--#include virtual="/gscVirtual/include/Anchor.asp"-->	

                </div>		
			 </div>				   

          </div> 
 
	  <%end if %>

	  <input type="hidden" name="elenco_coobbligati"     id="elenco_coobbligati"     value = "<%=elenco%>">
	  <input type="hidden" name="num_elenco_coobbligati" id="num_elenco_coobbligati" value = "<%=numElenco%>">
