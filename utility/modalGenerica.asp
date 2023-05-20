<!--#include virtual="/gscVirtual/include/includeStd.asp"-->

<%
IdTipoRichiesta=ucase(Trim(Request("IdTipoRichiesta")))

If IdTipoRichiesta="SEL_INDIRIZZO" then 
   IdAccount   = cdbl("0" & TestNumeroPos("0" & Request("IdAccount")))
   ListaCampi  = Request("ListaCampi")
   if IdAccount=0 then 
      IdTipoRichiesta=""
   end if 
elseIf IdTipoRichiesta="SEL_PEC" then 
   IdAccount   = cdbl("0" & TestNumeroPos("0" & Request("IdAccount")))
   ListaCampi  = Request("ListaCampi")
   if IdAccount=0 then 
      IdTipoRichiesta=""
   end if 
elseIf IdTipoRichiesta="SET_COOBBLIGATI" then 
   IdCauzione  = cdbl("0" & TestNumeroPos("0" & Request("IdCauzione")))
   Azione      = Request("azione")
   ListaCampi  = Request("ListaCampi")
   if IdCauzione=0 then 
      IdTipoRichiesta=""
   end if    
elseIf IdTipoRichiesta="SET_ATI" then 
   IdCauzione  = cdbl("0" & TestNumeroPos("0" & Request("IdCauzione")))
   Azione      = Request("azione")
   ListaCampi  = Request("ListaCampi")
   if IdCauzione=0 then 
      IdTipoRichiesta=""
   end if     
elseIf IdTipoRichiesta="SET_CIG" then 
   IdCauzione  = cdbl("0" & TestNumeroPos("0" & Request("IdCauzione")))
   Azione      = Request("azione")
   ListaCampi  = Request("ListaCampi")
   if IdCauzione=0 then 
      IdTipoRichiesta=""
   end if   
else
   IdTipoRichiesta=""
end if 

If IdTipoRichiesta="" then 
   response.end
end if 

SubTitolo=""
If IdTipoRichiesta="SEL_INDIRIZZO" then 
   Titolo="Selezione indirizzo"
end if 
If IdTipoRichiesta="SEL_PEC" then 
   Titolo="Selezione PEC"
end if 
If IdTipoRichiesta="SET_COOBBLIGATI" then 
   Titolo="Elenco Coobbligati"
end if
If IdTipoRichiesta="SET_ATI" then 
   Titolo="Elenco ATI"
   SubTitolo="Assicurarsi che le aziende abbiano le certificazioni necessarie."
end if
If IdTipoRichiesta="SET_CIG" then 
   Titolo="Elenco altri Lotti di partecipazione "
end if
%>


<div class="modal fade" id="confirmModalGenerica"  aria-hidden="true" role="dialog">
  <!-- a pieno schermo -->
  <div class="modal-dialog modal-xl">
    <div class="modal-content">
        <div class="row">
			<div class="col-10"><h4>&nbsp;&nbsp;<%=Titolo%></h4></div>
			<div class="col-2">
                 <button type="button" Id="dismissModalGenerica" class="close" data-dismiss="modal">
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
		<div>
	  <%
	  session("EsitoCallSelIndirizzo")="KO"
	  If IdTipoRichiesta="SEL_INDIRIZZO" then 
	     session("params_IdAccount")  = IdAccount
		 Session("params_ListaCampi") = ListaCampi
         session("EsitoCallSelIndirizzo") = "" 
	     callP=VirtualPath & "utility/SelIndirizzo.asp"
         Server.Execute(callP) 
	  end if 
	  If IdTipoRichiesta="SEL_PEC" then 
	     session("params_IdAccount")  = IdAccount
		 Session("params_ListaCampi") = ListaCampi
         session("EsitoCallSelIndirizzo") = "" 
	     callP=VirtualPath & "utility/SelPec.asp"
         Server.Execute(callP) 
	  end if 	  
	  If IdTipoRichiesta="SET_COOBBLIGATI" then 
	     session("params_IdCauzione") = IdCauzione
		 Session("params_Azione")     = Azione
		 Session("params_ListaCampi") = ListaCampi
         session("EsitoCallSelIndirizzo") = "" 
	     callP=VirtualPath & "utility/SetCoobbligato.asp"
         Server.Execute(callP) 
	  end if 	  
	  If IdTipoRichiesta="SET_ATI" then 
	     session("params_IdCauzione") = IdCauzione
		 Session("params_Azione")     = Azione
		 Session("params_ListaCampi") = ListaCampi
         session("EsitoCallSelIndirizzo") = "" 
	     callP=VirtualPath & "utility/SetATI.asp"
         Server.Execute(callP) 
	  end if 	  
	  If IdTipoRichiesta="SET_CIG" then 
	     session("params_IdCauzione") = IdCauzione
		 Session("params_Azione")     = Azione
		 Session("params_ListaCampi") = ListaCampi
         session("EsitoCallSelIndirizzo") = "" 
	     callP=VirtualPath & "utility/SetCIG.asp"
         Server.Execute(callP) 
	  end if 	 	  
	  %>

		</div>		  
      </div> 
	  <%
		 if session("EsitoCallSelIndirizzo")="" then 
      %>
      <div class="modal-footer">
         <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
		 <!-- viene scritta dalla call precedente -->
         <button type="button" class="btn btn-primary" onclick="confirmModalProcedi();";>Procedi</button>
      </div>
	  <%
	     end if 
      %>
    </div>
  </div>
</div>