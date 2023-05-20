   <!-- riceve in input IdTipoStato , se non impostato si mette a TUTTI -->
   <%
   IdTipoStato = Trim(Request("statoRichiesta"))
   if IdTipoStato = "" then 
      IdTipoStato = "TUTTI"
   end if 
   rowOpen = true 
   conta   = 4
   %>
   <div class="row ">
      <div class="col-2"><b>Stato Richieste</B></div>
      <div class="col-2">
           <div class="form-group ">
               <%
               if IdTipoStato = "TUTTI" then 
                  checked = " checked "
               else
                  checked = ""
               end if 
               
               %>
               <input class="form-check-input" type="radio" <%=checked%> onclick="Sottometti();"
               name="statoRichiesta" id="statoTUTTI" value="TUTTI">Tutte  
           </div>
      </div>
   <!-- riceve in input Tipo Utente da session -->
   <%
   LTU = Session("LoginTipoUtente")
   qSel = "" 
   qSel = qSel & " select distinct IdTipoStato,DescTipoStato,ordine" 
   qSel = qSel & " from Statoflusso"
   qSel = qSel & " Where ordine > 0 "
   qSel = qSel & " and (IdTipoUtente='*' or IdTipoUtente like '%" & LTU & "%')"
   qSel = qSel & " order by Ordine" 
   
   Set rsFL = Server.CreateObject("ADODB.Recordset")
   RsFL.CursorLocation = 3 
   RsFL.Open qSel, ConnMsde
   Do While Not RsFL.EOF
      conta = conta + 2
      if conta > 12 then %>
	     </div>
         <div class="row ">
            <div class="col-2"></div>  
      <%  
	     conta = 4
      end if
	  if IdTipoStato = RsFL("IdTipoStato") then 
         checked = " checked "
      else
         checked = ""
      end if 
	  statoExp = RsFL("IdTipoStato")
	  StatoDes = RsFL("DescTipoStato")
	  %>
      <div class="col-2">
           <div class="form-group ">
               <input class="form-check-input" type="radio" <%=checked%> onclick="Sottometti();"
               name="statoRichiesta" id="stato<%=statoExp%>" value="<%=statoExp%>"><%=statoDes%>
           </div>
      </div>	  
	  <%
      RsFL.MoveNext 
   loop 
   %>  
   
   <%if rowOpen then %>  
   </div>
   <%end if %>
