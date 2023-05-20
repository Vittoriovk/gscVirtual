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
  
  'recupero cliente 
  MyContQ = ""
  MyContQ = MyContQ & " select * from ServizioRichiesto "
  MyContQ = MyContQ & " Where IdAttivita = 'CAUZ_PROV'  And IdNumAttivita = " & IdCauzione
  
  IdAccountCliente = LeggiCampo(MyContQ,"IdAccountCliente")
  
  MyContQ = ""
  MyContQ = MyContQ & " select A.*,IsNull(b.IdAccountATI,0) as checkAti "
  MyContQ = MyContQ & " from AccountAti A "
  MyContQ = MyContQ & " left join CauzioneATI B"
  MyContQ = MyContQ & " on A.IdAccountATI = b.IdAccountATI"
  MyContQ = MyContQ & " Where A.IdAccount = " & IdAccountCliente
  MyContQ = MyContQ & " and (A.IdAccountATI = b.IdAccountATI or "
  MyContQ = MyContQ & "     (A.ValidoDal<=" & DTos() & " and ValidoAl>=" & DTos() & ")"
  MyContQ = MyContQ & "     )"
  MyContQ = MyContQ & " order By RagSoc "


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
        if cdbl(rs("checkAti"))>0 then 
           checked=" checked "
           Elenco    = Elenco & Rs("RagSoc") & " ; "
           Numelenco = NumElenco + 1 		   
        end if 
        id = Rs("IdAccountATI")

     %>
        <div class="row">
               <div class="col-1">
                  <div class="form-group ">
                     <%xx=ShowLabel("Sel.")%><br>
                     <input type="checkbox" class="big-checkbox" <%=checked%> disabled >
                  </div>        
               </div>    
               <div class="col-5">
                  <div class="form-group ">
                     <%xx=ShowLabel("Ragione Sociale")%>
                     <input type="text" readonly class="form-control" value="<%=Rs("RagSoc")%>" >
                  </div>        
               </div>
               <div class="col-3">
                  <div class="form-group ">
                     <%xx=ShowLabel("Partita Iva")%>
                     <input type="text" readonly class="form-control" value="<%=Rs("PI")%>" >
                  </div>        
               </div>
			 <% if readonly="" then %>
             <div class="col-2">
                <div class="form-group ">
                <%xx=ShowLabel("Azioni")%>
                <br>
				<%
				if cdbl(rs("checkAti"))=0 then
                   RiferimentoA=";#;;2;inse;aggiorna;;ati_registra(" & id & ",'insert');N"
				else
				   RiferimentoA=";#;;2;dele;aggiorna;;ati_registra(" & id & ",'delete');N"
				end if 
				%>
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
%>

    <input type="hidden" name="elenco_ATI"     id="elenco_ATI"     value = "<%=elenco%>">
    <input type="hidden" name="num_elenco_ATI" id="num_elenco_ATI" value = "<%=numElenco%>">
