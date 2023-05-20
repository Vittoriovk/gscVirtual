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
  MyContQ = MyContQ & " select * from CauzioneCIG "
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
        id = Rs("IdCauzioneCIG")
        Elenco    = Elenco & Rs("CIG") & " ; "
        Numelenco = NumElenco + 1 
     %>
        <div class="row">
    
               <div class="col-2">
                  <div class="form-group ">
                     <%xx=ShowLabel("CIG")
                       nn=prefix & "CIG" & id
                     %>
                     <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("CIG")%>" >
                  </div>        
               </div>
               <div class="col-8">
                  <div class="form-group ">
                     <%xx=ShowLabel("Descrizione")
                     nn=prefix & "DescCIG" & id
                     %>
                     <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("DescCIG")%>" >
                  </div>        
               </div>
			   
			 <% if readonly="" then %>
             <div class="col-2">
                <div class="form-group ">
                <%xx=ShowLabel("Azioni")%>
                <br>
                <%RiferimentoA=";#;;2;upda;aggiorna;;cig_registra(" & id & ",'');N"%>
                <!--#include virtual="/gscVirtual/include/Anchor.asp"-->                    
                <%RiferimentoA=";#;;2;dele;cancella;;cig_registra(" & id & ",'delete');N"%>
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
          <div class="col-2">
              <div class="form-group ">
                <%xx=ShowLabel("CIG")
                  nn=prefix & "CIG" & id
                %>
                <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="" >
                </div>        
             </div>
             <div class="col-8">
                <div class="form-group ">
                <%xx=ShowLabel("Descrizione")
                  nn=prefix & "DescCIG" & id
                %>
                <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="" >
                </div>        
             </div>
             <div class="col-2">
                <div class="form-group ">
                <%xx=ShowLabel("Azioni")%>
                <br>
                <%RiferimentoA=";#;;2;inse;Inserisci;;cig_registra(0,'');N"%>
                <!--#include virtual="/gscVirtual/include/Anchor.asp"-->    

                </div>        
             </div>                   

      </div> 
 
    <%end if %>

    <input type="hidden" name="elenco_CIG"     id="elenco_CIG"     value = "<%=elenco%>">
    <input type="hidden" name="num_elenco_CIG" id="num_elenco_CIG" value = "<%=numElenco%>">
