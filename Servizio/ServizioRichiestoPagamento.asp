<%
  NomePagina="ServizioRichiestoPagamento.asp"
  titolo="Pagamento Servizio "
  default_check_profile="Coll,Clie,BackO"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionServizioRichiesto.asp"-->
<!--#include virtual="/gscVirtual/common/FunProcedure.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionMovimento.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionEvento.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionPagamentiServizi.asp"-->
<!--#include virtual="/gscVirtual/common/FunMailWithAttach.asp"-->

<!--#include virtual="/gscVirtual/ProcessoElaborativo/FunctionCheckProcesso.asp"-->
<!--#include virtual="/gscVirtual/common/FunCallOtherPage.asp"-->


<%
  livelloPagina="00"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include virtual="/gscVirtual/include/head.asp"-->
<!-- Custom styles for this template -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">

</head>
<script>

function LMPS_CheckPagamentoOk()  {
   xx=myConfirm("Pagamento Servizio","Autorizzazione al pagamento","PAGA");
}
function myConfirmYes()
  {
    var act = $("#myConfirmAction").val();
    ImpostaValoreDi("Oper",act);
    document.Fdati.submit();
   
   
}
function registra(act)
{
    xx=ImpostaValoreDi("DescLoaded","0");
    xx=ElaboraControlli();
    
     if (xx==false)
       return false;

    ImpostaValoreDi("Oper",act);
    document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%
PaginaReturn   = ""
PaginaReturnOk = ""
CMP_FlagDebug = false 
   
   'Attivita' gestite 
   '   FORMAZIONE 
   if FirstLoad then 
      PaginaReturn    = getCurrentValueFor("PaginaReturn")
      PaginaReturnOk  = getCurrentValueFor("PaginaReturnOk")
      IdAttivita      = getCurrentValueFor("IdAttivita")
      IdNumAttivita   = getCurrentValueFor("IdNumAttivita")
   else
      PaginaReturn    = getValueOfDic(Pagedic,"PaginaReturn")
      PaginaReturnOk  = getValueOfDic(Pagedic,"PaginaReturnOk")
      IdAttivita      = getValueOfDic(Pagedic,"IdAttivita")
      IdNumAttivita   = getValueOfDic(Pagedic,"IdNumAttivita")
   end if 
   if PaginaReturnOk="" then 
      PaginaReturnOk=Session("LoginHomePage")
   end if 
   
   IdAttivita    = ucase(Trim(IdAttivita))
   IdNumAttivita = Cdbl("0" & IdNumAttivita)
   'response.write IdAttivita & " " & IdNumAttivita 
   'response.end 

   if IdAttivita="" or cdbl(IdNumAttivita)=0 then 
      response.redirect RitornaA(PaginaReturn)
      response.end 
   end if 

   'recupero il servizio 
   set dizServizio = GetDizServizioRichiesto(0,IdAttivita,IdNumAttivita)
   IdServizioRichiesto  = getDiz(dizServizio,"N_IdServizioRichiesto")
   IdAccountCliente     = getDiz(dizServizio,"N_IdAccountCliente")
   IdAccountRichiedente = getDiz(dizServizio,"N_IdAccountRichiedente")
   IdCompagnia          = getDiz(dizServizio,"N_IdCompagnia") 
   IdFornitore          = getDiz(dizServizio,"N_IdFornitore")
   IdAccountFornitore   = getDiz(dizServizio,"N_IdAccountFornitore")
   IdStatoServizio      = getDiz(dizServizio,"S_IdStatoServizio")
   IdProdotto           = getDiz(dizServizio,"N_IdProdotto")
   TotaleServizio       = getDiz(dizServizio,"N_ImptTotaleServizio")
   TipoPagaClieSelected = getDiz(dizServizio,"S_IdTipoCreditoClie")
   TipoPagaRequSelected = getDiz(dizServizio,"S_IdTipoCreditoRequ")
   
   IdAnagServizio         = getDiz(dizServizio,"S_IdAnagServizio")
   DescAnagServizio       = LeggiCampo("select * from AnagServizio Where IdAnagServizio='" & IdAnagServizio & "'","DescAnagServizio")
   DescProdotto           = LeggiCampo("select * from Prodotto Where IdProdotto=" & IdProdotto,"DescProdotto")
   DescAnagCaratteristica = getDiz(dizServizio,"S_DescAnagCaratteristica")
   DescServizioRichiesto  = getDiz(dizServizio,"S_DescServizioRichiesto")
   
   FlagStatoFinale      = getDiz(dizServizio,"N_FlagStatoFinale")
   if instr("PAGA_PAGC",IdStatoServizio)>0 then 
      canSelModPaga = false
   else 
      canSelModPaga = true 
   end if 
   if canSelModPaga = true and Cdbl("0" & FlagStatoFinale)=1 then 
      canSelModPaga = false
   end if 
   
   prevStato              = IdStatoServizio
   
   IdProcessoElaborativo  = getDiz(dizServizio,"S_IdProcessoElaborativo")

            
   Set Rs = Server.CreateObject("ADODB.Recordset")
   Rs.CursorLocation = 3 
   Rs.Open "Select * from cliente Where IdAccount=" & IdAccountCliente, ConnMsde
   DenomCliente      = Rs("Denominazione")
   CFPI              = Rs("CodiceFiscale") & "/" & Rs("PartitaIva")  
   IdAccountLivello1 = Rs("IdAccountLivello1")
   rs.close   
   
   solaLettura = " readonly "
   
   if Oper="PROSEGUI" then 
      xx=RemoveSwap()
      Session("swap_IdAttivita")    = IdAttivita
      Session("swap_IdNumAttivita") = IdNumAttivita
      Session("swap_PaginaReturn") = "Servizio/" & NomePagina
      response.redirect RitornaA("Servizio/ServizioRichiestoPagamentoGestione.asp")
      response.end    
   end if 
   
   if Oper="PAGA" and Session("TimeStamp")<>TimePage then 
      err.clear
	  IdTipoCreditoClie = Request("ListaModPagServizioCLIE")
	  IdTipoCreditoRequ = Request("ListaModPagServizioREQU")	  
      qUpd     = ""

      qUpd = qUpd & " Update ServizioRichiesto set "
      qUpd = qUpd & " IdTipoCreditoClie='" & IdTipoCreditoClie & "'"
      qUpd = qUpd & ",IdTipoCreditoRequ='" & IdTipoCreditoRequ & "'"
      qUpd = qUpd & " Where IdAttivita='" & apici(IdAttivita) & "'"
      qUpd = qUpd & " and IdNumAttivita=" & NumForDb(IdNumAttivita)
 
      xx=writeTraceAttivita("aggiorno modalita' : " & qUpd,IdAttivita,IdNumAttivita)  
      ConnMsde.execute qUpd
      
      if err.number=0 then 
         ep=pagaServizioRichiesto(IdAttivita,IdNumAttivita)
		 yy=writeTraceAttivita("dopo pagamento : " & ep & " " & err.description ,IdAttivita,IdNumAttivita)  
         if ep<>"" then 
            MsgErrore=ep & " : contattare assistenza."
         else 
		    ' creo la struttura dati 
			xx=creaStrutturaDati(IdServizioRichiesto)
            ' cambio il processo in caso di cauzione provvisoria 
		    if IdAttivita="CAUZ_PROV" then 
			   AtiDaValidare = false
			   qSel = ""
			   qSel = qSel & "select * from Cauzione Where IdCauzione="& NumForDb(IdNumAttivita)
			   'response.write qSel 
			   elencoAti = LeggiCampo(qSel,"ElencoAti")
			   'response.write elencoAti 
			   if elencoAti<>"" then  
			      qSel = ""
				  qSel = qSel & " select top 1 IdAccountAti "
				  qSel = qSel & " from AccountATI "
				  qSel = qSel & " where IdAccountATI in (" & ElencoAti & ")"
				  qSel = qSel & " and IdStatoValidazione<>'AFFI'"
				  'response.write qSel 
				  trov = "0" & LeggiCampo(qSel,"IdAccountAti")
				  'response.write trov
				  if Cdbl(trov)>0 then 
				     AtiDaValidare = true 
				  end if 
			   end if 

               'se ci sono ATI devo controllare se devono essere validate 
			   if AtiDaValidare then 
			      xx=UpdateFlussoProcessoServizioRichiesto(0,IdAttivita,IdNumAttivita,"VAL_ATI")
			   else 
                  xx=UpdateFlussoProcessoServizioRichiesto(0,IdAttivita,IdNumAttivita,"ATT_EMIS")
               end if 
			   'response.end 
			end if 

		    xx=writeTraceAttivita("controllo processo checkProcessoElaborativoServizio " & IdServizioRichiesto ,IdAttivita,IdNumAttivita)  
            xx = checkProcessoElaborativoServizio(IdProdotto,IdAccountFornitore,IdServizioRichiesto)

			xx=RemoveSwap()
            Session("swap_IdAttivita")     = IdAttivita
            Session("swap_IdNumAttivita")  = IdNumAttivita			
            response.redirect "ServizioRichiestoPagamentoOk.asp"
            response.end 
         end if 
      end if 
   end if 
  
   xx=setValueOfDic(Pagedic,"PaginaReturn"      ,PaginaReturn)
   xx=setValueOfDic(Pagedic,"PaginaReturnOk"    ,PaginaReturnOk)
   xx=setValueOfDic(Pagedic,"IdAttivita"        ,IdAttivita)
   xx=setValueOfDic(Pagedic,"IdNumAttivita"     ,IdNumAttivita)   
   xx=setCurrent(NomePagina,livelloPagina) 
   
   NameLoaded= ""

%>

<div class="d-flex" id="wrapper">
    <%
      callP=VirtualPath & "bar/" & Session("sideBar_" & Session("LoginIdAccount")) 
      Server.Execute(callP) 
    %>
    <!-- Page Content -->
    <div id="page-content-wrapper">
    <%
      callP=VirtualPath & "bar/" & Session("TopBar_" & Session("LoginIdAccount")) 
      Server.Execute(callP) 
    %>    

        <div class="container-fluid">
            <form name="Fdati" Action="<%=NomePagina%>" method="post">
 
            <div class="row">
                <%
                if PaginaReturn<>"" then 
                   RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"
                %>
                   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <% end if %>
                <div class="col-11"><h3>Riepilogo Pagamento Servizio</h3>
                </div>
            </div>
            <%
			if isBackOffice() then %>
            <div class="row">
               <div class="col-4">
                  <div class="form-group ">
                     <%
					 tipo=LeggiCampo("select * from Account Where IdAccount=" & IdAccountRichiedente,"IdTipoAccount")
					 if ucase(tipo)="COLL" then 
					    xx=ShowLabel("Richiedente (collaboratore) ")
					 else
					    xx=ShowLabel("Richiedente (cliente) ")
					 end if 
					 DenomRichiedente=LeggiCampo("select * from Account Where IdAccount=" & IdAccountRichiedente,"Nominativo")
					 %>
                     <input type="text" readonly class="form-control" value="<%=DenomRichiedente%>" >
                  </div>        
               </div>
            </div>
			<%
			end if 
			
            FlagAddInter=1
            
            if isCliente()=false then 
               FlagAddInter = 0 
               
            %>
            <div class="row">
               <div class="col-4">
                  <div class="form-group ">
                     <%xx=ShowLabel("Contraente")%>
                     <input type="text" readonly class="form-control" value="<%=DenomCliente%>" >
                  </div>        
               </div>
               <div class="col-4">
                  <div class="form-group ">
                     <%xx=ShowLabel("Cod.fiscale/PartitaIva")%>
                     <input type="text" readonly class="form-control" value="<%=CFPI%>" >
                  </div>        
               </div>               
            </div>
            <%end if %>
  
            <div class="row">
               <div class="col-4">
                  <div class="form-group ">
                     <%xx=ShowLabel("Tipo Servizio ")%>
                     <input type="text" readonly class="form-control" value="<%=DescAnagServizio%>" >
                  </div>        
               </div>
               <div class="col-4">
                  <div class="form-group ">
                     <%
                     xx=ShowLabel("Prodotto")
                     
                     %>
                     <input type="text" readonly class="form-control"  
                     value="<%=DescProdotto%>" >
                  </div>
                </div>
                <%if DescAnagCaratteristica<>"" then %>
               <div class="col-4">
                  <div class="form-group ">
                     <%
                     xx=ShowLabel("Dettaglio Prodotto")
                     
                     %>
                     <input type="text" readonly class="form-control"  
                     value="<%=DescAnagCaratteristica%>" >
                  </div>
                </div>
                
                <%end if %>
               </div>
               <div class="row">
                <div class="col-9">
                  <div class="form-group ">
                     <%
                     xx=ShowLabel("Descrizione Servizio Attivato")
 
                     %>
                     <input type="text" Readonly class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
                     value="<%=DescServizioRichiesto%>" >
                  </div>
                </div>
               <div class="col-3">
                  <div class="form-group ">
                     <%xx=ShowLabel("Stato")
                     DescStato = GetDiz(dizServizio,"S_DescStatoServizio")
                     
                     %>
                     <input type="text" readonly class="form-control" value="<%=DescStato%>" >
                  </div>        
               </div>				
            </div>

            <div class="row">

  
               <div class="col-2">
                  <div class="form-group ">
                     <%
                     xx=ShowLabel("Totale da Pagare")
                 
                     %>
                     <input type="text" readonly class="form-control text-success text-center font-weight-bold"
                     value="<%=InsertPoint(TotaleServizio,2)%> &euro;" >
                  </div>    
               </div>             
          
   
            </div>
            <br>
            <%
              
            ImptDaPagare           = cdbl(TotaleServizio)
            IdAccountModPagCliente = IdAccountCliente
			if IsBackOffice() then 
			   IdAccountRequest = IdAccountRichiedente
			else
               IdAccountRequest = Session("LoginIdAccount")
			end if 
            'sono recuperati in precedenza da ServizioRichiesto  
            'IdProcessoElaborativo  = processo elaborativo: default STANDARD
            
            %>
            <!--#include virtual="/gscVirtual/configurazioni/pagamenti/ListaModPagServizioRichiesto.asp"-->
            

            <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

            <input type="hidden" name="localVirtualPath" id="localVirtualPath" value = "<%=VirtualPath%>">
            <input type="hidden" name="IdAccountCliente" id="IdAccountCliente" value = "<%=IdAccountCliente%>">
     <%if false and isBackOffice() and instr("PAGA_PAGC",IdStatoServizio)>0 then%>
	    <div class="row">
           <div class="mx-auto">
              <button type="button" onclick="registra('prosegui')" class="btn btn-success">Prosegui</button>
           </div>   
     </div>
      <%end if %> 
 
            <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
            <!--#include virtual="/gscVirtual/include/paginazione.asp"-->

   
        <div class="row">
            <div class="col">
                &nbsp;
            </div>
        </div>
            
            </form>
        </div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->

<!--#include virtual="/gscVirtual/include/scripts.asp"-->

  <!-- Menu Toggle Script -->
  <script>
    $("#menu-toggle").click(function(e) {
      e.preventDefault();
      $("#wrapper").toggleClass("toggled");
    });
  </script>
  <script>
    $(document).ready(function(){
      $('[data-toggle="tooltip"]').tooltip();   
    });
  </script>
  <script>
$('.btn').onClick(function(e){
  e.preventDefault();
});  
</script>
</body>

</html>
