<%
  NomePagina="ServizioRichiestoPagamentoGestione.asp"
  titolo="Gestione Pagamento Servizio "
  default_check_profile="BackO"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionServizioRichiesto.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionPagamentiServizi.asp"-->
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
function registra(act)  {
var descOper="";
   if (act=="ANPAG")
      descOper="Annullamento pagamento ";
   if (act=="ANSER")
      descOper="Annullamento servizio";
   if (act=="CONFE")
      descOper="Conferma Pagamento";
	  
   xx=myConfirm(descOper,"Conferma operazione",act);
}
function myConfirmYes()
  {
    var act = $("#myConfirmAction").val();
    ImpostaValoreDi("Oper",act);
    document.Fdati.submit();
  
}
function salva(act)
  {
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
   
   'Attivita' gestite 
   '   FORMAZIONE 
   if FirstLoad then 
      PaginaReturn    = getCurrentValueFor("PaginaReturn")
      IdAttivita      = getCurrentValueFor("IdAttivita")
      IdNumAttivita   = getCurrentValueFor("IdNumAttivita")
   else
      PaginaReturn    = getValueOfDic(Pagedic,"PaginaReturn")
      IdAttivita      = getValueOfDic(Pagedic,"IdAttivita")
      IdNumAttivita   = getValueOfDic(Pagedic,"IdNumAttivita")
   end if 
   
   IdAttivita    = ucase(Trim(IdAttivita))
   IdNumAttivita = Cdbl("0" & IdNumAttivita)

   'recupero il servizio 
   set dizServizio = GetDizServizioRichiesto(0,IdAttivita,IdNumAttivita)
   IdServizioRichiesto  = getDiz(dizServizio,"N_IdServizioRichiesto")
   IdAccountCliente     = getDiz(dizServizio,"N_IdAccountCliente")
   IdAccountRichiedente = getDiz(dizServizio,"N_IdAccountRichiedente")
   IdCompagnia          = getDiz(dizServizio,"N_IdCompagnia") 
   IdFornitore          = getDiz(dizServizio,"N_IdFornitore")
   IdAccountFornitore   = getDiz(dizServizio,"N_IdAccountFornitore")
   IdStatoServizio      = getDiz(dizServizio,"S_IdStatoServizio")
   IdStatoServizioPrec  = getDiz(dizServizio,"S_IdStatoServizioPrec")
   DescStatoServizio    = GetDiz(dizServizio,"S_DescStatoServizio")
   IdProdotto           = getDiz(dizServizio,"N_IdProdotto")
   TotaleServizio       = getDiz(dizServizio,"N_ImptTotaleServizio")
   TipoPagaClieSelected = getDiz(dizServizio,"S_IdTipoCreditoClie")
   TipoPagaRequSelected = getDiz(dizServizio,"S_IdTipoCreditoRequ")
   
   IdAnagServizio         = getDiz(dizServizio,"S_IdAnagServizio")
   DescAnagServizio       = LeggiCampo("select * from AnagServizio Where IdAnagServizio='" & IdAnagServizio & "'","DescAnagServizio")
   DescProdotto           = LeggiCampo("select * from Prodotto Where IdProdotto=" & IdProdotto,"DescProdotto")
   DescAnagCaratteristica = getDiz(dizServizio,"S_DescAnagCaratteristica")
   DescServizioRichiesto  = getDiz(dizServizio,"S_DescServizioRichiesto")
   
   NoteStatoServizio      = getDiz(dizServizio,"S_NoteStatoServizio")
   NoteServizio           = getDiz(dizServizio,"S_NoteServizio")
   NoteServizioCliente    = getDiz(dizServizio,"S_NoteServizioCliente")
   if oper<>"" and instr("ANPAG_REGISTRA_CONFE",oper)>0 then 
      NoteServizio           = Request("NoteServizio0")
      NoteServizioCliente    = Request("NoteServizioCliente0")   
      xx = UpdateNoteServizioRichiesto(IdAttivita,IdNumAttivita,NoteServizio,NoteServizioCliente)
	  
	  'devo stornare i pagamenti
	  if Oper="ANPAG" then 
	     xx = stornaPagamenti(IdServizioRichiesto)
         IdStatoServizio = IdStatoServizioPrec
	  end if 
	  if Oper="CONFE" then 
	     IdStatoServizio = "PAGA"
		 DescStatoServizio = "Pagamento"
	  end if 
      NoteStatoServizio  = Request("NoteStatoServizio0")
      xx = UpdateStatoServizioRichiesto(IdAttivita,IdNumAttivita,IdStatoServizio,NoteStatoServizio)
	  if IdStatoServizio = "ANNU" then 
	     MsgErrore = "Servizio annullato"
	  end if 
   end if 

   Set Rs = Server.CreateObject("ADODB.Recordset")
   Rs.CursorLocation = 3 
   Rs.Open "Select * from cliente Where IdAccount=" & IdAccountCliente, ConnMsde
   DenomCliente      = Rs("Denominazione")
   CFPI              = Rs("CodiceFiscale") & "/" & Rs("PartitaIva")  
   IdAccountLivello1 = Rs("IdAccountLivello1")
   rs.close   
   
   solaLettura = "  "
   if Instr("ANNU",IdStatoServizio)>0 then 
      solaLettura = " readonly "
   end if 
   
   xx=setValueOfDic(Pagedic,"PaginaReturn"      ,PaginaReturn)
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
                <div class="col-11"><h3>Gestione Pagamento Servizio</h3>
                </div>
            </div>
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
                <div class="col-8">
                  <div class="form-group ">
                     <%
                     xx=ShowLabel("Descrizione Servizio Attivato")
 
                     %>
                     <input type="text" Readonly class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
                     value="<%=DescServizioRichiesto%>" >
                  </div>
                </div>

            </div>

            <div class="row">
               <div class="col-3">
                  <div class="form-group ">
                     <%xx=ShowLabel("Stato")
                     %>
                     <input type="text" readonly class="form-control text-info font-weight-bold" value="<%=DescStatoServizio%>" >
                  </div>        
               </div>
  
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
			<div class="row">
			   <div class="col-11">
                  <div class="form-group ">
				     <%xx=ShowLabel("Motivo variazione")
                       idDaPulire ="NoteStatoServizio0"
					   idDaSalvare=""
					   if solaLettura<>"" then 
					      idDaPulire=""
					   end if 
					 %>
					 <!--#include virtual="/gscVirtual/include/pulisciSalva.asp"-->					 
					 <input type="text" <%=solaLettura%> id="NoteStatoServizio0" name="NoteStatoServizio0" class="form-control" value="<%=NoteStatoServizio%>" >
                  </div>		
			   </div>
			</div>			
			<div class="row">
			   <div class="col-11">
                  <div class="form-group ">
				     <%xx=ShowLabel("Annotazioni interne")
                       idDaPulire ="NoteServizio0"
					   idDaSalvare=""
					   if solaLettura<>"" then 
					      idDaPulire=""
					   end if 
					 %>
					 <!--#include virtual="/gscVirtual/include/pulisciSalva.asp"-->					 
					 <input type="text" <%=solaLettura%> id="NoteServizio0" name="NoteServizio0" class="form-control" value="<%=NoteServizio%>" >
                  </div>		
			   </div>
			</div>
	  
			<div class="row">
			   <div class="col-11">
                  <div class="form-group ">
				     <%xx=ShowLabel("Annotazioni per il cliente")
                       idDaPulire ="NoteServizioCliente0"
					   idDaSalvare=""
					   if solaLettura<>"" then 
					      idDaPulire=""
					   end if 
					 %>
					 <!--#include virtual="/gscVirtual/include/pulisciSalva.asp"-->
					 <input type="text" <%=solaLettura%> id="NoteServizioCliente0" name="NoteServizioCliente0" class="form-control" value="<%=NoteServizioCliente%>" >
                  </div>		
			   </div>
 
			</div>				
            <br>
          <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
		   <%if IdStatoServizio <> "ANNU" then%>
	      <div class="row">
		     <div class="mx-auto">
		         <button type="button" onclick="salva('REGISTRA')" class="btn btn-success">Registra</button>
		     </div>
			 
			 <%if IdStatoServizio = "PAGC" or IdStatoServizio ="PAGA" then%>
		     <div class="mx-auto">
		         <button type="button" onclick="registra('ANPAG')" class="btn btn-warning">Annulla Pagamento</button>
		     </div>
			 <%end if %>

             <%if IdStatoServizio ="PAGC" then%>
		     <div class="mx-auto">
		         <button type="button" onclick="registra('CONFE')" class="btn btn-success">Conferma Pagamento</button>
             </div>
		     <%end if %>
   
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
