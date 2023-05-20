<%
  NomePagina="ServizioRichiestoPagamentoOk.asp"
  titolo="Pagamento Servizio "
  default_check_profile="Coll,Clie,BackO"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionServizioRichiesto.asp"-->

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
      IdAttivita      = getCurrentValueFor("IdAttivita")
      IdNumAttivita   = getCurrentValueFor("IdNumAttivita")
   else
      IdAttivita      = getValueOfDic(Pagedic,"IdAttivita")
      IdNumAttivita   = getValueOfDic(Pagedic,"IdNumAttivita")
   end if 
   
   IdAttivita    = ucase(Trim(IdAttivita))
   IdNumAttivita = Cdbl("0" & IdNumAttivita)

   'recupero il servizio 
   set dizServizio = GetDizServizioRichiesto(0,IdAttivita,IdNumAttivita)
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
            
   Set Rs = Server.CreateObject("ADODB.Recordset")
   Rs.CursorLocation = 3 
   Rs.Open "Select * from cliente Where IdAccount=" & IdAccountCliente, ConnMsde
   DenomCliente      = Rs("Denominazione")
   CFPI              = Rs("CodiceFiscale") & "/" & Rs("PartitaIva")  
   IdAccountLivello1 = Rs("IdAccountLivello1")
   rs.close   
   
   solaLettura = " readonly "
   
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
                <div class="col-11"><h3>Pagamento Servizio eseguito </h3>
                </div>
            </div>
            <%
          
            if isCliente()=false then 
               
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
