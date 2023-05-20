<%
  NomePagina="NuovaRichiesta.asp"
  titolo="Menu - Nuova Richiesta Di Affidamento Cliente"
  default_check_profile="Coll,Clie,BackO"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAffidamento.asp"-->
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
function localIns()
{
   xx=ImpostaValoreDi("Oper","CALL_INS");
   document.Fdati.submit();  
}
function localNext()
{
   xx=ImpostaValoreDi("Oper","CALL_NEXT");
   document.Fdati.submit();  
}
</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%
PaginaReturn = ""
Oggi = Dtos() 
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Ds = Server.CreateObject("ADODB.Recordset")

if FirstLoad then 
   opzioneSidebar = getCurrentValueFor("opzioneSidebar")
   PaginaReturn   = getCurrentValueFor("PaginaReturn")
   if isCliente() then 
      IdCliente = Session("LoginIdCliente") 
      IdAccount = Session("LoginIdAccount")
   else
      IdCliente = getCurrentValueFor("IdCliente")  
      IdAccount = LeggiCampo("select IdAccount From Cliente Where IdCliente=" & IdCliente,"IdAccount")
   end if 
else
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
   IdCliente      = getValueOfDic(Pagedic,"IdCliente")
   IdAccount      = getValueOfDic(Pagedic,"IdAccountCliente")
   opzioneSidebar = getValueOfDic(Pagedic,"opzioneSidebar")
end if 

if opzioneSidebar="" then 
   opzioneSidebar="affi"
end if 
if PaginaReturn="" and Session("LoginTipoUtente")=ucase("Clie") then 
   PaginaReturn="link/ClienteAffidamento.asp"
end if 
if Session("LoginTipoUtente")=ucase("Clie") then 
   IdAccount = Session("LoginIdAccount")
end if 

xx=setValueOfDic(Pagedic,"PaginaReturn"      ,PaginaReturn)
xx=setValueOfDic(Pagedic,"IdCliente"         ,IdCliente)
xx=setValueOfDic(Pagedic,"IdAccountCliente"  ,IdAccount)
xx=setValueOfDic(Pagedic,"opzioneSidebar"    ,opzioneSidebar)

xx=setCurrent(NomePagina,livelloPagina) 

if Oper="CALL_NEXT" then 
   xx=RemoveSwap()
   Session("swap_PaginaReturn")           = PaginaReturn
   Session("swap_TipoRicercaExt")         = "1"
   qKey = "select * from Cliente Where IdCliente=" & IdCliente
   descCliente = LeggiCampo(qKey,"Denominazione")
   Session("swap_testo_ricercaExt")       = descCliente 
   response.redirect RitornaA("configurazioni/Clienti/Affidamento/ListaRichiesta.asp")
   Response.end 
end if 

if Oper="CALL_INS" then 
   
   'carico la richiesta di affidamento 
   MyQ = "" 
   MyQ = MyQ & "insert into AffidamentoRichiesta (IdAccountRichiedente,IdAccountCliente,DataRichiesta,IdStatoAffidamento)"
   MyQ = MyQ & " values (" & Session("LoginIdAccount") & "," & IdAccount & "," & Dtos() & ",'Compila')"
   ConnMsde.execute MyQ
   IdAffidamentoRichiesta = GetTableIdentity("AffidamentoRichiesta")

   MyQ = "" 
   MyQ = MyQ & "insert into AffidamentoRichiestaComp (IdAffidamentoRichiesta,IdAccountCliente,IdCompagnia"
   MyQ = MyQ & ",DataRichiesta,IdStatoAffidamento)"
   MyQ = MyQ & " values (" & IdAffidamentoRichiesta & "," & IdAccount & ",0"
   MyQ = MyQ & "," &  Dtos() & ",'Compila')"
   ConnMsde.execute MyQ
   IdAffidamentoRichiestaComp = GetTableIdentity("AffidamentoRichiestaComp")

   'documenti validi per anno solare 
   DataDocumenti = Right("20" & Year(Date()),4) & "1231"
   excSql = "CheckDocAffidamento " & IdAffidamentoRichiesta & "," & DataDocumenti
   'response.write excSql 
   'response.end 
   ConnMsde.execute excSql
      
   'mi sposto alla gestione 
   xx=RemoveSwap()
      
   Session("swap_IdAffidamentoRichiesta")     = IdAffidamentoRichiesta
   Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
   Session("swap_IdAccountCliente")           = IdAccount
   Session("swap_PaginaReturn")               = PaginaReturn
   response.redirect RitornaA("configurazioni/Clienti/Affidamento/DocumentazioneRichiesta.asp")
   Response.end 
end if 

MySql = "select * from Cliente Where IdCliente = " & IdCliente 
Rs.CursorLocation = 3
Rs.Open MySql, ConnMsde
DescCliente = Rs("Denominazione")
cf          = Rs("CodiceFiscale")
PI          = Rs("PartitaIva")
Rs.close 

%>

<div class="d-flex" id="wrapper">
    <%
      Session("opzioneSidebar")=opzioneSidebar
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
                <div class="col-11"><h3>Nuova Richiesta Di Affidamento</h3>
                </div>
            </div>
            <% if isCliente()=false then %>
            <div class="row">
               <div class="col-3">
                  <div class="form-group ">
                     <%xx=ShowLabel("Cliente")%>
                     <input type="text" readonly class="form-control" value="<%=DescCliente%>" >
                  </div>        
               </div>            
               <div class="col-2">
                  <div class="form-group ">
                     <%xx=ShowLabel("Cod.fiscale")%>
                     <input type="text" readonly class="form-control" value="<%=cf%>" >
                  </div>        
               </div>      
               <div class="col-2">
                  <div class="form-group ">
                     <%xx=ShowLabel("Partita Iva")%>
                     <input type="text" readonly class="form-control" value="<%=pi%>" >
                  </div>        
               </div>               
            </div>
            <% end if %>
            
   <%
   qSel = ""
   qSel = qSel & " Select * from AffidamentoRichiesta A, StatoServizio B"
   qSel = qSel & " Where IdAccountCliente = " & IdAccount
   qSel = qSel & " and A.IdStatoAffidamento = 'COMPILA'"
   esiste = LeggiCampo(qSel,"IdStatoAffidamento")
   if esiste="" then
      qSel = ""
      qSel = qSel & " Select * from AffidamentoRichiesta A, StatoServizio B"
      qSel = qSel & " Where IdAccountCliente = " & IdAccount
      qSel = qSel & " and A.IdStatoAffidamento = B.IdStatoServizio"
      qSel = qSel & " and B.FlagStatofinale = 0 "
   end if          
   esiste = LeggiCampo(qSel,"IdStatoAffidamento")
   FlagRichiedi = true 
   if esiste<>"" then 
       FlagRichiedi=false
       MsgErrore="Siamo spiacenti; non e' possibile chiedere un nuovo affidamento; E' presente gia' una richiesta in corso"
       MsgInfo=MsgErrore
            %>
        <div class="row">
            <div class="col">
                &nbsp;
            </div>
        </div>			
<!--#include virtual="/gscVirtual/include/showInfoDivRow.asp"-->           

        <div class="row"><div class="mx-auto">
        <br>
        <div class="center">

             <a href='#' title="Vai A Gestione " onclick="localNext();">
                <i class="fa fa-2x fa-arrow-right"></i> Vai A Gestione Richieste</a>  
             </div>

        </div></div>
 
            <%
   end if 

   if FlagRichiedi=true then
      MsgErrore="E' possibile richiedere affidamenti per le compagnie disponibili;"
      MsgInfo=MsgErrore
            %>
        <div class="row">
            <div class="col">
                &nbsp;
            </div>
        </div>			
<!--#include virtual="/gscVirtual/include/showInfoDivRow.asp"-->           

        <div class="row"><div class="mx-auto">
        <br>
        <div class="center">

             <a href='#' title="Prosegui " onclick="localIns();">
                <i class="fa fa-2x fa-arrow-right"></i> Procedi con la richiesta di affidamento</a>  
             </div>

        </div></div>
<%end if%>
            <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
            <!--#include virtual="/gscVirtual/include/paginazione.asp"-->

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
