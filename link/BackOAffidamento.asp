<%NomePagina="BackOAffidamento.asp"%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->

<%
  Set Session("PERCORSO") = Server.CreateObject("Scripting.Dictionary")
  livelloPagina="01"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->
<% 
  
  'xx=DumpDic(SessionDic,NomePagina)
%>

<!DOCTYPE html>
<html lang="en">
<head>
  <%
     titolo="Menu Supervisor - Dashboard"
  %>
  <!--#include virtual="/gscVirtual/include/head.asp"-->

  <!-- Custom styles for this template -->
  <link href="/gscVirtual/css/simple-sidebar.css" rel="stylesheet">

</head>

<body>

  <div class="d-flex" id="wrapper">
   <%
     TitoloNavigazione="Affidamento per cliente"
     Session("opzioneSidebar")="affi"
     callP=VirtualPath & "bar/" & Session("sideBar_" & Session("LoginIdAccount")) 
     Server.Execute(callP) 
   %>   

    <!-- Page Content -->
    <div id="page-content-wrapper">
   <%
    callP=VirtualPath & "bar/" & Session("TopBar_" & Session("LoginIdAccount")) 
    Server.Execute(callP)
    xx=RemoveSwap()
    Session("swap_PaginaReturn") = "link/" & NomePagina   
   %>   
      <div class="container-fluid bg-light">
         <div class="row">
            <!-- Column -->
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>configurazioni/clienti/Affidamento/ListaRichiesta.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Gestione Richieste</h6>
                     </div>
                  </a>
               </div>
            </div>
         </div>
		 
		 
         <div class="row">
            <!-- Column -->
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>configurazioni/clienti/ValidazioneCoobbligatoBackO.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Gestione Validazione Coobbligati</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>configurazioni/clienti/ValidazioneCoobbligatoBackOStorico.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Storico Validazione Coobbligati</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>configurazioni/clienti/ValidazioneATIBackO.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Gestione Validazione <br>A.T.I.</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>configurazioni/clienti/ValidazioneATIBackOStorico.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Storico Validazione <br>A.T.I</h6>
                     </div>
                  </a>
               </div>
            </div>
			
         </div>		 
         <div class="row">
            <!-- Column -->
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>configurazioni/clienti/AffidamentoClienteGestioneBackO.asp">
                     <div class="box bg-warning text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Gestione Importi Affidati per compagnia</h6>
                     </div>
                  </a>
               </div>
            </div>
         </div>
		 <form name="FdatiCli" Action="<%=VirtualPath%>utility/SwapCercaCliente.asp" method="post">
		 <input type="hidden" name="CurrentFun"     id="CurrentFun"     value="">
		 <input type="hidden" name="PageToCall"     id="PageToCall"     value="">
		 <input type="hidden" name="OperAmmesse"    id="OperAmmesse"    value="">
		 <input type="hidden" name="PaginaReturn"   id="PaginaReturn"   value="<%=Session("swap_PaginaReturn")%>">
		 <input type="hidden" name="opzioneSidebar" id="opzioneSidebar" value="<%=Session("opzioneSidebar")%>">

		 </form>
      </div>
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

</body>

</html>
