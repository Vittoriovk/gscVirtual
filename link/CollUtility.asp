<%NomePagina="CollUtility.asp"%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->

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
     TitoloNavigazione="Utility"
     Session("opzioneSidebar")="util"
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
                  <a href="<%=VirtualPath%>configurazioni/CambioPassword.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Cambio Password</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>configurazioni/prodotti/ProdottiDisponibiliColl.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Prodotti Disponibili</h6>
                     </div>
                  </a>
               </div>
            </div>						
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>configurazioni/prodotti/ProdottiListinoColl.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Listino Personale</h6>
                     </div>
                  </a>
               </div>
            </div>			
         </div>
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
