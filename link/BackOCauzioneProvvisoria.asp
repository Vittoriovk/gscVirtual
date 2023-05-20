<%NomePagina="BackOCauzioneProvvisoria.asp"
  default_check_profile="BackO"
%>
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
     TitoloNavigazione="Cauzioni per cliente"
     Session("opzioneSidebar")="cauz"
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

	showProv = false
	showDefi = false 

	if isServizioAttivo("CAUZ_PROV") then 
	   showProv = true 
	end if 
	if isServizioAttivo("CAUZ_DEFI") then 
	   showDefi = true 
	end if 	
   %>   
      <div class="container-fluid bg-light">
	     <%if showProv = true then %>
         <div class="row">
            <!-- Column -->
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>cauzioneProvvisoria/CauzioneClienteGestioneBackO.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Gestione Richieste Cauzioni Provvisorie</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>cauzioneProvvisoria/swapBackCauzAtti.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Cauzioni Provvisorie Attive in conferma</h6>
                     </div>
                  </a>
               </div>
            </div>         
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>cauzioneProvvisoria/CauzioneClienteGestioneBackOStorico.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Storico Richieste Cauzioni Provvisorie</h6>
                     </div>
                  </a>
               </div>
            </div>  			
         </div>
		 <%end if %>
		  <%if showDefi = true then %>
         <div class="row">
            <!-- Column -->
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>cauzioneDefinitiva/CauzioneDefinitivaGestione.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Gestione Richieste Cauzioni Definite</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>cauzioneDefinitiva/CauzioneDefinitivaStorico.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Storico Richieste Cauzioni Definitive</h6>
                     </div>
                  </a>
               </div>
            </div>         
         </div>
		 <%end if %>		 
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
