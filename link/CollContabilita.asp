<%
  NomePagina="CollContabilita.asp"
  default_check_profile="Coll"
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
     titolo="Menu Collaboratore - Contabilit&agrave;"
  %>
  <!--#include virtual="/gscVirtual/include/head.asp"-->

  <!-- Custom styles for this template -->
  <link href="/gscVirtual/css/simple-sidebar.css" rel="stylesheet">

</head>

<body>

  <div class="d-flex" id="wrapper">
	<%
	  TitoloNavigazione="Contabilit&agrave;"
	  Session("opzioneSidebar")="cont"
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
						<a href="<%=VirtualPath%>estrattoConto/CreaEstrattoCliente.asp">
							<div class="box bg-success text-center">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Estratto Conto Clienti</h6>
							</div>
						</a>
					</div>
				</div>				
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class="card card-hover">
						<a href="<%=VirtualPath%>estrattoConto/creaEstrattoCollaboratore.asp">
							<div class="box bg-success text-center">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Estratto Conto Collaboratori</h6>
							</div>
						</a>
					</div>
				</div>				
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class="card card-hover">
						<a href="<%=VirtualPath%>EstrattoConto/GestioneEstrattoCollaboratore.asp">
							<div class="box bg-success text-center">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Gestione Estratto Conto</h6>
							</div>
						</a>
					</div>
				</div>
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class="card card-hover">
					    <a href="<%=VirtualPath%>formazione/GestioneEstrattoStorico.asp">
							<div class="box bg-success text-center">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Storico Estratto Conto</h6>
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
