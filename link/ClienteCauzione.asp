<%NomePagina="ClienteCauzione.asp"
  default_check_profile="Clie"
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
	flagCauzProv = false 
	flagCauzDefi = false 
	
	if instr(Session("Login_servizi_attivi"),"CAUZ_PROV")>0 then 
	   flagCauzProv = true  
	end if 
	if instr(Session("Login_servizi_attivi"),"CAUZ_DEFI")>0 then 
	   flagCauzDefi = true  
	end if 
	
	  TitoloNavigazione="Cauzioni cliente"
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
	%>	
		<div class="container-fluid bg-light">
		   <% if flagCauzProv = true then %>
		   <div class="row">
				<!-- Column -->
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class="card card-hover">
						<a href="<%=VirtualPath%>CauzioneProvvisoria/SwapClieCauz.asp">
							<div class="box bg-success text-center">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Gestione Cauzione Provvisoria</h6>
							</div>
						</a>
					</div>
				</div>
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class="card card-hover">
						<a href="<%=VirtualPath%>CauzioneProvvisoria/SwapClieCauzAtti.asp">
							<div class="box bg-success text-center">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Cauzioni Provvisorie Attive</h6>
							</div>
						</a>
					</div>
				</div>
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class="card card-hover">
					    <a href="<%=VirtualPath%>CauzioneProvvisoria/SwapClieCauzStor.asp">
							<div class="box bg-success text-center">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Storico Cauzioni Provvisorie</h6>
							</div>
						</a>
					</div>
				</div>			
				
			</div>
			<%end if %>
			
			<% if flagCauzDefi = true then %>
		<div class="row">
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneDefinitiva/NuovaCauzioneDefinitiva.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Nuova Cauzione Definitiva</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneDefinitiva/CauzioneDefinitivaGestione.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Gestione Cauzioni Definitive</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneDefinitiva/CauzioneDefinitivaStorico.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Storico Cauzioni Definitive</h6>
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
