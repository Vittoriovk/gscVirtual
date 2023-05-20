<%NomePagina="SupervisorConfigurazioni.asp"%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->

<%
  Set Session("PERCORSO") = Server.CreateObject("Scripting.Dictionary")
  livelloPagina="01"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->
<% 
  'xx=DumpDic(SessionDic,NomePagina)
%>
<!---->
<!DOCTYPE html>
<html lang="en">
<head>
  <%
     titolo="Menu Supervisor - Dashboard"
  %>
  <!-- Custom styles for this  -->
<!-- plugins:css -->
<link rel="stylesheet" href=" vendors/feather/feather.css">
<link rel="stylesheet" href=" vendors/ti-icons/css/themify-icons.css">
<link rel="stylesheet" href=" vendors/css/vendor.bundle.base.css">
<!-- endinject -->
<!-- Plugin css for this page -->
<link rel="stylesheet" href=" vendors/select2/select2.min.css">
<link rel="stylesheet" href=" vendors/select2-bootstrap-theme/select2-bootstrap.min.css">
<link rel="stylesheet" href=" vendors/mdi/css/materialdesignicons.min.css">
<!-- End plugin css for this page -->
<!-- inject:css -->
<link rel="stylesheet" href=" css/vertical-layout-light/style.css">
  <link href="css/simple-sidebar.css" rel="stylesheet">

</head>

<body>

  <div class="d-flex" id="wrapper">

	<div>
		<%
		callP=VirtualPath & "bar/" & Session("TopBar_" & Session("LoginIdAccount")) 
		Server.Execute(callP) 
		%>
	</div>
	<div class="container-fluid page-body-wrapper">
		<div>
			<%
			TitoloNavigazione="Configurazioni"
			Session("opzioneSidebar")="conf"
			callP=VirtualPath & "bar/" & Session("sideBar_" & Session("LoginIdAccount")) 
			Server.Execute(callP) 
			%>
		</div>

    <!-- Page Content -->
       <div class="container-fluid bg-light">
		   <div class="row">
				<!-- Column -->
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class="h-100 card ">
						<a href="configurazioni/tabelle/ramo.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							<div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Ramo</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Gestione dei rami per i servizi previsti" ></i>
								<%end if %>
							</div>
						</a>
					</div>
				</div>
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class="h-100 card">
						<a href="configurazioni/tabelle/caratteristica.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							<div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Template Rischi</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Gestione dei template dei servizi, ad esempio le suddivisioni delle cauzioni definite" ></i>								
								<%end if %>   
							</div>
						</a>
					</div>
				</div>				
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class="h-100 card ">
						<a href="configurazioni/tabelle/rischio.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							<div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Rischio</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Gestione dei rischi per i rami previsti" ></i>
								<%end if %>
							</div>
						</a>
					</div>
				</div>

				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-90 card">
						<a href="configurazioni/tabelle/elenco.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Elenchi</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Gestione elenchi di valori predefiniti" ></i>								
								<%end if %>   								
							</div>
						</a>
					</div>
				</div>		
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-90 card card-hover">
						<a href="configurazioni/tabelle/datoTecnico.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Dati Tecnici</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Gestione Dati Aggiuntivi " ></i>								
								<%end if %> 								
							</div>
						</a>
					</div>
				</div>
			</div>
			
			<div class="row">
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/documento.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Documenti</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Gestione Dei documenti Da utilizzare" ></i>								
								<%end if %> 									
							</div>
						</a>
					</div>
				</div>
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/listaDocumento.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Liste Documenti</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Creazione di liste di documenti " ></i>								
								<%end if %> 									
							</div>
						</a>
					</div>
				</div>
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/prodotti/ProdottoTemplate.asp">
						    <%if default_show_info=false then
							     base_color = "bg-warning"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Template Prodotto</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Definizione dei prodotti e configurazione " ></i>								
								<%end if %> 									
							</div>
						</a>
					</div>
				</div>				
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/compagnia.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Compagnie</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Gestione Compagnie Assicurative e/o produttori di servizi " ></i>								
								<%end if %> 								
							</div>
						</a>
					</div>
				</div>


				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/TratFisc.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Tratt.Fiscali</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Definizione aliquote iva e imposte " ></i>								
								<%end if %> 									
							</div>
						</a>
					</div>
				</div>
				<%if false then %>
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/ServizioFasciaRibasso.asp">
						    <%if default_show_info=false then
							     base_color = "bg-warning"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Fasce Ribasso Cauzione Definitiva</h6>
								<%if default_show_info=true then %>
								<i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								title="Gestione delle Fascie di ribasso per le cauzioni definite" ></i>								</h6>
								<%end if %>
	
							</div>
						</a>
					</div>
				</div>
				<%end if %>
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/Certificazione.asp">
						    <%if default_show_info=false then
							     base_color = "bg-warning"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Certificazione per Cauzioni</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Definizione delle certificazioni disponibili per le cauzioni " ></i>								
								<%end if %> 								
							</div>
						</a>
					</div>
				</div>				

				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/ServizioDocumentoCoobbligato.asp">
						    <%if default_show_info=false then
							     base_color = "bg-warning"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Documentazione per Coobbligati</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Definizione dei documenti necessari per i coobbligati " ></i>								
								<%end if %> 								
							</div>
						</a>
					</div>
				</div>				
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/ServizioDocumentoATI.asp">
						    <%if default_show_info=false then
							     base_color = "bg-warning"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Documentazione per ATI</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Definizione dei documenti necessari per i le ATI (Associazione temporanea di impresa) " ></i>								
								<%end if %> 								
							</div>
						</a>
					</div>
				</div>				
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/Parametri.asp">
						    <%if default_show_info=false then
							     base_color = "bg-warning"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Parametri Generali</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Definizione di parametri validi per tutti i prodotti " ></i>								
								<%end if %> 									
							</div>
						</a>
					</div>
				</div>
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/prodotti/GruppoProdotti.asp">
						    <%if default_show_info=false then
							     base_color = "bg-warning"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Raggruppamento Prodotti</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Definizione di gruppi di prodotto " ></i>								
								<%end if %> 									
							</div>
						</a>
					</div>
				</div>				
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/fornitore.asp">
						    <%if default_show_info=false then
							     base_color = "bg-info"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Fornitore</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Definizione dei fornitori dei prodotti e configurazioni " ></i>								
								<%end if %> 									
							</div>
						</a>
					</div>
				</div>			
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/tabelle/DirittiEmissione.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Diritti di Emissione</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Definizione dei diritti di emissione " ></i>								
								<%end if %> 									
							</div>
						</a>
					</div>
				</div>							
				<div class="col-md-6 col-lg-2 col-xlg-3">
					<div class=" h-100 card card-hover">
						<a href="configurazioni/prodotti/ProdottiAttiva.asp">
						    <%if default_show_info=false then
							     base_color = "bg-success"
							  else
							     base_color = "bg-info"
							  end if%>
							  <div class="box <%=base_color%> text-center ">
								<h1 class="font-light text-white"></h1>
								<h6 class="text-white">Attiva Prodotti</h6>
								<%if default_show_info=true then %>
								   <i class="fa fa-info-circle" data-toggle="tooltip" data-placement="top" 
								   title="Attivazione, Disattivazione dei prodotti" ></i>								
								<%end if %> 									
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
<script src=" vendors/js/vendor.bundle.base.js"></script>
<!-- endinject -->
<!-- Plugin js for this page -->
<script src=" vendors/typeahead.js/typeahead.bundle.min.js"></script>
<script src=" vendors/select2/select2.min.js"></script>
<!-- End plugin js for this page -->
<!-- inject:js -->
<script src=" js/off-canvas.js"></script>
<script src=" js/hoverable-collapse.js"></script>
<script src=" js/template.js"></script>
<script src=" js/settings.js"></script>
<script src=" js/todolist.js"></script>
<!-- endinject -->
<!-- Custom js for this page-->
<script src=" js/file-upload.js"></script>
<script src=" js/typeahead.js"></script>
<script src=" js/select2.js"></script>
<!-- End custom js for this page-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

<script TYPE="text/javascript">
	function mouseover(elem) {
		elem.style.color = '#FF0000';
	}
	function mouseout(elem) {
		elem.style.color = '#4B49AC';
	}
</script>
<script>
	(function() {    
            var dialog = document.getElementById('myFirstDialog');    
            document.getElementById('show').onclick = function() {    
                dialog.show();    
            };    
            document.getElementById('hide').onclick = function() {    
                dialog.close();    
            };    
    })();  
</script>
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
