<%NomePagina="CollCauzioni.asp"%>
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
     titolo="Menu Collaboratore - Cauzioni"
  %>
  <!--#include virtual="/gscVirtual/include/head.asp"-->

  <!-- Custom styles for this template -->
  <link href="/gscVirtual/css/simple-sidebar.css" rel="stylesheet">

</head>
			      <script>
                    function send01(op,de)
                    {
                      xx=ImpostaValoreDi("CurrentFun",de);
					  xx=ImpostaValoreDi("PageToCall","<%=VirtualPath%>" + op);
                      document.FdatiCli.submit();
                    }
				  </script>
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
   %>   
      <div class="container-fluid bg-light">
	  <% if flagCauzProv = true then %>
         <div class="row">
            <!-- Column -->
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">

				  <%
				  op="CauzioneProvvisoria/CheckCauzioneProvvisoria.asp"
				  de="Nuova richiesta di Cauzione Provvisoria"
				  
				  %>
                  <a href="#" onclick="send01('<%=op%>','<%=de%>')">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Provvisoria - Nuova Richiesta</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneProvvisoria/CauzioneGestione.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Provvisoria - richieste da inviare</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneProvvisoria/SwapCauzProvAttEmis.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Provvisoria - richiesta in attesa emissione</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneProvvisoria/SwapCauzProvAttConf.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Provvisoria - emesse in attesa di conferma</h6>
                     </div>
                  </a>
               </div>
            </div>						
         </div>
		<div class="row">			

            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneProvvisoria/CauzioneGestione.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Provvisoria - polizze emesse e confermate</h6>
                     </div>
                  </a>
               </div>
            </div>						
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneProvvisoria/SwapCauzProvAnnulla.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Provvisoria - polizze annullate</h6>
                     </div>
                  </a>
               </div>
            </div>									
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneProvvisoria/CauzioneGestione.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Provvisoria - polizze in attesa di svincolo</h6>
                     </div>
                  </a>
               </div>
            </div>									
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneProvvisoria/CauzioneGestione.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Provvisoria - polizze svincolate</h6>
                     </div>
                  </a>
               </div>
            </div>									
			
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneProvvisoria/CauzioneStorico.asp">
                     <div class="box bg-success text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">&nbsp;Storico Cauzioni Provvisorie - da rimuovere</h6>
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

				  <%
				  op="CauzioneDefinitiva/NuovaCauzioneDefinitiva.asp"
				  de="Nuova richiesta di Cauzione Definitiva"
				  
				  %>
                  <a href="#" onclick="send01('<%=op%>','<%=de%>')">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">&nbsp;&nbsp;Nuova Cauzione &nbsp;&nbsp;Definitiva</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneDefinitiva/CauzioneDefinitivaGestione.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">&nbsp;Gestione Cauzioni Definitive</h6>
                     </div>
                  </a>
               </div>
            </div>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>CauzioneDefinitiva/CauzioneDefinitivaStorico.asp">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">&nbsp;&nbsp;Storico Cauzioni&nbsp;&nbsp;&nbsp; Definitive</h6>
                     </div>
                  </a>
               </div>
            </div>         
         </div>
		 <%end if %>
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
