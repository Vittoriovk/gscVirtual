<%NomePagina="BackODashboard.asp"%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
  <%
     titolo="Menu Utente - Dashboard"
  %>
  <!--#include virtual="/gscVirtual/include/head.asp"-->

  <!-- Custom styles for this template -->
  <link href="<%=VirtualPath%>css/simple-sidebar.css" rel="stylesheet">

</head>


<body>

<div class="d-flex" id="wrapper">
	<%
	  Session("opzioneSidebar")="dash"
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
        <h2 class="mt-4">Attivit&agrave; in corso</h2>
        <p></p>
		
		<%
   MySql = "" 
   MySql = MySql & " select top 1 a.IdStatoAffidamento as flag" 
   MySql = MySql & " from AffidamentoRichiesta A, StatoServizio B "
   MySql = MySql & " Where A.IdStatoAffidamento = B.IdStatoServizio "
   MySql = MySql & " and   B.FlagStatoFinale = 0"
   MySql = MySql & " and   A.IdStatoAffidamento not in ('COMPILA')"
   MySql = MySql & " and   A.IdAccountBackOffice in (0," &  Session("LoginIdAccount") & ")"
   
   Trovato=LeggiCampo(MySql,"flag")
   if Trovato<>"" then 
      TitoloR="Ci sono richieste di affidamento da gestire"
	  linkRef="/gscVirtual/configurazioni/clienti/Affidamento/ListaRichiesta.asp"
   %>
   <!--#include virtual="/gscVirtual/include/showInfoVai.asp"-->
   <%
   end if 
  
   MySql = "" 
   MySql = MySql & " select top 1 A.IdAccountCoobbligato as flag" 
   MySql = MySql & " from AccountCoobbligato A, StatoServizio C "
   MySql = MySql & " Where A.IdStatoValidazione = C.IdStatoServizioo "
   MySql = MySql & " and   C.FlagStatoFinale = 0"
   MySql = MySql & " and   A.IdStatoValidazione <>''"
   MySql = MySql & " and   A.IdStatoValidazione not in ('DOCU')"
   MySql = MySql & " and   A.TipoGestore like '%BACKO%'"
   MySql = MySql & " and   A.IdAccountBackOffice in (0," &  Session("LoginIdAccount") & ")"
   Trovato=LeggiCampo(MySql,"flag")
   if Trovato<>"" then 
      TitoloR="Ci sono richieste di validazione Coobbligato da gestire"
	  linkRef="/gscVirtual/configurazioni/clienti/ValidazioneCoobbligatoBackO.asp"
   %>
   <!--#include virtual="/gscVirtual/include/showInfoVai.asp"-->
   <%
   end if   

   MySql = "" 
   MySql = MySql & " select top 1 A.IdAccountATI as flag" 
   MySql = MySql & " from AccountCoobbligato A, StatoServizio C "
   MySql = MySql & " Where A.IdStatoValidazione = C.IdStatoServizio "
   MySql = MySql & " and   C.FlagStatoFinale = 0"
   MySql = MySql & " and   A.IdStatoValidazione <>''"
   MySql = MySql & " and   A.IdStatoValidazione not in ('DOCU')"
   MySql = MySql & " and   A.TipoGestore like '%BACKO%'"
   MySql = MySql & " and   A.IdAccountBackOffice in (0," &  Session("LoginIdAccount") & ")"
   Trovato=LeggiCampo(MySql,"flag")
   if Trovato<>"" then 
      TitoloR="Ci sono richieste di validazione A.T.I. da gestire"
	  linkRef="/gscVirtual/configurazioni/clienti/ValidazioneATIBackO.asp"
   %>
   <!--#include virtual="/gscVirtual/include/showInfoVai.asp"-->
   <%
   end if   
   
   'recupero Evento cauzione 

   MySql = "" 
   MySql = MySql & " select top 1 A.IdNumAttivita as flag" 
   MySql = MySql & " from ServizioRichiesto A, StatoServizio C "
   MySql = MySql & " Where A.IdStatoServizio = C.IdStatoServizio "
   MySql = MySql & " and   A.IdAttivita = 'CAUZ_DEFI'"
   MySql = MySql & " and   C.FlagStatoFinale = 0"
   MySql = MySql & " and   A.IdStatoServizio <>''"
   MySql = MySql & " and   A.IdStatoServizio not in ('DOCU')"
   MySql = MySql & " and   A.IdAccountGestore in (0," &  Session("LoginIdAccount") & ")"

   Trovato=LeggiCampo(MySql,"flag")
   if Trovato<>"" then 
      TitoloR="Ci sono Cauzioni Definitive da gestire"
	  linkRef="/gscVirtual/CauzioneDefinitiva/CauzioneDefinitivaGestione.asp"
   %>
   <!--#include virtual="/gscVirtual/include/showInfoVai.asp"-->
   <%
   end if      
		%>
		
		
		
		
        <!--#include virtual="/gscVirtual/utility/RiepilogoEventi.asp"-->        
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
