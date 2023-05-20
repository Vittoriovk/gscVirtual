<%NomePagina="CollDashboard.asp"%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
  <%
     titolo="Menu Collaboratore - Dashboard"
     livelloPagina="00"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->	 
    <%
    xx = DestroyCurrent()
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
        <h1 class="mt-4">DashBoard Collaboratore</h1>
		
   <%
   xx=RemoveSwap()
   MySql = "" 
   MySql = MySql & " select top 1 A.IdAccountCoobbligato as flag" 
   MySql = MySql & " from AccountCoobbligato A, StatoServizio C "
   MySql = MySql & " Where A.IdStatoValidazione = C.IdStatoServizio "
   MySql = MySql & " and   C.FlagStatoFinale = 0"
   MySql = MySql & " and   A.IdStatoValidazione <>''"
   MySql = MySql & " and   A.IdStatoValidazione not in ('DOCU')"
   MySql = MySql & " and   A.TipoGestore like '%COLL%'"
   MySql = MySql & " and   A.IdAccountBackOffice in (0," &  Session("LoginIdAccount") & ")"
   Trovato=LeggiCampo(MySql,"flag")
   if Trovato<>"" then 
      TitoloR="Ci sono richieste di validazione Coobbligato da gestire"
	  linkRef="/gscVirtual/configurazioni/clienti/ValidazioneCoobbligatoColl.asp"
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
   MySql = MySql & " and   A.TipoGestore like '%COLL%'"
   MySql = MySql & " and   A.IdAccountBackOffice in (0," &  Session("LoginIdAccount") & ")"
   Trovato=LeggiCampo(MySql,"flag")
   if Trovato<>"" then 
      TitoloR="Ci sono richieste di validazione A.T.I. da gestire"
	  linkRef="/gscVirtual/configurazioni/clienti/ValidazioneATIColl.asp"
   %>
   <!--#include virtual="/gscVirtual/include/showInfoVai.asp"-->
   <%
   end if   
   
		%>		
		
        <!--#include virtual="/gscVirtual/utility/RiepilogoEventi.asp"-->
		<%
		 'response.write "attivi:" & Session("Login_servizi_attivi")
		 'For Each obj in session("Login_Parametri").keys
         '    response.write "Key: " & obj & " Value: " & session("Login_Parametri")(obj)
         'Next
		%>
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
