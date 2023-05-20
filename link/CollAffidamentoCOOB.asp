<%NomePagina="CollAffidamentoCOOB.asp"%>
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
     titolo="Menu Collaboratore - Affidamento"
  %>
  <!--#include virtual="/gscVirtual/include/head.asp"-->

  <!-- Custom styles for this template -->
  <link href="/gscVirtual/css/simple-sidebar.css" rel="stylesheet">

</head>

<body>

  <div class="d-flex" id="wrapper">
   <%
     TitoloNavigazione="Affidamento Coobbligato per cliente"
     Session("opzioneSidebar")="affc"
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
  <script>
	function send01(op,de)
	{
	  xx=ImpostaValoreDi("CurrentFun",de);
	  xx=ImpostaValoreDi("PageToCall","<%=VirtualPath%>" + op);
	  document.FdatiCli.submit();
	}
  </script>
				  
      <div class="container-fluid bg-light">
         <div class="row">
            <!-- Column -->
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
				  <%
				  op="configurazioni/clienti/ClienteCoobbligati.asp"
				  de="Gestione ATI per cliente "
				  
				  %>
                  <a href="#" onclick="send01('<%=op%>','<%=de%>')">
                     <div class="box bg-info text-center">
                        <h1 class="font-light text-white"></h1>
                        <h6 class="text-white">Gestione Coobbligato Cliente</h6>
                     </div>
                  </a>
               </div>
            </div>			
         </div>			
		 
		 <% 
		 flagGesCoob = false 
		 tmp=GetDiz(session("Login_Parametri") ,"VAL_COB")
		 if instr(tmp,"COLL")>0 then 
		    flagGesCoob = true 
		 end if 
		 flagGesAti  = false 
		 tmp=GetDiz(session("Login_Parametri") ,"VAL_ATI")
		 if instr(tmp,"COLL")>0 then 
		    flagGesAti = true 
		 end if 
		 if session("LivelloAccount")=1 and (flagGesCoob=true or flagGesAti=true) then 
		 %>
         <div class="row">
            <!-- Column -->
			<%if flagGesCoob=true then %>
            <div class="col-md-6 col-lg-2 col-xlg-3">
               <div class="card card-hover">
                  <a href="<%=VirtualPath%>configurazioni/clienti/ValidazioneCoobbligatoColl.asp">
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
			<%end if %>
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
