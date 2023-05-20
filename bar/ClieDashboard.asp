<%NomePagina="ClieDashboard.asp"%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
  <%
     titolo="Menu Cliente - Dashboard"
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
        <h1 class="mt-4">DashBoard Cliente</h1>
		<%
		
		'response.write "attivi:"  & Session("Login_servizi_attivi") & "<br>"
		'response.write "us.prod:" & Session("LoginIdAccountProdotti") & "<br>"
		
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
