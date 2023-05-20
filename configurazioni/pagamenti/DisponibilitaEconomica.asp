<%
  NomePagina="DisponibilitaEconmica.asp"
  titolo="Menu - Disponibilita' economica"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAffidamento.asp"-->
<%
  livelloPagina="00"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include virtual="/gscVirtual/include/head.asp"-->
<!-- Custom styles for this template -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">

</head>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%
PaginaReturn = Session("swap_PaginaReturn")
if PaginaReturn="" then 
   PaginaReturn=Session("LoginHomePage")
end if 

IdAccount    = Cdbl("0" & Session("swap_IdAccount"))

if Cdbl(IdAccount)=0 then 
   response.redirect RitornaA(PaginaReturn)
   response.end 
end if 
 
%>

<div class="d-flex" id="wrapper">
	<%
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
			<form name="Fdati" Action="<%=NomePagina%>" method="post">
 
			<div class="row">
				<%
				if PaginaReturn<>"" then 
				   RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"
				%>
				   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<% end if %>
				<div class="col-11"><h3>Disponibilit&agrave; economica utente</h3>
				</div>
			</div>
            <%if cdbl(IdAccount)<>cdbl(session("LoginIdAccount"))  then %>
			<div class="row">
			   <div class="col-3">
			   
                  <div class="form-group ">
				     <%xx=ShowLabel("Utente")
					 qSel = "Select * from Account Where IdAccount=" & IdAccount
					 'response.write qSel
					 DenomUtente=LeggiCampo(qSel,"Nominativo")
					 
					 %>
					 <input type="text" readonly class="form-control" value="<%=DenomUtente%>" >
                  </div>		
			   </div>
			</div>
			<br>
            <%end if
              IdAccountModPag = IdAccount 
			  OpDocAmm = "L"
			  ContaModLMP = 0
			%>
            <!--#include virtual="/gscVirtual/configurazioni/pagamenti/ListaModPag.asp"-->

			</form>
		</div> <!-- container fluid -->
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
  <script>
    $(document).ready(function(){
      $('[data-toggle="tooltip"]').tooltip();   
    });
  </script>
  <script>
$('.btn').onClick(function(e){
  e.preventDefault();
});  
</script>
</body>

</html>
