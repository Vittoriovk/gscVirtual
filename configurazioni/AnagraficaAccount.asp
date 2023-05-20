<%
  NomePagina="AnagraficaAccount.asp" 
  titolo="Menu Supervisor - Dashboard"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<%
  livelloPagina="00"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->
<%
  SoloLettura=false
  'xx=DumpDic(Pagedic,NomePagina)
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include virtual="/gscVirtual/include/head.asp"-->
<!-- Custom styles for this template -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">
</head>

<body> 
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,Titolo,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<!--#include file="AnagraficaAccountLogica.asp"-->
 
<div class="d-flex" id="wrapper">

<%
   TitoloNavigazione="profilo"
   Session("opzioneSidebar")="prof"
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
				<div class="col"><h3>Gestione Profilo : <%=Session("LoginNominativo") %></h3>
				</div>
			</div>
			
			<!--#include file="AnagraficaAccountElabora.asp"-->
			<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
			</form>
		</div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->


</html>
