<%
  NomePagina="FornitoreDettaglio.asp"
  titolo="Menu Supervisor - Dashboard"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionAccount.asp"-->

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
<!--#include virtual="/gscVirtual/js/functionTable.js"-->
<script language="JavaScript">

function localFun(Op,Id)
{
	xx=ImpostaValoreDi("DescLoaded","0");
	xx=ImpostaColoreFocus("IdSezioneRui" + Id,"","white");
	xx=ImpostaColoreFocus("NumeroRui" + Id,"","white");
	xx=ImpostaColoreFocus("DataIscrizioneRui" + Id,"","white");

	xx=ElaboraControlli();
	
 	if (xx==false)
	   return false;
	
	var conta = 0;
	var sz = trim(ValoreDi("IdSezioneRui" + Id));
	if (sz!="-1")
	   conta++;
	var nr = trim(ValoreDi("NumeroRui" + Id));
	if (nr!="")
	   conta++;	
	var di = trim(ValoreDi("DataIscrizioneRui" + Id));	
	if (di!="")
	   conta++;		
	
	if (conta>0 && conta<3 ) {
	    xx=ControllaCampo("IdSezioneRui" + Id,"LI");
		xx=ControllaCampo("NumeroRui" + Id,"TE");
		xx=ControllaCampo("DataIscrizioneRui" + Id,"DTO");
		alert('Dati Non Validi');
		return false;
	}
	
	ImpostaValoreDi("Oper","update");
	document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<!--#include file="FornitoreDettaglioLogica.asp"-->
<% 
  'xx=DumpDic(SessionDic,NomePagina)
%>
<div class="d-flex" id="wrapper">

	<%
	  TitoloNavigazione="Configurazioni"
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
			<%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h3>Gestione Profilo fornitore:</b> <%=DescPageOper%> </h3>
				</div>
			</div>

            <!--#include file="FornitoreDettaglioElabora.asp"-->
			<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
			</form>
			<!--#include virtual="/gscVirtual/include/FormSoggetti.asp"-->
		</div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->
  <script>
    $(document).ready(function(){
      $('input[name="DescCognome0"]').focus();
    });
  </script>
  


</body>

</html>
