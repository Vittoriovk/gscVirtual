<%
  NomePagina="ListaClientiBack.asp"
  titolo="Menu - Clienti"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
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
<script>

function setClie(id)
{
	xx=ImpostaValoreDi("ItemToModify",id);
}
function callFun(action)
{
	var t=ValoreDi("ItemToModify");
	xx=ImpostaValoreDi("Oper",action);
	document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%

if FirstLoad then 
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   OperAmmesse   = getValueOfDic(Pagedic,"OperAmmesse")
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
   if OperAmmesse="" then 
      OperAmmesse = Session("swap_OperAmmesse")
   end if    
else
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   OperAmmesse   = getValueOfDic(Pagedic,"OperAmmesse")
end if 
OperAmmesse="CRUD"

'registro i dati della pagina 
xx=setValueOfDic(Pagedic,"PaginaReturn"  ,PaginaReturn)
xx=setValueOfDic(Pagedic,"OperAmmesse"   ,OperAmmesse)
xx=setCurrent(NomePagina,livelloPagina) 

if Oper="CALL_AFFI" then
   xx=RemoveSwap()
   itemId = Cdbl("0" & Request("ItemToModify"))
   Session("swap_IdCliente") = itemId
   Session("swap_PaginaReturn")    = "configurazioni/Clienti/" & NomePagina
   Session("swap_OperAmmesse")     = OperAmmesse
   response.redirect RitornaA("configurazioni/Clienti/ClienteAffidamentoCompagnia.asp")
   response.end 
   
end if  

ItemToModify = "0"

%>

<div class="d-flex" id="wrapper">
	<%
	  Session("opzioneSidebar")="clie"
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
				<div class="col-11"><h3>Lista Clienti</h3>
				</div>
			</div>
			<%
			AddRow=true
			dim CampoDb(10)
			ElencoOption = ";0;Nominativo;1;Codice Fiscale;2;Partita Iva;2;"
            CampoDB(1)   = "a.Denominazione"
			CampoDb(2)   = "a.CodiceFiscale"
			CampoDb(3)   = "a.PartitaIva"
			
			%>
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
		
<%
			'caricamento tabella 
			if Condizione<>"" then 
				Condizione = " And " & Condizione
			end if 

			Set Rs = Server.CreateObject("ADODB.Recordset")
			MySql = "" 
			MySql = MySql & " Select A.*,c.FlagAttivo,c.DescBlocco From Cliente A "
			MySql = MySql & " left join Account C on a.idAccount = C.IdAccount"
			MySql = MySql & " Where a.IdAzienda = " & Session("IdAziendaWork") 
			'MySql = MySql & " and   a.IdAccountLivello1 = " & Session("LoginRefAccountLev1")
			'if session("LivelloAccount")>=2 then 
			'   MySql = MySql & " and   a.IdAccountLivello2 = " & Session("LoginRefAccountLev2")
			'end if 
			'if session("LivelloAccount")>=3 then 
			'   MySql = MySql & " and   a.IdAccountLivello3 = " & Session("LoginRefAccountLev3")
			'end if 			
            MySql = MySql & " and c.FlagAttivo<>'N' "
			MySql = MySql & Condizione & " order By Denominazione"

			Rs.CursorLocation = 3 
			Rs.Open MySql, ConnMsde

			DescLoaded=""
			NumCols = numC + 1
			NumRec  = 0
			ShowNew    = true
			ShowUpdate = false
			MsgNoData  = ""
			'elenco azioni 
			if rs.EOF=false then 
			%>
			<div class="table-responsive">
			<table class="table"><tbody>
			<thead>
				<tr>
					<th scope="col">azioni</th>
				</tr>
			</thead>
			<tr><td>
			 <div class="row">
				<!-- Column -->
				<div class="col-md-6 col-lg-2 col-xlg-3">
				   <div class="card card-hover">
					  <a href="#" onclick="callFun('CALL_AFFI')">
						 <div class="box bg-success text-center">
							<h1 class="font-light text-white"></h1>
							<h6 class="text-white">Affidamenti</h6>
						 </div>
					  </a>
				   </div>
				</div>		
			</div>
			</td></tr>
			</tbody></table></div>
			<%
			end if 
%>

<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

			<div class="table-responsive"><table class="table"><tbody>
			<thead>
				<tr>
				<th scope="col">Sel</th>
				<th scope="col">Cliente</th>
		        <th scope="col">Codice fiscale</th>
		        <th scope="col">Partita Iva</th>
		        <th scope="col">Riferimento</th>
				</tr>
			</thead>
            
<%
			if MsgNoData="" then 
				if PageSize>0 then 
					Rs.PageSize = PageSize
					pageTotali = rs.PageCount
					NumRec=0
					if Cpag<=0 then 
						Cpag =1
					end if 
					if Cpag>PageTotali then 
						CPag=PageTotali
					end if  
					Rs.absolutepage=CPag
				end if
				NumRec=0
				Primo=0
				Do While Not rs.EOF and (NumRec<PageSize or Pagesize<=0)
					Primo=Primo+1
					NumRec=NumRec+1
					Id=Rs("IdCliente")
					selected=""
					if Cdbl(ItemToModify)=0 then 
					   ItemToModify=Id
					   Selected=" checked "
					end if 
					DescLoaded=DescLoaded & Id & ";"
		%>
			<tr scope="col">
			
			<td><input  type="radio" name="SelClie" id="SelClie" <%=Selected%> onclick="setClie(<%=Id%>)">
			</td>			
				<td>
					<input class="form-control" type="text" readonly value="<%=Rs("Denominazione")%>">
				</td>
				<td>
					<input class="form-control" type="text" readonly value="<%=Rs("Codicefiscale")%>">
				</td>
				<td>
					<input class="form-control" type="text" readonly value="<%=Rs("PartitaIva")%>">
				</td>

				<td>
				    <%
					DescRif=""
					IdRif=0 
				   IdRif=Rs("IdAccountLivello1")

					if session("LivelloAccount")=2 and Rs("IdAccountLivello3")>0 then 
					   IdRif=Rs("IdAccountLivello3")
					end if 
					if IdRif>0 then 
					   DescRif = LeggiCampo("Select * from Account Where IdAccount=" & IdRif,"Nominativo")
					end if 
					
					
					%>
					<input class="form-control" type="text" readonly value="<%=DescRif%>">
				</td>				
				

			</tr>
		<%	
		rs.MoveNext
	Loop
end if 
rs.close

%>

</tbody></table></div> <!-- table responsive fluid -->

			<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
			<!--#include virtual="/gscVirtual/include/paginazione.asp"-->

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
