<%
  NomePagina="ListaProvvigioniForn.asp"
  titolo="Menu - Provvigioni Fornitore"
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

function localIns()
{
	xx=ImpostaValoreDi("Oper","CALL_INS");
	document.Fdati.submit();
}
function localUpd(id)
{
	xx=ImpostaValoreDi("ItemToRemove",id);
	xx=ImpostaValoreDi("Oper","CALL_UPD");
	document.Fdati.submit();
}
function localConf(id)
{
	xx=ImpostaValoreDi("ItemToRemove",id);
	xx=ImpostaValoreDi("Oper","CALL_CFG");
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
   IdFornitore   = getValueOfDic(Pagedic,"IdFornitore")
    
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
   if OperAmmesse="" then 
      OperAmmesse = Session("swap_OperAmmesse")
   end if    
   if Cdbl("0" & IdFornitore)=0 then 
      IdFornitore = Session("swap_IdFornitore")
   end if    
else
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   OperAmmesse   = getValueOfDic(Pagedic,"OperAmmesse")
   IdFornitore   = getValueOfDic(Pagedic,"IdFornitore")
end if 

OperAmmesse="CRUD"
'registro i dati della pagina 
response.write 
if PaginaReturn="" then 
   PaginaReturn="tabelle/Fornitore.asp"
end if 

IdFornitore = cdbl("0" & IdFornitore)
if IdFornitore=0 then 
   response.redirect VirtualPath & PaginaReturn 
   response.end 
end if 

xx=setValueOfDic(Pagedic,"PaginaReturn"  ,PaginaReturn)
xx=setValueOfDic(Pagedic,"OperAmmesse"   ,OperAmmesse)
xx=setValueOfDic(Pagedic,"IdFornitore"   ,IdFornitore)
xx=setCurrent(NomePagina,livelloPagina) 

if Oper="CALL_INS" or Oper="CALL_UPD" then
   xx=RemoveSwap()
   itemId = Cdbl("0" & Request("ItemToRemove"))
   Session("swap_OperTabella")     = Oper
   Session("swap_IdRegolaProvvigione") = itemId
   Session("swap_IdFornitore")     = IdFornitore
   Session("swap_PaginaReturn")    = "configurazioni/Provvigioni/" & NomePagina
   Session("swap_OperAmmesse")     = OperAmmesse
   Session("swap_TipoProvvigione") = "FORN"
   response.redirect RitornaA("configurazioni/Provvigioni/ModificaProvvigioneForn.asp")
   response.end 
   
end if  

if Oper="DEL" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	if Cdbl(KK)>0 then 
		MyQ = "" 
		MyQ = MyQ & " delete from RegolaProvvigione "
		MyQ = MyQ & " where IdRegolaProvvigione = " & KK
		MsgErrore=VerificaDel("RegolaProvvigione",KK)
		if MsgErrore="" then 	
			ConnMsde.execute MyQ 
			If Err.Number <> 0 Then 
				MsgErrore = ErroreDb(Err.description)
			End If
		End if 
	End if 
	DescIn=""
End if
DescForn=LeggiCampo("Select * from Fornitore Where IdFornitore=" & IdFornitore,"DescFornitore")
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

				RiferimentoA="col-1 text-center;" & VirtualPath & paginaReturn & ";;2;prev;Indietro;;"
				%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				
				<div class="col-11"><h3>Regole Provvigioni Da Fornitore : <%=DescForn%></h3>
				</div>
			</div>
			<%
			AddRow=true
			dim CampoDb(10)
			ElencoOption = ";0;Descrizione;"
            CampoDB(1)   = "a.DescRegolaProvvigione"
			
			%>
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->


<%
			'caricamento tabella 
			if Condizione<>"" then 
				Condizione = " And " & Condizione
			end if 

			Set Rs = Server.CreateObject("ADODB.Recordset")
			MySql = "" 
			MySql = MySql & " Select * From RegolaProvvigione  "
			MySql = MySql & " Where TipoRegola = 'FORN' "
			MySql = MySql & " And IdFornitore=" & IdFornitore
			MySql = MySql & Condizione & " order By DescRegolaProvvigione"

			Rs.CursorLocation = 3 
			Rs.Open MySql, ConnMsde

			DescLoaded=""
			NumCols = numC + 1
			NumRec  = 0
			ShowNew    = true
			ShowUpdate = false
			MsgNoData  = ""
%>

<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

			<div class="table-responsive"><table class="table"><tbody>
			<thead>
				<tr>
					<th scope="col"> Regola Provvigione
						<%
						  if instr(OperAmmesse,"C")>0 then 
						  RiferimentoA="col-2;#;;2;inse;Inserisci;;localIns();N"
						  %>
						<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                          <%end if %>						
					</th>
		<th scope="col">Azioni</th>
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
					Id=Rs("IdRegolaProvvigione")
					DescLoaded=DescLoaded & Id & ";"
		%>
			<tr scope="col">
				<td>
					<input class="form-control" type="text" readonly value="<%=Rs("DescRegolaProvvigione")%>">
				</td>
	
			<td>
					<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;localUpd('" & id & "');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->		

					<%RiferimentoA="col-2;#;;2;dele;Cancella;;SalvaSingoloEdAttiva('DEL'," & Id & ",true,'','','');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->

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
