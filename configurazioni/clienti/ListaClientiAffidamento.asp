<%
  NomePagina="ListaClientiAffidamento.asp"
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
function localUpd(id)
{
	xx=ImpostaValoreDi("ItemToRemove",id);
	xx=ImpostaValoreDi("Oper","CALL_UPD");
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

'registro i dati della pagina 
if PaginaReturn="" then 
   PaginaReturn=Session("LoginHomePage")
end if 
xx=setValueOfDic(Pagedic,"PaginaReturn"  ,PaginaReturn)
xx=setValueOfDic(Pagedic,"OperAmmesse"   ,OperAmmesse)
xx=setCurrent(NomePagina,livelloPagina) 

if Oper="CALL_INS" or Oper="CALL_UPD" then
   xx=RemoveSwap()
   Session("swap_OperTabella")     = Oper
   Session("swap_IdCliente") = itemId
   Session("swap_PaginaReturn")    = "configurazioni/Clienti/" & NomePagina
   Session("swap_OperAmmesse")     = OperAmmesse
   response.redirect RitornaA("configurazioni/Clienti/ClienteModifica.asp")
   response.end 
   
end if  

%>

<div class="d-flex" id="wrapper">
	<%
	  Session("opzioneSidebar")="affi"
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

				RiferimentoA="col-1  text-center;" & VirtualPath & "link/AdminIntermediari.asp" & ";;2;prev;Indietro;;"
				%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				
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
	IdCollaboratore=0
	IdCollMaster   =0
	IdCollMasterLev=0
	if Session("LoginTipoUtente")=ucase("Coll") then 
	   if ucase(Session("LoginTipoCollaboratore"))<>"SEGN" then 
	      IdCollMaster   =Session("LoginIdAccount")
		  IdCollMasterLev=session("LivelloAccount")
       end if 
	end if 
	
	if IdCollMaster>0 and IdCollMasterLev>0 then %>
	<div class="row no-row-margin " style="margin-top: 10px;margin-bottom: 10px;" >

      <div class="col-1 s1 no-margin font-weight-bold">
	     Collaboratore
	  </div>
	  
      <div class="col-9 no-margin">
	  <%
	    stdClass="class='form-control form-control-sm'"
	    IdCollaboratore=Cdbl("0" & Request("IdCollaboratore0"))
        inC = "Select IdCollaboratore From Collaboratore Where IdAccountLivello" & IdCollMasterLev & "=" & IdCollMaster
	    q="Select * from Collaboratore Where IdCollaboratore in (" & inC & ") "
		q=Q & " order by Denominazione "
		'Where 
	    response.write ListaDbChangeCompleta(q,"IdCollaboratore0",IdCollaboratore ,"IdCollaboratore","Denominazione" ,1,"","","","","",stdClass)
	  
	  %>
	  </div>	
      <div class="col-2 no-margin">
	  </div>	

	</div>
	<%end if %>
			
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
			MySql = MySql & " and   a.IdAccountLivello1 = " & Session("LoginRefAccountLev1")
			if session("LivelloAccount")>=2 then 
			   MySql = MySql & " and   a.IdAccountLivello2 = " & Session("LoginRefAccountLev2")
			end if 
			if session("LivelloAccount")>=3 then 
			   MySql = MySql & " and   a.IdAccountLivello3 = " & Session("LoginRefAccountLev3")
			end if 			
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
%>

<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

			<div class="table-responsive"><table class="table"><tbody>
			<thead>
				<tr>
		<th scope="col">Cliente</th>
		<th scope="col">Codice fiscale</th>
		<th scope="col">Partita Iva</th>
		<th scope="col" width="20">Attivo</th>
		<th scope="col">Riferimento</th>
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
					Id=Rs("IdCliente")
					DescLoaded=DescLoaded & Id & ";"
		%>
			<tr scope="col">
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
					  'response.write Rs("FlagAttivo")
					  if Rs("FlagAttivo")="S" then 
					     FlagAttivo="SI" 
					  else
					     FlagAttivo="NO:" & Rs("DescBlocco")
					end if %>
					
					<input class="form-control" type="text" readonly value="<%=FlagAttivo%>">
				</td>
				<td>
				    <%
					DescRif=""
					IdRif=0 
					if session("LivelloAccount")=1 and Rs("IdAccountLivello2")>0 then 
					   IdRif=Rs("IdAccountLivello2")
					end if 
					if session("LivelloAccount")=2 and Rs("IdAccountLivello3")>0 then 
					   IdRif=Rs("IdAccountLivello3")
					end if 
					if IdRif>0 then 
					   DescRif = LeggiCampo("Select * from Account Where IdAccount=" & IdRif,"Nominativo")
					end if 
					
					
					%>
					<input class="form-control" type="text" readonly value="<%=DescRif%>">
				</td>				
				
			<td>
			    <% if FlagAttivo="SI" then %>
					<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;localUpd('" & id & "');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->		
					<%RiferimentoA="col-2;#;;2;tecn;Configurazioni;;localConf('" & id & "');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->	
					<%
					if instr(OperAmmesse,"D")>0 then 
						RiferimentoA="col-2;#;;2;dele;Cancella;;SalvaSingoloEdAttiva('DEL'," & Id & ",true,'','','');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					<%end if %>
				<%else 
						RiferimentoA="col-2;#;;2;upda;Riattiva;;SalvaSingoloEdAttiva('ATT'," & Id & ",true,'','','');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				
				<%end if %>
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
