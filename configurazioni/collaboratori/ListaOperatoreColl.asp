<%
  NomePagina="ListaOperatoreColl.asp"
  titolo="Back Office Collaboratore"
  default_check_profile="COLL"
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
</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%

if FirstLoad then 
   v_attivi="S"
   v_Cessati=""
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   OperAmmesse   = getValueOfDic(Pagedic,"OperAmmesse")
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
   if OperAmmesse="" then 
      OperAmmesse = Session("swap_OperAmmesse")
   end if    
else
   v_attivi  = Request("attivi")
   v_Cessati = Request("cessati")
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   OperAmmesse   = getValueOfDic(Pagedic,"OperAmmesse")
end if 
if isAdmin() then 
   OperAmmesse="CRUD"
end if 
'registro i dati della pagina 
PaginaReturn="" 
xx=setValueOfDic(Pagedic,"PaginaReturn"  ,PaginaReturn)
xx=setValueOfDic(Pagedic,"OperAmmesse"   ,OperAmmesse)
xx=setCurrent(NomePagina,livelloPagina) 

if Oper="CALL_INS" or Oper="CALL_UPD" then
   xx=RemoveSwap()
   itemId = Cdbl("0" & Request("ItemToRemove"))
   Session("swap_OperTabella")     = Oper
   Session("swap_IdUtente")        = itemId
   Session("swap_PaginaReturn")    = "configurazioni/Collaboratori/" & NomePagina
   Session("swap_OperAmmesse")     = OperAmmesse
   response.redirect RitornaA("configurazioni/Collaboratori/UtenteCollModifica.asp")
   response.end 
   
end if  
 
if Oper="DEL" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	IdAccount ="0" & LeggiCampo("select IdAccount From Utente Where IdUtente=" & KK,"IdAccount")
	if Cdbl(IdAccount)>0 then 	
	   MyQ = "Delete From Account Where IdAccount=" & IdAccount
	   ConnMsde.execute MyQ
	   MyQ = "Delete From Utente Where IdAccount=" & IdAccount
       ConnMsde.execute MyQ

	End if 
	DescIn=""
End if
if Oper="ATT" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	IdAccount ="0" & LeggiCampo("select IdAccount From Utente Where IdUtente=" & KK,"IdAccount")
	if Cdbl(IdAccount)>0 then 
		MyQ = "" 
		MyQ = MyQ & " update Account set "
		MyQ = MyQ & "  FlagAttivo='N'"
		MyQ = MyQ & " ,Abilitato=0"
		MyQ = MyQ & " ,DescBlocco='Riattivato da " & apici(Session("LoginNominativo")) & " il " & Stod(Dtos()) & "'"
		MyQ = MyQ & " where IdAccount = " & IdAccount
		ConnMsde.execute MyQ 
		If Err.Number <> 0 Then 
			MsgErrore = ErroreDb(Err.description)
		End If
	End if 
	DescIn=""
End if 
%>

<div class="d-flex" id="wrapper">
	<%
	  Session("opzioneSidebar")="oper"
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
				   RiferimentoA="col-1 text-center;" & VirtualPath & paginaReturn & ";;2;prev;Indietro;;"
				   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				end if 
				%>
							
				
				<div class="col-11"><h3>Lista Back Office</h3>
				</div>
			</div>
			<%
			AddRow=true
			dim CampoDb(10)
			ElencoOption = ";0;Nominativo;1;Codice Fiscale;2;e-mail;3;"
            CampoDB(1)   = "DescUtente"
			CampoDb(2)   = "a.CodiceFiscale"
			CampoDb(3)   = "eMail"
			
			%>
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
<%
check_cessati=""
if v_cessati <> "" then 
   check_Cessati = "checked=""checked"""  
end if
%>

	<div class="row no-row-margin " style="margin-top: 10px;margin-bottom: 10px;" >

      <div class="col-1 s1 no-margin font-weight-bold">
	     mostra
	  </div>
	  
      <div class="col-1 no-margin">
      <label>
	    <input type="checkbox" name="cessati"  id="cessati" <%=check_cessati%> value="on">
        <span class="font-weight-bold">Cessati</span>
      </label>
	  </div>	
 
	  
	</div>
			
<%
			'caricamento tabella 
			if Condizione<>"" then 
				Condizione = " And " & Condizione
			end if 

			Set Rs = Server.CreateObject("ADODB.Recordset")
			MySql = "" 
			MySql = MySql & " Select a.*,b.FlagAttivo,b.DescBlocco From Utente A , Account B   "
			MySql = MySql & " Where a.IdAzienda = " & Session("IdAziendaWork") 
			MySql = MySql & " and   a.IdAccount = B.IdAccount" 
			MySql = MySql & " and   a.IdAccountLivello1 = " & Session("LoginIdAccount")
			if v_Cessati = "" then
		       MySql = MySql & " and b.FlagAttivo in ('S','N') "
			end if 
			MySql = MySql & Condizione & " order By DescUtente"

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
					<th scope="col"> Utente
						<%
						  RiferimentoA="col-2;#;;2;inse;Inserisci;;localIns();N"
						  %>
						<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					</th>
		<th scope="col">Cod.Fiscale</th>
		<th scope="col">Mail</th>
		<th scope="col">Stato</th>
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
					Id=Rs("IdUtente")
					DescLoaded=DescLoaded & Id & ";"
		%>
			<tr scope="col">
				<td>
					<input class="form-control" type="text" readonly value="<%=Rs("DescUtente")%>">
				</td>
				<td>
					<input class="form-control" type="text" readonly value="<%=Rs("CodiceFiscale")%>">
				</td>
				<td>
					<input class="form-control" type="text" readonly value="<%=Rs("eMail")%>">
				</td>				
				<td>
				    <%if Rs("FlagAttivo")="C" then 
					     FlagAttivo=Rs("DescBlocco")
					  elseif Rs("FlagAttivo")="S" then 
					     FlagAttivo="Attivo" 
					  else
					     FlagAttivo="Non Attivo:" & Rs("DescBlocco")
					end if 
					%>
					
					<input class="form-control" type="text" readonly value="<%=FlagAttivo%>">
				</td>
		
			<td>
			    <% if Rs("FlagAttivo")<>"C" then %>
					<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;localUpd('" & id & "');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					<%
					RiferimentoA="col-2;#;;2;dele;Cancella;;SalvaSingoloEdAttiva('DEL'," & Id & ",true,'','','');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
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
