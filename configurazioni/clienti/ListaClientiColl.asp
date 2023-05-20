<%
  NomePagina="ListaClientiColl.asp"
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
function attivaForm()
{
	xx=$('#confirmModal').modal('toggle');
}

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
function localConf(id,op)
{
	xx=ImpostaValoreDi("ItemToRemove",id);
	xx=ImpostaValoreDi("Oper",op);
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
OperAmmesse="CRUD"
'registro i dati della pagina 
PaginaReturn=""
if PaginaReturn="" and false then 
   PaginaReturn="link/AdminIntermediari.asp"
end if 
xx=setValueOfDic(Pagedic,"PaginaReturn"  ,PaginaReturn)
xx=setValueOfDic(Pagedic,"OperAmmesse"   ,OperAmmesse)
xx=setCurrent(NomePagina,livelloPagina) 

if Oper="CALL_PRO" then 
   xx=RemoveSwap()
   IdCliente = Cdbl("0" & Request("ItemToRemove"))
   IdAccount = cdbl("0" & LeggiAccount("Cliente",IdCliente))
   
   if Cdbl(IdAccount)>0 then 
      Session("swap_IdAccountPadre") = Session("LoginIdAccount") 
	  Session("swap_IdAccount")      = IdAccount
      Session("swap_PaginaReturn")   = "configurazioni/Clienti/" & NomePagina
      response.redirect RitornaA("configurazioni/prodotti/ProdottiAccount.asp")
      response.end 
	  
   end if    

end if 

if Oper="CALL_INS" or Oper="CALL_UPD" then
   xx=RemoveSwap()
   itemId = Cdbl("0" & Request("ItemToRemove"))
   tipoId = ""
   tipoCo = ""
   if Oper="CALL_INS" then 
      tipoId = Request("gruppo1")
	  if tipoId="SEGN" then 
	     tipoId="PEFI"
		 tipoCO="Segn"
	  end if 
	  Session("swap_IdTipoCliente") = tipoCO
	  Session("swap_IdPersCliente") = tipoID
   end if    
   Session("swap_OperTabella")     = Oper
   Session("swap_IdCliente") = itemId
   Session("swap_PaginaReturn")    = "configurazioni/Clienti/" & NomePagina
   Session("swap_OperAmmesse")     = OperAmmesse
   response.redirect RitornaA("configurazioni/Clienti/ClienteModifica.asp")
   response.end 
   
end if  
if Oper="CALL_CFG" then
   xx=RemoveSwap()
   itemId = Cdbl("0" & Request("ItemToRemove"))
   Session("swap_IdCliente") = itemId
   Session("swap_PaginaReturn")    = "configurazioni/Clienti/" & NomePagina
  
   response.redirect RitornaA("configurazioni/Clienti/ClienteConfigura.asp")
   response.end 
end if  
if Oper="CALL_LOG" then
   xx=RemoveSwap()
   itemId = Cdbl("0" & Request("ItemToRemove"))
   Session("swap_IdCliente") = itemId
   Session("swap_PaginaReturn")    = "configurazioni/Clienti/" & NomePagina
  
   response.redirect RitornaA("configurazioni/Clienti/ClienteConfiguraLogin.asp")
   response.end 
end if  
if Oper="CALL_CER" then
   xx=RemoveSwap()
   itemId = Cdbl("0" & Request("ItemToRemove"))
   Session("swap_IdCliente")    = itemId
   Session("swap_PaginaReturn") = "configurazioni/Clienti/" & NomePagina
  
   response.redirect RitornaA("configurazioni/Clienti/ClienteCertificati.asp")
   response.end 
end if  
if Oper="CALL_COO" then
   xx=RemoveSwap()
   itemId = Cdbl("0" & Request("ItemToRemove"))
   Session("swap_IdCliente")    = itemId
   Session("swap_PaginaReturn") = "configurazioni/Clienti/" & NomePagina
  
   response.redirect RitornaA("configurazioni/Clienti/ClienteCoobbligati.asp")
   response.end 
end if 

if Oper="DEL" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	Acc = Cdbl("0" & LeggiCampo("select IdAccount from Cliente where IdCliente=" & kk,"IdAccount"))
	if Acc>0 then 
	   MyQ = "" 
	   MyQ = MyQ & " update Account Set "
	   MyQ = MyQ & " FlagAttivo='N',Abilitato=0"
	   MyQ = MyQ & ",DescBlocco='Cancellato da " & apici(Session("LoginNominativo")) & " il " & Stod(Dtos()) & "'"
	   MyQ = MyQ & " where IdAccount = " & Acc
	   MsgErrore=VerificaDel("Cliente",KK)
	   if MsgErrore="" then 	
		  ConnMsde.execute MyQ 
		  If Err.Number <> 0 Then 
			MsgErrore = ErroreDb(Err.description)
		  End If
	   end if 
	End if 
	DescIn=""
End if
if Oper="ATT" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	Acc = Cdbl("0" & LeggiCampo("select IdAccount from Cliente where IdCliente=" & kk,"IdAccount"))
	if Acc>0 then 	
       MyQ = "" 
	   MyQ = MyQ & " update Account set "
	   MyQ = MyQ & "  FlagAttivo='S'"
	   MyQ = MyQ & " ,DescBlocco='Riattivato da " & apici(Session("LoginNominativo")) & " il " & Stod(Dtos()) & "'"
	   MyQ = MyQ & " where IdAccount = " & Acc
       ConnMsde.execute MyQ 
	End if 
	DescIn=""
End if
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
				<%
                RiferimentoA="col-1  text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"
				if PaginaReturn<>"" then 
				%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<%end if %>
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
check_Attivi=""
if v_attivi <> "" then 
   check_Attivi = "checked=""checked"""  
end if
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
	    <input type="checkbox" name="attivi"  id="attivi" <%=check_attivi%> value="on">
        <span class="font-weight-bold">Attivi</span>
      </label>
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
			if (v_attivi <>"" and v_Cessati = "") or (v_attivi = "" and v_Cessati <> "") then
			   if v_attivi <>"" then
			       MySql = MySql & " and c.FlagAttivo<>'N' "
			   else
			       MySql = MySql & " and c.FlagAttivo='N'"
			   end if 
			end if 
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
					<th scope="col"> Cliente
						<%
						  if instr(OperAmmesse,"C")>0  then 
						  RiferimentoA="col-2;#;;2;inse;Inserisci;;attivaForm();N"
						  %>
						<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                          <%end if %>						
					</th>
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
					<%RiferimentoA="col-2;#;;2;tecn;Configurazioni;;localConf('" & id & "','CALL_CFG');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					<%RiferimentoA="col-2;#;;2;logi;Dati di Accesso;;localConf('" & id & "','CALL_LOG');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					<%RiferimentoA="col-2;#;;2;cert;Certificazioni;;localConf('" & id & "','CALL_CER');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                    <%if false then%>					
					<%RiferimentoA="col-2;#;;2;clie;Coobligati;;localConf('" & id & "','CALL_COO');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->						
					<%end if %>
				    <%RiferimentoA="col-2;#;;2;prod;Prodotti;;AttivaFunzione('CALL_PRO','" & Id & "');N"
					  abilitato=GetDiz(session("Login_Parametri") ,"ASS_PRO")
					  if abilitato="S" then 
						%>
				       <!--#include virtual="/gscVirtual/include/Anchor.asp"-->											
					<%end if 
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

<div class="modal fade" id="confirmModal"  aria-hidden="true" role="dialog">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">

        <h2>Seleziona Tipo Cliente </h2> 
        <button type="button" class="close" data-dismiss="modal">
          <span aria-hidden="true">×</span><span class="sr-only">Chiudi</span>
        </button>
      </div>

      <div class="modal-body"> 
		<div>
		  <div class="form-check">
			<input name="gruppo1" type="radio" id="radio2" value="PEGI" >
			<label for="radio2">Persona giuridica</label>
		  </div>
		  <div class="form-check">
			<input name="gruppo1" type="radio" id="radio3" value="DITT">
			<label for="radio3">Ditta individuale</label>
		  </div>		  
		</div>		  
      </div> 

      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
        <button type="button" class="btn btn-primary" onclick="localIns();";>Seleziona</button>
      </div>
    </div>
  </div>
</div>
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
