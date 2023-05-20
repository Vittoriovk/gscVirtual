<%
  NomePagina="ProdottoFornitore.asp"
  titolo="Prodotti per fornitore "
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
<!-- Custom styles for this  -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">
</head>

<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->


<%

Set Rs = Server.CreateObject("ADODB.Recordset")
IdFornitore=0
IdAccount  =0
DescFornitore=""

if FirstLoad then 
   IdFornitore   = Session("swap_IdFornitore")
   IdAccount     = Session("swap_IdAccount")
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
   IdFornitore = cdbl("0" & IdFornitore)
   if cdbl(IdFornitore)= 0 then
      IdFornitore   = cdbl("0" & getValueOfDic(Pagedic,"IdFornitore"))
   end if 
   
   if cdbl(IdFornitore)>0 then 
      Rs.CursorLocation = 3 
      Rs.Open "Select * from Fornitore where IdFornitore=" & IdFornitore, ConnMsde   
      IdAccount     = Rs("IdAccount")
	  DescFornitore = Rs("DescFornitore")
      Rs.close 
   end if    
   
else
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   IdFornitore   = getValueOfDic(Pagedic,"IdFornitore")
   DescFornitore = getValueOfDic(Pagedic,"DescFornitore")
   IdAccount     = getValueOfDic(Pagedic,"IdAccount")
end if 

if cdbl(IdFornitore)=0 then 
   response.redirect virtualpath & PaginaReturn 
   response.end 
end if 

on error resume next 
if Oper="CALL_LIST" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	Session("swap_IdAccount")          = IdAccount
	Session("swap_IdProdotto")         = KK
	Session("swap_IdAccountFornitore") = IdAccount
	Session("swap_OperTabella")   = Oper
    Session("swap_PaginaReturn")  = "configurazioni/prodotti/ProdottoFornitore.asp"
    response.redirect virtualPath & "configurazioni/prodotti/ProdottiFornitoreListino.asp"	
end if 

if Oper="CALL_TECN" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
    KK=Request("ItemToRemove")
    if Cdbl("0" & KK ) > 0 then 
        Session("swap_IdProdotto")      = KK
        Session("swap_IdTabellaKeyInt") = KK
        Session("swap_OperTabella")   = Oper
        Session("swap_PaginaReturn")  = "configurazioni/prodotti/ProdottoFornitore.asp"
        response.redirect virtualPath & "configurazioni/prodotti/ProdottoListaDati.asp"
        response.end 
    end if 
End if 

if Oper="CALL_AFFI" or Oper="DATI_TECN" or Oper="CALL_DOCU" or Oper="CALL_FIRM" or Oper="CALL_COST" or Oper="CALL_MAIL" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	Session("swap_IdFornitore")  = IdFornitore
	Session("swap_IdProdotto")   = KK
	Session("swap_OperTabella")   = Oper
    Session("swap_PaginaReturn")  = "configurazioni/prodotti/ProdottoFornitore.asp"
	if Oper="CALL_AFFI" then 
	   Session("swap_TipoDoc")    = "AFFI"
       response.redirect virtualPath & "configurazioni/prodotti/ProdottiFornitoreDocAff.asp"
	elseif Oper="CALL_DOCU" then 
	   Session("swap_TipoDoc")    = "PROD"
	   response.redirect virtualPath & "configurazioni/prodotti/ProdottiFornitoreDocAff.asp"
    elseif Oper="CALL_FIRM" then 
	   response.redirect virtualPath & "configurazioni/prodotti/ProdottiFornitoreFirma.asp"
    elseif Oper="CALL_MAIL" then 
	   response.redirect virtualPath & "configurazioni/prodotti/ProdottiFornitoreMail.asp"
    elseif Oper="DATI_TECN" then 
	   response.redirect virtualPath & "configurazioni/prodotti/ProdottiFornitoreDatoTecn.asp"
    else 
	   response.redirect virtualPath & "configurazioni/prodotti/ProdottiFornitoreCosti.asp"
	end if 
    response.end 
End if 

  'registro i dati della pagina 
xx=setValueOfDic(Pagedic,"PaginaReturn"  ,PaginaReturn)
xx=setValueOfDic(Pagedic,"IdFornitore"   ,IdFornitore)
xx=setValueOfDic(Pagedic,"DescFornitore" ,DescFornitore)
xx=setValueOfDic(Pagedic,"IdAccount"     ,IdAccount)
xx=setCurrent(NomePagina,livelloPagina) 
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
			<%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn &  ";;2;prev;Indietro;;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h3>Elenco Prodotti per fornitore: <%=DescFornitore%></h3>
				</div>
			</div>

			<%
			AddRow=true
			dim CampoDb(10)
			CampoDB(1)="DescProdotto"	
			CampoDB(2)="DescCompagnia"	
			ElencoOption=";0;Descrizione;1;Compagnia;2"
			%>		
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->

<%
'caricamento tabella 
if Condizione<>"" then 
	Condizione=" and " & Condizione
end if 
		


MySql = "" 
MySql = MySql & " Select a.*,B.MailDocumentazione,b.CodiceProdotto as CodProF,C.DescCompagnia "
MySql = MySql & " From Prodotto a,AccountProdotto B,Compagnia C "
MySql = MySql & " Where a.IdProdotto > 0 "
MySql = MySql & " and   a.IdProdotto  = B.IdProdotto "
MySql = MySql & " and   B.IdAccount   = " & IdAccount
MySql = MySql & " and   a.IdCompagnia = C.idCompagnia "
MySql = MySql & Condizione & " order By DescProdotto"

'response.write MySql 

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

DescLoaded=""
NumCols = numC + 1
NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
%>

<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

<div class="table-responsive"><table class="table"><tbody>
<thead>
	<tr>
		<th scope="col">Prodotto</th>
		<th scope="col">Compagnia</th>
		<th scope="col">Codice</th>
		<th scope="col">Mail documenti affidamento/gestione</th>
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
		Id=Rs("IdProdotto")
		MailDoc=Rs("MailDocumentazione")
		err.clear 
		%>
		
		<tr scope="col">
			<td>
			<input class="form-control" type="text" name="Desc<%=Id%>" readonly value="<%=Rs("DescProdotto")%>">
			</td>
			<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescCompagnia")%>">
			</td>
			<td>
			<input class="form-control" type="text" readonly value="<%=Rs("CodProF")%>">
			</td>			
			<td>
			<input class="form-control" type="text" readonly value="<%=MailDoc%>">
			</td>			
            <td>
			
				<%RiferimentoA="col-2;#;;2;docu;Documenti di prodotto;;AttivaFunzione('CALL_DOCU','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				
				<%RiferimentoA="col-2;#;;2;tecn;Configura parametri;;AttivaFunzione('CALL_MAIL','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				
				<% if Instr("CAUZ_DEFI,CAUZ_PROV",Rs("IdAnagServizio"))=0 and trim(Rs("FlagPrezzoFisso"))="1" then
				      RiferimentoA="col-2;#;;2;money;Listino prodotto;;AttivaFunzione('CALL_LIST','" & Id & "');N"%>
				      <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				<% end if %>
				
				
				<% if Rs("IdAnagServizio")="CAUZ_DEFI" then %>
				<%RiferimentoA="col-2;#;;2;hand;Documentazione per Istruttoria;;AttivaFunzione('CALL_AFFI','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<%end if%>
				
				<% if Rs("IdAnagServizio")="CAUZ_PROV" then %>
				<%RiferimentoA="col-2;#;;2;hand;Documentazione per affidamento;;AttivaFunzione('CALL_AFFI','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<%RiferimentoA="col-2;#;;2;penn;Firme;;AttivaFunzione('CALL_FIRM','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				<%end if%>
				<% if Instr("CAUZ_PROV|CAUZ_DEFI",Rs("IdAnagServizio"))>0 then %>
				<%RiferimentoA="col-2;#;;2;money;Costi cauzione;;AttivaFunzione('CALL_COST','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				
				<%end if%>
				
				<% if Rs("IdAnagServizio")="FORMAZ" then %>
				<%RiferimentoA="col-2;#;;2;hand;Documentazione per attivazione;;AttivaFunzione('CALL_DOCU','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				<% end if %>
				
                <%RiferimentoA="col-2;#;;2;manu;Dati Aggiuntivi;;AttivaFunzione( 'CALL_TECN','" & Id & "');N"%>
                <!--#include virtual="/gscVirtual/include/Anchor.asp"-->    
					  
				<% if Rs("IdAnagServizio")="FORMAZ" and false then %>
				<%RiferimentoA="col-2;#;;2;manu;Dati Tecnici Aggiuntivi;;AttivaFunzione('DATI_TECN','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				<%end if %>
			<td>
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
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

</body>

</html>
