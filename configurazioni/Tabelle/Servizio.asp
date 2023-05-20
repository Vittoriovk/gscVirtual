<%
  NomePagina="Servizio.asp"
  titolo="Configurazione Servizio"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
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

if FirstLoad then 
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
else
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
end if 

on error resume next 
if Oper="CALL_DOB" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
    Session("swap_IdAnagServizio") = KK
	Session("swap_PaginaReturn")  = "configurazioni/tabelle/Servizio.asp"
	response.redirect virtualPath & "configurazioni/tabelle/ServizioDocumentoCoobbligato.asp"
    response.end 
 
End if 
if Oper="CALL_FRI" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
    Session("swap_IdAnagServizio") = KK
	Session("swap_PaginaReturn")  = "configurazioni/tabelle/Servizio.asp"
	response.redirect virtualPath & "configurazioni/tabelle/ServizioFasciaRibasso.asp"
    response.end 
 
End if 

  'registro i dati della pagina 
xx=setValueOfDic(Pagedic,"PaginaReturn"   ,PaginaReturn)
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
			<%RiferimentoA="col-1 text-center;" & VirtualPath & "SupervisorConfigurazioni.asp;;2;prev;Indietro;;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h3>Elenco Servizi : Configurazioni</b></h3>
				</div>
			</div>

			<%
			AddRow=true
			dim CampoDb(10)
			CampoDB(1)="DescServizo"	
			ElencoOption=";0;Descrizione;1"
			%>		
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->

<%
'caricamento tabella 
if Condizione<>"" then 
	Condizione=" and " & Condizione
end if 
		
Set Rs = Server.CreateObject("ADODB.Recordset")

MySql = "" 
MySql = MySql & " Select A.*,B.DescAnagRamo "
MySql = MySql & " From AnagServizio A,AnagRamo B "
MySql = MySql & " Where A.IdAnagRamo = B.IdAnagRamo "
MySql = MySql & Condizione & " order By DescAnagServizio"

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
		<th scope="col">Servizio</th>
		<th scope="col">Ramo</th>
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
		Id=Rs("IdAnagServizio")
		err.clear 
		%>
		
		<tr scope="col">
			<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescAnagServizio")%>">
			</td>
			<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescAnagRamo")%>">
			</td>
			
            <td>
				<%if Id="CAUZ_PROV" then 
					RiferimentoA="col-2;#;;2;clie;Documenti Coobbligati;;AttivaFunzione('CALL_DOB','" & Id & "');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				<%end if %>
				<%if Id="CAUZ_DEFI" then 
					RiferimentoA="col-2;#;;2;perc;Fascie di Ribasso;;AttivaFunzione('CALL_FRI','" & Id & "');N"%>
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
