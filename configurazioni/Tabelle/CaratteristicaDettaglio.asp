<%
  NomePagina="CaratteristicaDettaglio.asp"
  titolo="Anagrafica Template Servizi"
  default_check_profile="SuperV"
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

if FirstLoad then 
   PaginaReturn    = getCurrentValueFor("PaginaReturn")
   IdAnagServizio  = getCurrentValueFor("IdAnagServizio")
else
   PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
   IdAnagServizio   = getValueOfDic(Pagedic,"IdAnagServizio")
end if 
'response.write "qq=" & IdAnagServizio
'response.end 

on error resume next 
if Oper="CALL_INS" or Oper="CALL_UPD" or Oper="CALL_DEL" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	Session("swap_OperTabella")          = Oper
	Session("swap_IdAnagServizio")       = IdAnagServizio
	Session("swap_IdAnagCaratteristica") = KK
    Session("swap_PaginaReturn")         = "configurazioni/tabelle/CaratteristicaDettaglio.asp"
    response.redirect virtualPath & "configurazioni/tabelle/CaratteristicaDettaglioModifica.asp"
    response.end 
End if 

  'registro i dati della pagina 
  
xx=setValueOfDic(Pagedic,"PaginaReturn"    ,PaginaReturn)
xx=setValueOfDic(Pagedic,"IdAnagServizio"  ,IdAnagServizio)
xx=setCurrent(NomePagina,livelloPagina) 
'xx=DumpDic(SessionDic,NomePagina)

DescAnagServizio = LeggiCampo("select * from AnagServizio where IdAnagServizio='" & IdAnagServizio & "'","DescAnagServizio")

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
			<%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h3>Elenco Template</b></h3>
				</div>
			</div>
	        <div class="row">
	           <div class="col-1">
	           </div>
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("Servizio")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescAnagServizio%>" >
               </div>	
			</div>
			<%
			AddRow=true
			dim CampoDb(10)
			CampoDB(1)="DescAnagCaratteristica"	
			ElencoOption=";0;Descrizione;1"
			%>		
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->

<%
'caricamento tabella 
if Condizione<>"" then 
	Condizione=" and " & Condizione
end if 

if FirstLoad=false then		
Set Rs = Server.CreateObject("ADODB.Recordset")

MySql = "" 
MySql = MySql & " Select * "
MySql = MySql & " From AnagCaratteristica "
MySql = MySql & " Where IdAnagServizio='" & idAnagServizio & "'"
MySql = MySql & Condizione & " order By DescAnagCaratteristica"

'serve per velocizzare il reload 




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

<div class="table-responsive">
    <table class="table">
<tbody>
<thead>
	<tr>
		<th data-field="descrizione" data-filter-control="input" scope="col">Descrizione Template
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
		Id=Rs("IdAnagCaratteristica")
		err.clear 
		%>
		
		<tr scope="col">
			<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescAnagCaratteristica")%>">
			</td>
	
            <td><%if rs("flagModificabile")=1 then %>
				<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;AttivaFunzione('CALL_UPD','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			

				
				<%RiferimentoA="col-2;#;;2;dele;Cancella;;AttivaFunzione('CALL_DEL','" & Id & "');N"%>
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
<%end if %>
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
