<%
  NomePagina="SubRamo.asp"
  titolo="Caratterizzazione di un ramo"
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
else
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
end if 

on error resume next 
if Oper="CALL_INS" or Oper="CALL_UPD" or Oper="CALL_DEL" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	Session("swap_IdSubRamo")     = KK
	Session("swap_OperTabella")   = Oper
    Session("swap_PaginaReturn")  = "configurazioni/tabelle/SubRamo.asp"
    response.redirect virtualPath & "configurazioni/tabelle/SubRamoDettaglio.asp"
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
	  Session("opzioneSidebar")="conf"
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
				<div class="col-11"><h3>Elenco Caratteristiche Rami</b></h3>
				</div>
			</div>

			<%
			AddRow=true
			dim CampoDb(10)
			CampoDB(1)="DescSubRamo"	
			CampoDB(2)="DescRamo"	
			ElencoOption=";0;Descrizione;1;Ramo;2"
			%>		
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->

<%
'caricamento tabella 
if Condizione<>"" then 
	Condizione=" and " & Condizione
end if 
		
Set Rs = Server.CreateObject("ADODB.Recordset")

MySql = "" 
MySql = MySql & " Select A.*,B.DescRamo "
MySql = MySql & " From SubRamo a inner join Ramo B "
MySql = MySql & " on A.IdRamo = B.IdRamo "
MySql = MySql & " Where A.IdSubRamo > 0 "
MySql = MySql & Condizione & " order By DescRamo"

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
		<th scope="col">Descrizione Caratteristica
		<a href="#" title="Inserisci" onclick="AttivaFunzione('CALL_INS','0');">
		<i class="fa fa-2x fa-plus-square"></i></a>
		</th>
        <th scope="col">Ramo Riferimento</th>
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
		Id=Rs("IdSubRamo")
		err.clear 
		%>
		
		<tr scope="col">
			<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescSubRamo")%>">
			</td>
			<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescRamo")%>">
			</td>			
            <td>
			
				<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;AttivaFunzione('CALL_UPD','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<%RiferimentoA="col-2;#;;2;dele;Cancella;;AttivaFunzione('CALL_DEL','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
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
