<%
  NomePagina="ProdottiListinoColl.asp"
  titolo="Menu Supervisor - Dashboard"
  default_check_profile="Coll"
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
NameLoaded=NameLoaded & ""

   Set Rs = Server.CreateObject("ADODB.Recordset")

   IdAccount          = Session("LoginIdAccount")
   if FirstLoad then 
      PaginaReturn       = getCurrentValueFor("PaginaReturn")   
   else
      PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
   end if 

  on error resume next


  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"PaginaReturn"       ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

%>

<div class="d-flex" id="wrapper">

	<%
	  TitoloNavigazione="utility"
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
				<div class="col-11"><h3>Listino Prodotto Personale</h3>
				</div>
			</div>
			<%
			AddRow=true
			dim CampoDb(10)
			ElencoOption = ";0;Prodotto;1"
            CampoDB(1)   = "DescProdotto"
			
			%>
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
			
			<%
			
if Condizione<>"" then 
	Condizione = " And " & Condizione
end if 

MySql = "" 
MySql = MySql & " Select A.*,B.DescProdotto "
MySql = MySql & " From AccountProdottoListino a, Prodotto B "
MySql = MySql & " Where IdAccount = " & IdAccount
MySql = MySql & " and A.IdProdotto = B.IdProdotto "
MySql = MySql & Condizione & " order By ValidoDal Desc"

'response.write MySql 

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde


DescLoaded=""
NumCols = 4
NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
%>
	<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

<div class="table-responsive">
 
<table class="table"><tbody>
<thead>
	<tr>
	    <th scope="col">Prodotto</th>
		<th scope="col">Valido Dal</th>
		<th scope="col">Prezzo Distribuzione</th>
		<th scope="col">Prezzo Listino</th>
	</tr>
</thead>

<%
elencoDettaglio=""
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
	Do While Not rs.EOF 
		Primo=Primo+1
		NumRec=NumRec+1
		Id=Rs("ValidoDal")
		DescLoaded=DescLoaded & Id & ";"


		%> 
	<tr scope="col"> 
		<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescProdotto")%>">
		</td>
		<td>
			<input class="form-control" type="text" readonly value="<%=Stod(Rs("ValidoDal"))%>">
		</td>
		<td>
			<input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("PrezzoDistribuzione"),2)%>">
		</td>		
		<td>
			<input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("PrezzoListino"),2)%>">
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
