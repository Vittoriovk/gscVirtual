<%
  NomePagina="ProdottiDisponibiliColl.asp"
  titolo="Menu Supervisor - Dashboard"
  default_check_profile="Coll,clie"
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
				<div class="col-11"><h3>Prodotti Disponibili</h3>
				</div>
			</div>
            <div class="row">
			   <div class="col-1 font-weight-bold">Servizio</div>
			   <div class="col-4">
			   <div class="form-group ">
				     <% 
					 IdAnagServizio=Request("IdAnagServizio0")
					 if IdAnagServizio="-1" then
					    IdAnagServizio=""
					 end if 
					 stdClass="class='form-control form-control-sm'"
					 q = ""
	                 q = q & " Select * from AnagServizio "
	                 q = q & " order by DescAnagServizio  "
                     response.write ListaDbChangeCompleta(q,"IdAnagServizio0",IdAnagServizio ,"IdAnagServizio","DescAnagServizio" ,1,"Sottometti();","","","","",stdClass)
                   %>
                  </div>		
			   
			   
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
MySql = MySql & " Select DescProdotto,Min(ValidoDal) as ValidoDal,Max(ValidoAl) as ValidoAl "
MySql = MySql & " From ProdottoSessione "
MySql = MySql & " Where IdAccount = " & Session("LoginIdAccount")
MySql = MySql & " and   IdSessione = '" & Session.SessionId & "'"
MySql = MySql & Condizione 
if IdAnagServizio<>"" then 
   MySql = MySql & " and IdAnagServizio = '" & apici(IdAnagServizio) & "'"
end if 
MySql = MySql & " Group By DescProdotto"
MySql = MySql & " order By DescProdotto"
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
		<th scope="col">Valido Al</th>
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
        If Rs("ValidoDal")=0 then 
		   ValidoDal="n.d."
		else 
		   ValidoDal=Stod(Rs("ValidoDal"))
        end if 
        If Rs("ValidoAl")=0 then 
		   ValidoAl="n.d."
		else 
		   ValidoAl=Stod(Rs("ValidoAl"))
        end if 		
		%> 
	<tr scope="col"> 
		<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescProdotto")%>">
		</td>
		<td>
			<input class="form-control" type="text" readonly value="<%=ValidoDal%>">
		</td>
		<td>
			<input class="form-control" type="text" readonly value="<%=ValidoAl%>">
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
