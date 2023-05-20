<%
  NomePagina="FornitoreListino.asp"
  titolo="Gestione Listino Fornitore"
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
<script>
function updMail() {
	xx=ImpostaValoreDi("NameLoaded","mail,EM");
	xx=ImpostaValoreDi("DescLoaded","0");
    xx=ElaboraControlli();
 	if (xx==false) {
	   return false;
	} 
	
	ImpostaValoreDi("Oper","UPDMAIL");
	document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<%
NameLoaded=NameLoaded & ""

Set Rs = Server.CreateObject("ADODB.Recordset")

IdAccountFornitore = 0

DescFornitore= ""
if FirstLoad then 
   IdAccountFornitore = getCurrentValueFor("IdAccountFornitore")
   PaginaReturn       = getCurrentValueFor("PaginaReturn")   
   IdAccountFornitore = cdbl("0" & IdAccountFornitore)
   if cdbl(IdAccountFornitore)>0 then 
      Rs.CursorLocation = 3 
      Rs.Open "Select * from Fornitore where IdAccount=" & IdAccountFornitore, ConnMsde   
	  DescFornitore = Rs("DescFornitore")
      Rs.close 
   end if      
else
   IdAccount          = "0" & getValueOfDic(Pagedic,"IdAccount")
   IdAccountFornitore = "0" & getValueOfDic(Pagedic,"IdAccountFornitore")
   DescFornitore  = getValueOfDic(Pagedic,"DescFornitore")
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
end if 

   if Cdbl(IdAccountFornitore)=0  then 
      response.redirect virtualPath & PaginaReturn
      response.end 
   end if 

  IdAccountFornitore = cdbl(IdAccountFornitore)

  on error resume next


  if Oper="CALL_INS" or Oper="CALL_UPD" then 
     xx=RemoveSwap()
     Session("TimeStamp")=TimePage
     KK=trim(Request("ItemToRemove"))

     if kk <> "" then 
	    Session("swap_IdAccountFornitore")       = IdAccountFornitore
        Session("swap_IdAccountProdottoListino") = kk
        Session("swap_PaginaReturn")             = "configurazioni/Tabelle/" & nomePagina
        response.redirect virtualPath & "configurazioni/Tabelle/FornitoreListinoModifica.asp"
        response.end 
     end if 
  End if 

  if Oper="DEL" and checkTimePageLoad() then 
     Session("TimeStamp")=TimePage
     KK="0" & Request("ItemToRemove")
	 if Cdbl(KK)>0 then 
        MyQ = "" 
	    MyQ = MyQ & " delete from AccountProdottoListino "
	    MyQ = MyQ & " where IdAccountProdottoListino = " & KK
	    ConnMsde.execute MyQ 
	    If Err.Number <> 0 Then 
		   MsgErrore = ErroreDb(Err.description)
	    End If
     End if
  end if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdAccountFornitore" ,IdAccountFornitore)
  xx=setValueOfDic(Pagedic,"DescFornitore"      ,DescFornitore)
  xx=setValueOfDic(Pagedic,"PaginaReturn"       ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 
  err.clear 
  
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
				<div class="col-11"><h3>Listino Prodotti Fornitore </h3>
				</div>
			</div>
	        <div class="row">
	           <div class="col-1">
	           </div>
               <div class="col-3 form-group ">
		          <%xx=ShowLabel("Fornitore")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescFornitore%>" >
               </div>
			</div>			
<%

MySql = "" 
MySql = MySql & " Select A.*"
MySql = MySql & " ,Isnull(B.DescProfiloProdotto,'') as DescProf"
MySql = MySql & " ,Isnull(C.DescProdotto,'') as DescProd"
MySql = MySql & " From AccountProdottoListino A "
MySql = MySql & " left join ProfiloProdotto B on A.IdProfiloProdotto = B.IdProfiloProdotto"
MySql = MySql & " left join Prodotto C on A.IdProdotto = C.IdProdotto"
MySql = MySql & " Where A.IdAccount = 0"
MySql = MySql & " And   A.IdAccountRegistratore = 0"
MySql = MySql & " And   A.tipoRegola = '" & Session("LoginTipoUtente") & "'"
MySql = MySql & " And   A.IdAccountFornitore = " & NumForDb(IdAccountFornitore)
MySql = MySql & Condizione & " order By ValidoDal Desc"

'response.write MySql 

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

RecCount=Rs.RecordCount 
if RecCount=0 then 
   RecCount=99
end if 

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
		<th scope="col" width="10%">Valido Dal
        <a href="#" title="Inserisci" onclick="AttivaFunzione('CALL_INS','0');">
        <i class="fa fa-2x fa-plus-square"></i></a>		
		</th>
		<th scope="col" >Gruppo Prodotto</th>		
		<th scope="col" >Prodotto</th>				
		<th scope="col" width="8%">Prezzo Compagnia</th>
		<th scope="col" width="8%">Prezzo Fornitore</th>
		<th scope="col" width="8%">Prezzo Distribuzione</th>
		<th scope="col" width="8%">Prezzo Listino</th>
		<th scope="col" width="6%">Azioni</th>
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
		Id=Rs("IdAccountProdottoListino")
		DescLoaded=DescLoaded & Id & ";"


		%> 
	<tr scope="col"> 
		<td>
			<input class="form-control" type="text" readonly value="<%=Stod(Rs("ValidoDal"))%>">
		</td>
		<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescProf")%>">
		</td>		
		<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescProd")%>">
		</td>		
		
		<td>
			<input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("PrezzoCompagnia"),2)%>">
		</td>		
		<td>
			<input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("PrezzoFornitore"),2)%>">
		</td>
		<td>
			<input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("PrezzoDistribuzione"),2)%>">
		</td>		
		<td>
			<input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("PrezzoListino"),2)%>">
		</td>		
		
		<td>
			<%RiferimentoA=";#;;2;upda;Modifica;;AttivaFunzione('CALL_UPD','" & Id & "');N"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->		
			<%RiferimentoA=";#;;2;dele;Cancella;;SalvaSingoloEdAttiva('DEL'," & Id & ",true,'','','');N"%>
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
