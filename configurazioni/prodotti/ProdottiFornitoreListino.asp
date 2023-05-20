<%
  NomePagina="ProdottiFornitoreListino.asp"
  titolo="Listino Prodotti fornitore"
  default_check_profile="SuperV,Admin,Coll"
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

IdAccount          = 0
IdProdotto         = 0
IdAccountFornitore = 0

DescFornitore= ""
DescProdotto = ""
if FirstLoad then 
   IdProdotto         = getCurrentValueFor("IdProdotto")
   IdAccount          = getCurrentValueFor("IdAccount")
   IdAccountFornitore = getCurrentValueFor("IdAccountFornitore")
   PaginaReturn       = getCurrentValueFor("PaginaReturn")   
   IdProdotto         = cdbl("0" & IdProdotto)
   IdAccount          = cdbl("0" & IdAccount)
   IdAccountFornitore = cdbl("0" & IdAccountFornitore)
   if cdbl(IdAccountFornitore)>0 then 
      Rs.CursorLocation = 3 
      Rs.Open "Select * from Fornitore where IdAccount=" & IdAccountFornitore, ConnMsde   
	  DescFornitore = Rs("DescFornitore")
      Rs.close 
   end if      
   if cdbl(IdProdotto)>0 then 
      Rs.CursorLocation = 3 
      Rs.Open "Select * from Prodotto where IdProdotto=" & IdProdotto, ConnMsde   
	  DescProdotto = Rs("DescProdotto")
      Rs.close 
   end if    
else
   IdProdotto         = "0" & getValueOfDic(Pagedic,"IdProdotto")
   IdAccount          = "0" & getValueOfDic(Pagedic,"IdAccount")
   IdAccountFornitore = "0" & getValueOfDic(Pagedic,"IdAccountFornitore")
   DescProdotto   = getValueOfDic(Pagedic,"DescProdotto")
   DescFornitore  = getValueOfDic(Pagedic,"DescFornitore")
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
end if 

   if Cdbl(IdAccount)=0 or cdbl(IdProdotto)=0 or Cdbl(IdAccountFornitore)=0  then 
      response.redirect virtualPath & PaginaReturn
      response.end 
   end if 

  IdAccount          = cdbl(IdAccount)
  IdProdotto         = cdbl(IdProdotto)
  IdAccountFornitore = cdbl(IdAccountFornitore)
  on error resume next


  if Oper="CALL_INS" or Oper="CALL_UPD" then 
     xx=RemoveSwap()
     Session("TimeStamp")=TimePage
     KK=trim(Request("ItemToRemove"))

     if kk <> "" then 
	    Session("swap_IdAccount")          = IdAccount
        Session("swap_IdProdotto")         = IdProdotto
		Session("swap_IdAccountFornitore") = IdAccountFornitore
        Session("swap_ValidoDal")          = KK
        Session("swap_PaginaReturn")  = "configurazioni/prodotti/" & nomePagina
        response.redirect virtualPath & "configurazioni/prodotti/ProdottiFornitoreListinoModifica.asp"
        response.end 
     end if 
  End if 

  if Oper="DEL" and checkTimePageLoad() then 
     Session("TimeStamp")=TimePage
     KK="0" & Request("ItemToRemove")
	 if Cdbl(KK)>0 then 
        MyQ = "" 
	    MyQ = MyQ & " delete from AccountProdottoListino "
	    MyQ = MyQ & " where ValidoDal = "          & KK
	    MyQ = MyQ & " and   IdProdotto = "         & IdProdotto
	    MyQ = MyQ & " and   IdAccount = "          & IdAccount
		MyQ = MyQ & " and   IdAccountFornitore = " & IdAccountFornitore
	
	    ConnMsde.execute MyQ 
	    If Err.Number <> 0 Then 
		   MsgErrore = ErroreDb(Err.description)
	    End If
     End if
  end if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdAccount"          ,IdAccount)
  xx=setValueOfDic(Pagedic,"IdProdotto"         ,IdProdotto)
  xx=setValueOfDic(Pagedic,"DescProdotto"       ,DescProdotto)
  xx=setValueOfDic(Pagedic,"IdAccountFornitore" ,IdAccountFornitore)
  xx=setValueOfDic(Pagedic,"DescFornitore"      ,DescFornitore)
  xx=setValueOfDic(Pagedic,"PaginaReturn"       ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 
  err.clear 
  
  DescAccount = ""
  if Cdbl(IdAccount)<>cdbl(IdAccountFornitore) then 
     DescAccount = LeggiCampo("select * from Account Where IdAccount =" & IdAccount,"Nominativo")
  end if 

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
				<div class="col-11"><h3>Listino Prodotto</h3>
				</div>
			</div>
	        <div class="row">
	           <div class="col-1">
	           </div>
			   <%if DescAccount<>"" then %>
               <div class="col-3 form-group ">
		          <%xx=ShowLabel("Listino per ")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescAccount%>" >
               </div>	
			   
			   <%end if %>
               <div class="col-3 form-group ">
		          <%xx=ShowLabel("Prodotto")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescProdotto%>" >
               </div>	
			   <%if IsCollaboratore()=false then %>
               <div class="col-3 form-group ">
		          <%xx=ShowLabel("Fornitore")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescFornitore%>" >
               </div>
			   <%end if %>
			</div>			
<%

MySql = "" 
MySql = MySql & " Select * From AccountProdottoListino "
MySql = MySql & " Where IdProdotto = " & IdProdotto
MySql = MySql & " and   IdAccount = " & IdAccount
MySql = MySql & " and   tipoRegola = '" & Session("LoginTipoUtente") & "'"
if IsCollaboratore() then 
   MySql = MySql & " And   IdAccountRegistratore = " & Session("LoginIdAccount")
else 
   MySql = MySql & " And   IdAccountRegistratore = 0"
end if 
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
		<th scope="col" width="15%">Valido Dal
        <a href="#" title="Inserisci" onclick="AttivaFunzione('CALL_INS','0');">
        <i class="fa fa-2x fa-plus-square"></i></a>		
		</th>
		<%if IsCollaboratore()=false then %>
		<th scope="col" width="15%">Prezzo Compagnia</th>
		<th scope="col" width="15%">Prezzo Fornitore</th>
		<%end if %>
		<th scope="col" width="15%">Prezzo Distribuzione</th>
		<th scope="col" width="15%">Prezzo Listino</th>
		<th scope="col">Azioni</th>
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
			<input class="form-control" type="text" readonly value="<%=Stod(Rs("ValidoDal"))%>">
		</td>
		<%if IsCollaboratore()=false then %>
		<td>
			<input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("PrezzoCompagnia"),2)%>">
		</td>		
		<td>
			<input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("PrezzoFornitore"),2)%>">
		</td>
		<%end if %>
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
