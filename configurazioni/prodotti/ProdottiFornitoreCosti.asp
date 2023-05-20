<%
  NomePagina="ProdottiFornitoreCosti.asp"
  titolo="costi prodotti per fascia"
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
<script>
function localUpdate(op,id) {
	xx=ImpostaValoreDi("DescLoaded",id);
    xx=ElaboraControlli();
 	if (xx==false) {
	   return false;
	} 
	
	ImpostaValoreDi("ItemToRemove",id);
	ImpostaValoreDi("Oper",op);
	document.Fdati.submit();
}

</script>
</head>

<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<%
NameLoaded=NameLoaded & "IdFascia,FLP;CostoFisso,FLZ;Percentuale,FLQ;Minimo,FLZ;"
Set Rs = Server.CreateObject("ADODB.Recordset")

IdProdotto   = 0
IdFornitore  = 0
IdAccount    = 0
DescFornitore    = ""
DescProdotto     = ""
DescTipoProdotto = ""
if FirstLoad then 
   IdProdotto   = "0" & Session("swap_IdProdotto")
   if Cdbl(IdProdotto)=0 then 
      IdProdotto = cdbl("0" & getValueOfDic(Pagedic,"IdProdotto"))
   end if 
   IdFornitore   = "0" & Session("swap_IdFornitore")
   if Cdbl(IdFornitore)=0 then 
      IdFornitore = cdbl("0" & getValueOfDic(Pagedic,"IdFornitore"))
   end if 
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
   IdFornitore = cdbl("0" & IdFornitore)
   if cdbl(IdFornitore)>0 then 
      Rs.CursorLocation = 3 
      Rs.Open "Select * from Fornitore where IdFornitore=" & IdFornitore, ConnMsde   
      IdAccount     = Rs("IdAccount")
	  DescFornitore = Rs("DescFornitore")
      Rs.close 
   end if      
   if cdbl(IdProdotto)>0 then 
      Rs.CursorLocation = 3 
      Rs.Open "Select * from Prodotto where IdProdotto=" & IdProdotto, ConnMsde   
	  DescProdotto = Rs("DescProdotto")
 
      if Rs("IdAnagServizio")="CAUZ_PROV" then 
	     DescTipoProdotto = "cauzione provvisoria"
	  end if 
	  if Rs("IdAnagServizio")="CAUZ_DEFI" then 
	     DescTipoProdotto = "cauzione definitiva"
	  end if 
      Rs.close 
   end if    
else
   IdProdotto       = "0" & getValueOfDic(Pagedic,"IdProdotto")
   IdAccount        = "0" & getValueOfDic(Pagedic,"IdAccount")
   IdFornitore      = "0" & getValueOfDic(Pagedic,"IdFornitore")
   DescProdotto     = getValueOfDic(Pagedic,"DescProdotto")
   DescTipoProdotto = getValueOfDic(Pagedic,"DescTipoProdotto")
   DescFornitore    = getValueOfDic(Pagedic,"DescFornitore")
   PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
end if 
if cdbl(IdProdotto)=0 or Cdbl(IdFornitore)=0  then 
   response.redirect virtualPath & PaginaReturn
   response.end 
end if 

IdFascia   =0
CostoFisso =0
CostoPrec  =0
Percentuale=0
Minimo     =0

IdAccount =cdbl(IdAccount)
IdProdotto=cdbl(IdProdotto)
on error resume next

if Oper="INS" then 
   Session("TimeStamp")=TimePage
   KK=0
   IdFascia   = request("IdFascia"    & kk)
   CostoFisso = request("CostoFisso"  & kk)
   CostoPrec  = request("CostoPrec"   & kk)
   CostoPrec  = 0
   Percentuale= request("Percentuale" & kk)
   Minimo     = request("Minimo"      & kk)

   MyQ = ""
   MyQ = MyQ & " Insert into AccountProdottoFascia (IdAccount,IdProdotto,IdFascia,CostoFisso,CostoFissoFasciaPrec,Percentuale,Minimo) "
   MyQ = MyQ & " values ("
   MyQ = MyQ & "  " & IdAccount
   MyQ = MyQ & " ," & IdProdotto
   MyQ = MyQ & " ," & NumForDb(IdFascia)
   MyQ = MyQ & " ," & NumForDb(CostoFisso)
   MyQ = MyQ & " ," & NumForDb(CostoPrec)
   MyQ = MyQ & " ," & NumForDb(Percentuale)
   MyQ = MyQ & " ," & NumForDb(Minimo)
   MyQ = MyQ & " )"
   
   ConnMsde.execute MyQ
   if err.number=0 then 
      IdFascia   =0
      CostoFisso =0
      CostoPrec  =0
      Percentuale=0
      Minimo     =0
   else
      MsgErrore = ErroreDb(err.description) 
   end if 

end if 

if Oper="MOD" then 
   Session("TimeStamp")=TimePage
   KK=Request("ItemToRemove")
   IdRow      = KK
   IdFascia   = request("IdFascia"    & kk)
   CostoFisso = request("CostoFisso"  & kk)
   CostoPrec  = request("CostoPrec"   & kk)
   Percentuale= request("Percentuale" & kk)
   Minimo     = request("Minimo"      & kk)	
   
   MyQ = "" 
   MyQ = MyQ & " update AccountProdottoFascia set "
   MyQ = MyQ & " IdFascia = "    & NumForDb(IdFascia)
   MyQ = MyQ & ",CostoFisso = "  & NumForDb(CostoFisso)
   MyQ = MyQ & ",CostoFissoFasciaPrec = " 
   MyQ = MyQ & ""                & NumForDb(CostoPrec)
   MyQ = MyQ & ",Percentuale = " & NumForDb(Percentuale)
   MyQ = MyQ & ",Minimo = "      & NumForDb(Minimo)
   MyQ = MyQ & " where IdRow = " & Idrow
   ConnMsde.execute MyQ
   if err.number<>0 then 
      MsgErrore = ErroreDb(err.description) 
   end if    
   IdFascia   =0
   CostoFisso =0
   CostoPrec  =0
   Percentuale=0
   Minimo     =0

End if 

if Oper="DEL" then 
   Session("TimeStamp")=TimePage
   KK=Request("ItemToRemove")
   IdRow      = KK
   
   MyQ = "" 
   MyQ = MyQ & " delete from AccountProdottoFascia "
   MyQ = MyQ & " where IdRow = " & Idrow
   ConnMsde.execute MyQ
   IdFascia   =0
   CostoFisso =0
   CostoPrec  =0
   Percentuale=0
   Minimo     =0
 
End if 

  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdProdotto"      ,IdProdotto)
  xx=setValueOfDic(Pagedic,"DescProdotto"    ,DescProdotto)
  xx=setValueOfDic(Pagedic,"DescTipoProdotto",DescTipoProdotto)
  
  xx=setValueOfDic(Pagedic,"IdFornitore"     ,IdFornitore)
  xx=setValueOfDic(Pagedic,"DescFornitore"   ,DescFornitore)
  xx=setValueOfDic(Pagedic,"IdAccount"       ,IdAccount)
  xx=setValueOfDic(Pagedic,"PaginaReturn"    ,PaginaReturn)
  
  xx=setCurrent(NomePagina,livelloPagina) 
  err.clear 
%>

<% 
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
			<%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h4>Costi per <%=DescTipoProdotto%></h4>
				</div>
			</div>
			<div class="row">
			   <div class="col-1"></div>
			   <div class="col-4">
                  <div class="form-group ">
				     <%xx=ShowLabel("Fornitore")%>
					 <input type="text" readonly class="form-control" value="<%=DescFornitore%>" >
                  </div>		
			   </div>
			   <div class="col-4">
                  <div class="form-group ">
				     <%xx=ShowLabel("Prodotto")%>
					 <input type="text" readonly class="form-control" value="<%=DescProdotto%>" >
                  </div>		
			   </div>			   
			</div>
			<br>
<%

MySql = "" 
MySql = MySql & " Select *"
MySql = MySql & " from AccountProdottoFascia "
MySql = MySql & " where IdProdotto = " & IdProdotto
MySql = MySql & " and   IdAccount = " & IdAccount
MySql = MySql & " order by IdFascia"

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
	
<div class="table-responsive"><table class="table"><tbody>
<thead>
	<tr>
		<th scope="col">Fascia Fino A &euro;</th>
		<th scope="col">Costo Fisso &euro;</th>
		<th scope="col" style="width:15%"  >Perc. Anno %</th>
		<th scope="col" style="width:15%"  >Minimo &euro;</th>
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
		Id=Rs("IdRow")
		DescLoaded=DescLoaded & Id & ";"

		%> 
	<tr scope="col"> 
		<td>
			<input class="form-control" Id="IdFascia<%=Id%>"    Name="IdFascia<%=Id%>"   type="text" value="<%=Rs("IdFascia")%>">
		</td>
		<td>
			<input class="form-control" Id="CostoFisso<%=Id%>"  Name="CostoFisso<%=Id%>" type="text" value="<%=Rs("CostoFisso")%>">
		</td>
		<td>
			<input class="form-control" Id="Percentuale<%=Id%>" Name="Percentuale<%=Id%>"  type="text" value="<%=Rs("Percentuale")%>">
		</td>		
		<td>
			<input class="form-control" Id="Minimo<%=Id%>"      Name="Minimo<%=Id%>"  type="text" value="<%=Rs("Minimo")%>">
		</td>		
		<td>
			<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;localUpdate('MOD','" & Id & "');N"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
			<%RiferimentoA="col-2;#;;2;dele;Cancella;;localUpdate('DEL','" & Id & "');N"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->						
		</td>		
	</tr> 
		<%	
		rs.MoveNext
	Loop
end if 
rs.close

    id=0
%>
	<tr scope="col"> 
		<td>
			<input type="text" class="form-control" Id="IdFascia<%=Id%>"    Name="IdFascia<%=Id%>"      value="<%=IdFascia%>">
		</td>
		<td>
			<input type="text" class="form-control" Id="CostoFisso<%=Id%>"  Name="CostoFisso<%=Id%>"    value="<%=CostoFisso%>">
		</td>
		<td>
			<input type="text" class="form-control" Id="Percentuale<%=Id%>" Name="Percentuale<%=Id%>"   value="<%=Percentuale%>">
		</td>		
		<td>
			<input type="text" class="form-control" Id="Minimo<%=Id%>"      Name="Minimo<%=Id%>"        value="<%=Minimo%>">
		</td>		
		<td>
			<%RiferimentoA="col-2;#;;2;inse;Inserisci;;localUpdate('INS','" & Id & "');N"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</td>		
	</tr> 
   
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
