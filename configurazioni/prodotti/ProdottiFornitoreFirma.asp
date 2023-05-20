<%
  NomePagina="ProdottiFornitoreFirma.asp"
  titolo="Menu Supervisor - Dashboard"
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
	xx=ImpostaValoreDi("NameLoaded","Costo,FLZ");
	xx=ImpostaValoreDi("DescLoaded",id);
    xx=ElaboraControlli();
 	if (xx==false) {
	   return false;
	} 
	
	ImpostaValoreDi("ItemToRemove",id);
	ImpostaValoreDi("Oper","UPD");
	document.Fdati.submit();
}

</script>
</head>

<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<%
NameLoaded=NameLoaded & "Costo,LI"

Set Rs = Server.CreateObject("ADODB.Recordset")

IdProdotto   = 0
IdFornitore  = 0
IdAccount    = 0
DescFornitore= ""
DescProdotto = ""
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
      Rs.close 
   end if    
else
   IdProdotto     = "0" & getValueOfDic(Pagedic,"IdProdotto")
   IdAccount      = "0" & getValueOfDic(Pagedic,"IdAccount")
   IdFornitore    = "0" & getValueOfDic(Pagedic,"IdFornitore")
   DescProdotto   = getValueOfDic(Pagedic,"DescProdotto")
   DescFornitore  = getValueOfDic(Pagedic,"DescFornitore")
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
end if 
if cdbl(IdProdotto)=0 or Cdbl(IdFornitore)=0  then 
   response.redirect virtualPath & PaginaReturn
   response.end 
end if 

DescDettaglio= "<b>Fornitore</b>:" & DescFornitore & " - <b>Prodotto</b>:" & DescProdotto
IdAccount =cdbl(IdAccount)
IdProdotto=cdbl(IdProdotto)
on error resume next
 
if Oper="UPD" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	obbl  = Request("checkObb" & KK)
	costo = Request("Costo"    & KK)
	MyQ = "" 
	MyQ = MyQ & " delete from AccountProdottoFirma "
	MyQ = MyQ & " where IdProdotto = " & IdProdotto
	MyQ = MyQ & " and   IdAccount  = " & IdAccount 
	MyQ = MyQ & " and   IdTipoFirma = '" & apici(KK) & "'" 	
	ConnMsde.execute MyQ

	if obbl="S" then 
		MyQ = "" 
		MyQ = MyQ & " insert into AccountProdottoFirma (IdAccount,IdProdotto,IdTipoFirma,CostoFirma) values ( "
		MyQ = MyQ & "  " & IdAccount 
		MyQ = MyQ & ", " & IdProdotto
        MyQ = MyQ & ",'" & apici(kk) & "'"
		MyQ = MyQ & ", " & replace(costo ,",",".")
		MyQ = MyQ & " )" 
		ConnMsde.execute MyQ
	end if 

End if 
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdProdotto"     ,IdProdotto)
  xx=setValueOfDic(Pagedic,"DescProdotto"   ,DescProdotto)
  xx=setValueOfDic(Pagedic,"IdFornitore"    ,IdFornitore)
  xx=setValueOfDic(Pagedic,"DescFornitore"  ,DescFornitore)
  xx=setValueOfDic(Pagedic,"IdAccount"      ,IdAccount)
  xx=setValueOfDic(Pagedic,"PaginaReturn"   ,PaginaReturn)
  
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
				<div class="col-11"><h4>Gestione Firma per : <%=DescDettaglio%></h4>
				</div>
			</div>
<%

MySql = "" 
MySql = MySql & " Select a.*,isnull(b.CostoFirma,0) as Costo,isnull(b.IdAccount,0) as associato"
MySql = MySql & " from TipoFirma a  left join AccountProdottoFirma B "
MySql = MySql & " on    A.IdTipofirma = B.IdTipoFirma "
MySql = MySql & " and   B.IdProdotto = " & IdProdotto
MySql = MySql & " and   B.IdAccount = " & IdAccount
MySql = MySql & Condizione & " order By A.DescTipofirma"

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
		<th scope="col">Tipologia di firma</th>
		<th scope="col">Richiesto</th>
		<th scope="col" style="width:10%"  >Costo Euro</th>
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
		Id=Rs("IdTipoFirma")
		DescLoaded=DescLoaded & Id & ";"
		if Rs("associato")=0 then 
		   Obbligatorio=""
		else
		   Obbligatorio=" checked "
		end if		

		%> 
	<tr scope="col"> 
		<td>
			<input class="form-control" Id="IdDocumento<%=Id%>" type="text" readonly value="<%=Rs("DescTipoFirma")%>">
		</td>
		<td><div class="form-check">
				<input id="checkObb<%=Id%>" <%=Obbligatorio%> name="checkObb<%=Id%>" type="checkbox" value = "S" class="big-checkbox">
			</div>		
		</td>
		<td>
			<input class="form-control" Id="Costo<%=Id%>" Name="Costo<%=Id%>" type="text"  value="<%=Rs("Costo")%>">		
		</td>		

		<td>
			<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;localUpdate('UPD','" & Id & "');N"%>
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
