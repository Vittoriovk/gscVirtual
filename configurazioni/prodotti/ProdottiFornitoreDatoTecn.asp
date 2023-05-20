<%
  NomePagina="ProdottiFornitoreDatoTecn.asp"
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
NameLoaded=NameLoaded & "IdDatoTecnico,LI"

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
 
FlagUpdRiferimento=false 

if Oper="INS" then 
    Session("TimeStamp")=TimePage
	KK="0"
	IdDatoTecnico = Request("IdDatoTecnico" & KK)
	if Cdbl(IdDatoTecnico)>0 and cdbl(IdProdotto)>0 then 
		MyQ = "" 
		MyQ = MyQ & " Insert into AccountProdottoDatoTecn ("
		MyQ = MyQ & " IdAccount,IdProdotto,IdDatoTecnico"
		MyQ = MyQ & ") values ("		
		MyQ = MyQ & "  " & IdAccount
		MyQ = MyQ & " ," & IdProdotto  		
		MyQ = MyQ & " ," & IdDatoTecnico 
		MyQ = MyQ & ")"

		ConnMsde.execute MyQ 
		If Err.Number <> 0 Then 
			MsgErrore = ErroreDb(Err.description)
		else
			DescIn=""
		End If
	END if 
End if 

if Oper="DEL" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	MyQ = "" 
	MyQ = MyQ & " delete from AccountProdottoDatoTecn "
	MyQ = MyQ & " where IdDatoTecnico = " & KK
	MyQ = MyQ & " and   IdProdotto = " & IdProdotto
	MyQ = MyQ & " and   IdAccount = "  & IdAccount
	
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	End If
	DescIn=""
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

  DescMail=LeggiCampo("select * from AccountProdotto Where IdProdotto=" & IdProdotto & " and IdAccount=" & IdAccount,"MailDocumentazione")
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
				<div class="col-11"><h5>Dati Tecnici Aggiuntivi <br> <%=DescDettaglio%></h5>
				</div>
			</div>

			<%
			AddRow=true
			dim CampoDb(10)
			CampoDB(1)="DescDatoTecnico"	
			ElencoOption=";0;Descrizione;1"
			%>		
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
<%
'caricamento tabella 
if Condizione<>"" then 
	Condizione=" and " & Condizione
end if 
		


MySql = "" 
MySql = MySql & " Select b.* From AccountProdottoDatoTecn a, DatoTecnico B "
MySql = MySql & " Where A.IdDatoTecnico = B.IdDatoTecnico "
MySql = MySql & " and   A.IdProdotto = " & IdProdotto
MySql = MySql & " and   A.IdAccount = " & IdAccount
MySql = MySql & Condizione & " order By B.DescDatoTecnico"

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
		<th scope="col">Dato Tecnico</th>
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
		Id=Rs("IdDatoTecnico")
		DescLoaded=DescLoaded & Id & ";"


		%> 
	<tr scope="col"> 
		<td>
			<input class="form-control" Id="IdDatoTecnico<%=Id%>" type="text" readonly value="<%=Rs("DescDatoTecnico")%>">
		</td>
		<td>
			<%RiferimentoA="col-2;#;;2;dele;Cancella;;SalvaSingoloEdAttiva('DEL'," & Id & ",true,'','','');N"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</td>		
	</tr> 
		<%	
		rs.MoveNext
	Loop
end if 
rs.close


%>
<%if ShowNew then 
	Id=0
%>
	<tr> 
		<td>
			<% 	

			inQue = "(select IdDatoTecnico from AccountProdottoDatoTecn Where IdAccount=" & IdAccount & " and  IdProdotto= " & IdProdotto & " ) "
			
			IdRef="IdDatoTecnico" & Id 	
			query = ""
			query = query & " Select * from DatoTecnico " 
			query = query & " Where IdDatoTecnico not in " & inQue 
			query = query & " order By DescDatoTecnico"
			'response.write query
			err.clear
			response.write ListaDbChangeCompleta(Query,IdRef,"0","IdDatoTecnico","DescDatoTecnico",0,"","","","","dati assenti","class='form-control form-control-sm'")
			response.write err.description
			xx="0" & LeggiCampo(query,"IdDatoTecnico")
			%>
		</td>
		<td align="left">
			<%if Cdbl(xx)>0 then %>
			<%RiferimentoA="col-2;#;;2;insert;Inserisci;;SaveWithOper('INS')"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
			<%end if %>
		</td>
	</tr>			
	   
<%end if%>
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
