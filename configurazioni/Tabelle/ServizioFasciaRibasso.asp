<%
 NomePagina="ServizioFasciaRibasso.asp"
 titolo="Fascia Ribasso"
 
 'forzo il controllo al profilo 
 default_check_profile="SUPERV"
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
NameLoaded=NameLoaded & "IdFascia,FLP;Percentuale,FLQ"
Set Rs = Server.CreateObject("ADODB.Recordset")

DescServizio   = ""
IdAnagServizio = ""
if FirstLoad then 
   IdAnagServizio = getCurrentValueFor("IdAnagServizio")
   DescServizio   = getCurrentValueFor("DescServizio")
   PaginaReturn   = getCurrentValueFor("PaginaReturn") 
else
   IdAnagServizio = getValueOfDic(Pagedic,"IdAnagServizio")
   DescServizio   = getValueOfDic(Pagedic,"DescServizio")
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
end if 

IdAnagServizio = trim(IdAnagServizio)
if IdAnagServizio=""  then 
   IdAnagServizio="CAUZ_DEFI"
end if 
if DescServizio="" then 
   DescServizio = LeggiCampo("Select * from AnagServizio where IdAnagServizio='" & IdAnagServizio & "'","DescAnagServizio") 
end if 

on error resume next

flagRecalc = false 
if Oper="INS" then 
   Session("TimeStamp")=TimePage
   KK=0
   IdFascia     = request("IdFascia"    & kk)
   Percentuale  = request("Percentuale" & kk)
   IdFasciaPrec = 0
   DescFascia = "Ribasso dal " & IdFasciaPrec & "%  al " & IdFascia & "%"
   MyQ = ""
   MyQ = MyQ & " Insert into ServizioFascia (IdAnagServizio,IdFascia,DescFascia,Percentuale,IdFasciaPrec) "
   MyQ = MyQ & " values ("
   MyQ = MyQ & " '" & IdAnagServizio & "'"
   MyQ = MyQ & ", " & NumForDb(IdFascia)
   MyQ = MyQ & ",'" & DescFascia & "'"   
   MyQ = MyQ & ", " & NumForDb(Percentuale)
   MyQ = MyQ & ", " & NumForDb(IdFasciaPrec)
   MyQ = MyQ & " )"
   
   ConnMsde.execute MyQ
   if err.number=0 then 
      flagRecalc = true 
      IdFascia   =0
      Percentuale=0
   else
      MsgErrore = ErroreDb(err.description) 
   end if 

end if 

if Oper="MOD" then 
   Session("TimeStamp")=TimePage
   KK=Request("ItemToRemove")
   IdRow      = KK
   IdFascia   = request("IdFascia"    & kk)
   Percentuale= request("Percentuale" & kk)
   
   MyQ = "" 
   MyQ = MyQ & " update ServizioFascia set "
   MyQ = MyQ & " IdFascia = "    & NumForDb(IdFascia)
   MyQ = MyQ & ",Percentuale = " & NumForDb(Percentuale)
   MyQ = MyQ & " where IdRow = " & Idrow
   ConnMsde.execute MyQ
   if err.number<>0 then 
      MsgErrore = ErroreDb(err.description) 
   else 
      flagRecalc = true 
   end if    
   IdFascia   =0
   Percentuale=0
End if 

if Oper="DEL" then 
   Session("TimeStamp")=TimePage
   KK=Request("ItemToRemove")
   IdRow      = KK
   
   MyQ = "" 
   MyQ = MyQ & " delete from ServizioFascia "
   MyQ = MyQ & " where IdRow = " & Idrow
   ConnMsde.execute MyQ
   flagRecalc = true 
   IdFascia   =0
   Percentuale=0
 
End if 
'ricalcolo le descrizione delle fascie 
if flagRecalc = true then 
   MySql = "" 
   MySql = MySql & " Select * "
   MySql = MySql & " from ServizioFascia "
   MySql = MySql & " where IdAnagServizio = '" & IdAnagServizio & "'"
   MySql = MySql & " order by IdFascia"

   Rs.CursorLocation = 3 
   Rs.Open MySql, ConnMsde
   Do While Not rs.EOF 
      qPrec = ""
      qPrec = qPrec & " select max(IdFascia) as idM from ServizioFascia "
      qPrec = qPrec & " Where IdAnagServizio='" & IdAnagServizio & "'"
      qPrec = qPrec & " and IdFascia< " & NumForDb(Rs("IdFascia"))
      IdFasciaPrec = cdbl("0" & LeggiCampo(qPrec,"idM"))
      DescFascia = "Ribasso dal " & IdFasciaPrec & "%  al " & Rs("IdFascia") & "%"
      MyQ = "" 
      MyQ = MyQ & " update ServizioFascia set "
      MyQ = MyQ & " IdFasciaPrec = " & NumForDb(IdFasciaPrec)
      MyQ = MyQ & ",DescFascia = '"  & DescFascia & "'"
      MyQ = MyQ & " where IdRow = " & Rs("IdRow")
      ConnMsde.execute MyQ  

      rs.MoveNext
   Loop
   rs.close
   
end if 


  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdAnagServizio" ,IdAnagServizio)
  xx=setValueOfDic(Pagedic,"DescServizio"   ,DescServizio)
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
				<div class="col-11"><h4>Fasce di Ribasso : <%=DescServizio%></h4>
				</div>
			</div>
<%

MySql = "" 
MySql = MySql & " Select * "
MySql = MySql & " from ServizioFascia "
MySql = MySql & " where IdAnagServizio = '" & IdAnagServizio & "'"
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
	    <th scope="col">Descrizione</th>
		<th scope="col">Percentuale di ribasso fino al</th>
		<th scope="col">Percentuale calcolo somma garantita</th>
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
	Do While Not rs.EOF 
		Primo=Primo+1
		NumRec=NumRec+1
		Id=Rs("IdRow")
		DescLoaded=DescLoaded & Id & ";"

		%> 

	<tr scope="col"> 
		<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescFascia")%>">
		</td>	
		<td>
			<input class="form-control" Id="IdFascia<%=Id%>"    Name="IdFascia<%=Id%>"   type="text" value="<%=Rs("IdFascia")%>">
		</td>
		<td>
			<input class="form-control" Id="Percentuale<%=Id%>" Name="Percentuale<%=Id%>"  type="text" value="<%=Rs("Percentuale")%>">
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
			<input type="text" class="form-control" readonly value="n.d.">
		</td>	
		<td>
			<input type="text" class="form-control" Id="IdFascia<%=Id%>"    Name="IdFascia<%=Id%>"      value="<%=IdFascia%>">
		</td>
		<td>
			<input type="text" class="form-control" Id="Percentuale<%=Id%>" Name="Percentuale<%=Id%>"   value="<%=Percentuale%>">
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
