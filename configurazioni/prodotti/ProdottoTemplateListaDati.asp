<%
  NomePagina="ProdottoTemplateListaDati.asp"
  titolo="Template Prodotto : dati aggiuntivi"
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
<!-- Custom styles for this template -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">
</head>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->


<%
  NameLoaded= ""

  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdProdottoTemplate=0
  if FirstLoad then 
     PaginaReturn        = getCurrentValueFor("PaginaReturn")
     IdProdottoTemplate  = "0" & getCurrentValueFor("IdProdottoTemplate")
     OperTabella         = Session("swap_OperTabella")
  else
	 IdProdottoTemplate  = "0" & getValueOfDic(Pagedic,"IdProdottoTemplate")
	 OperTabella         = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn        = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdProdottoTemplate = cdbl(IdProdottoTemplate)
  if Cdbl(IdProdottoTemplate)=0 then 
     response.redirect RitornaA(PaginaReturn)
     response.end 
  end if 

  if Oper="CALL_INS" or Oper="CALL_UPD" then 
     xx=RemoveSwap()
     Session("TimeStamp")=TimePage
     KK=trim(Request("ItemToRemove"))
     if kk <> "" then 
        Session("swap_IdProdottoTemplate") = IdProdottoTemplate
        Session("swap_IdOpzione")          = KK
        Session("swap_OperTabella")        = Oper
        Session("swap_PaginaReturn")       = "configurazioni/prodotti/ProdottoTemplateListaDati.asp"
        response.redirect virtualPath & "configurazioni/prodotti/ProdottoTemplateDatoModifica.asp"
        response.end 
     end if 
  End if 

  if Oper="CALL_INS_TEC" or Oper="CALL_UPD_TEC" then 
     xx=RemoveSwap()
     Session("TimeStamp")=TimePage
     KK=trim(Request("ItemToRemove"))
     if kk <> "" then 
        Session("swap_IdProdottoTemplate") = IdProdottoTemplate
        Session("swap_IdDatoTecnico")      = KK
        Session("swap_OperTabella")        = Oper
        Session("swap_PaginaReturn")       = "configurazioni/prodotti/ProdottoTemplateListaDati.asp"
        response.redirect virtualPath & "configurazioni/prodotti/ProdottoTemplateDatoTecnicoModifica.asp"
        response.end 
     end if 
  End if 
  
  if Oper="CALL_DEL" and CheckTimePageLoad() then 
     KK=Request("ItemToRemove")
     if Cdbl("0" & KK ) > 0 then 
	    q = ""
		q = q & "delete from ProdottoTemplateOpzione"
        q = q & " Where IdProdottoTemplate = " & IdProdottoTemplate
		q = q & " and   IdOpzione = '" & KK & "'"
		'response.write q
        ConnMsde.execute q
     end if 
  
  end if  
  if Oper="CALL_DEL_TEC" and CheckTimePageLoad() then 
     KK=Request("ItemToRemove")
     if Cdbl("0" & KK ) > 0 then 
	    q = ""
		q = q & "delete from ProdottoTemplateDatoTecnico"
        q = q & " Where IdProdottoTemplate = " & IdProdottoTemplate
		q = q & " and   IdDatoTecnico = " & KK 
		'response.write q
        ConnMsde.execute q
     end if 
  
  end if    
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdProdottoTemplate"   ,IdProdottoTemplate)
  xx=setValueOfDic(Pagedic,"OperTabella"  ,OperTabella)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  qProd = "select * from ProdottoTemplate where idProdottoTemplate=" & IdProdottoTemplate
  IdAnagServizio        = LeggiCampo(qProd,"IdAnagServizio")
  DescProdottoTemplate  = LeggiCampo(qProd,"DescProdottoTemplate")
  IdAnagCaratteristica  = LeggiCampo(qProd,"IdAnagCaratteristica")
  IdAnagCaratteristicaKey=""
  
  if cdbl(IdAnagCaratteristica)>0 then 
     q = ""
     q = q & " select IdAnagCaratteristicaKey From AnagCaratteristica "
     q = q & " where IdAnagCaratteristica =" & IdAnagCaratteristica
     IdAnagCaratteristicaKey = LeggiCampo(q,"IdAnagCaratteristicaKey")
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
				<div class="col-11"><h3>Elenco Dati Aggiuntivi</b></h3>
				</div>
			</div>
	        <div class="row">
	           <div class="col-1">
	           </div>
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("Template Prodotto")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescProdottoTemplate%>" >
               </div>	
			</div>			

<%

Set Rs = Server.CreateObject("ADODB.Recordset")

MySql = "" 
MySql = MySql & " Select A.*  "
MySql = MySql & " From Opzione A "
MySql = MySql & " Where  A.IdAnagServizio = '" & IdAnagServizio & "'"
MySql = MySql & " and  A.IdTipoOpzione = 'TECN'"
MySql = MySql & " and a.IdAnagCaratteristicaKey = '|NESSUNO|'" 
MySql = MySql & " order by DescInterna"

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

If Err.number<>0 then	
    MsgNoData = Err.description
elseIf Rs.EOF then	
    MsgNoData = "Nessun dettaglio in archivio"
End if

%>
<div class="row">
     <div class="col-12 bg-primary text-white font-weight-bold">
     Dati Standard 
     </div>
</div>
			
<div class="table-responsive"><table class="table"><tbody>

<%
if MsgNoData<>"" or MsgErrore<>"" then %>
	<tr>
		<td colspan='6' align='center'>
		<div class="bg-danger text-white"><%=server.htmlencode(MsgErrore) & " " & server.htmlencode(MsgNoData) %></div>
		</td>
	</tr>	
<%end if 

if MsgNoData="" then 
	Do While Not rs.EOF 
		%>
		
		<tr scope="col">
			<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescInterna")%>">
			</td>
		</tr>
		<%	
		rs.MoveNext
	Loop
end if 
rs.close
%>
</tbody></table></div> <!-- table responsive fluid -->
<%
MySql = "" 
MySql = MySql & " Select A.*,isNull(B.IdOpzione,'') as OpzioneRef,B.FlagObbligatorio,isNull(B.Ordine,9999) as Ordine  "
MySql = MySql & " ,isNull(B.Rigo,9999) as Rigo"
MySql = MySql & " From Opzione A Left join ProdottoTemplateOpzione B "
MySql = MySql & " on  A.IdOpzione = B.IdOpzione"
MySql = MySql & " and B.IdProdottoTemplate=" & IdProdottoTemplate
MySql = MySql & " Where  A.IdAnagServizio = '" & IdAnagServizio & "'"
MySql = MySql & " and  A.IdTipoOpzione = 'TECN'"
if IdAnagCaratteristicaKey="" then 
   MySql = MySql & " and a.IdAnagCaratteristicaKey='" & IdAnagCaratteristicaKey & "'" 
else
   MySql = MySql & " and a.IdAnagCaratteristicaKey like '%|" & IdAnagCaratteristicaKey & "|%'" 
end if 
MySql = MySql & " order By Rigo, Ordine ,DescInterna"

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

If Err.number<>0 then	
	MsgNoData = Err.description
elseIf Rs.EOF then	
	MsgNoData = "Nessun dettaglio in archivio"
End if


DescLoaded=""
NumCols = numC + 1
NumRec  = 0
ShowNew    = true
ShowUpdate = false
%>

<div class="row">
     <div class="col-12 bg-primary text-white font-weight-bold">
     Dati Standard Aggiuntivi 
     </div>
</div>

<div class="table-responsive"><table class="table"><tbody>
<thead>
	<tr>
	    <th scope="col">Sel.</th>
		<th scope="col">Dato Standard</th>
		<th scope="col" width="10%">Obbligatorio</th>
		<th scope="col" width="10%">Rigo</th>
		<th scope="col" width="10%">Ordine</th>
		<th scope="col" width="10%">Azioni</th>
	</tr>
</thead>

<%
if MsgNoData<>"" or MsgErrore<>"" then %>
	<tr>
		<td colspan='6' align='center'>
		<div class="bg-danger text-white"><%=server.htmlencode(MsgErrore) & " " & server.htmlencode(MsgNoData) %></div>
		</td>
	</tr>	
<% end if 




'non deve paginare
PageSize = 0
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
		Id=Rs("IdOpzione")
		err.clear 
		checked=""
		DescObbl="n.d."
		Ordine="n.d."
		if rs("OpzioneRef")<>"" then
		   checked = " checked "
		   if Rs("FlagObbligatorio")=1 then 
		      DescObbl="SI"
		   else
		      DescObbl="NO"
		   end if
		   Ordine=Rs("Ordine")
		   rigo  =Rs("rigo")
		end if 
		
		%>
		
		<tr scope="col">
		    <td>
			   <input type="checkbox" <%=checked%> class="big-checkbox" disabled>
			</td>
			<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescInterna")%>">
			</td>
			<td><input class="form-control" type="text" readonly value="<%=descObbl%>"></td>
			<td><input class="form-control" type="text" readonly value="<%=rigo%>"></td>
			<td><input class="form-control" type="text" readonly value="<%=ordine%>"></td>
			
            <td>
				<%if rs("OpzioneRef")="" then %>
  				    <%RiferimentoA="col-2;#;;2;plus;Inserisci;;AttivaFunzione('CALL_INS','" & Id & "');N"%>
				    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->							
				<%else %>
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

<div class="row">
     <div class="col-12 bg-primary text-white font-weight-bold">
     Dati Tecnici 
     </div>
</div>
<div class="table-responsive"><table class="table"><tbody>
	<tr>
	    <th scope="col">Sel.</th>
		<th scope="col">Dato Tecnico
		<a href="#" title="Inserisci" onclick="AttivaFunzione('CALL_INS_TEC','0');">
        <i class="fa fa-2x fa-plus-square"></i></a>
		</th>
		<th scope="col" width="10%">Obbligatorio</th>
		<th scope="col" width="10%">Rigo</th>
		<th scope="col" width="10%">Ordine</th>
		<th scope="col" width="10%">Azioni</th>
	</tr>

<%
MySql = "" 
MySql = MySql & " Select A.*,B.FlagObbligatorio,isNull(B.Ordine,9999) as Ordine  "
MySql = MySql & " ,isNull(B.Rigo,9999) as Rigo"
MySql = MySql & " From DatoTecnico A inner join ProdottoTemplateDatoTecnico B "
MySql = MySql & " on  A.IdDatoTecnico = B.IdDatoTecnico"
MySql = MySql & " and B.IdProdottoTemplate=" & IdProdottoTemplate
MySql = MySql & " order By Rigo, Ordine ,DescDatoTecnico"

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
%>

<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->
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
		Id=Rs("IdDatoTecnico")
		err.clear 
		
		DescObbl="n.d."
		Ordine="n.d."
	    if Rs("FlagObbligatorio")=1 then 
	       DescObbl="SI"
	    else
	       DescObbl="NO"
	    end if
	    Ordine=Rs("Ordine")
	    rigo  =Rs("rigo")
		
		%>
		
		<tr scope="col">
		    <td>
			   <input type="checkbox" 'checked' class="big-checkbox" disabled>
			</td>
			<td>
			<input class="form-control" type="text" readonly value="<%=Rs("DescDatoTecnico")%>">
			</td>
			<td><input class="form-control" type="text" readonly value="<%=descObbl%>"></td>
			<td><input class="form-control" type="text" readonly value="<%=rigo%>"></td>
			<td><input class="form-control" type="text" readonly value="<%=ordine%>"></td>
			
            <td>
			    <%RiferimentoA="col-2;#;;2;upda;Aggiorna;;AttivaFunzione('CALL_UPD_TEC','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
			    <%RiferimentoA="col-2;#;;2;dele;Cancella;;AttivaFunzione('CALL_DEL_TEC','" & Id & "');N"%>
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
