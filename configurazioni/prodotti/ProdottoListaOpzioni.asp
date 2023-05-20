<%
  NomePagina="ProdottoListaOpzioni.asp"
  titolo="Prodotto : Opzioni aggiuntive"
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
  IdProdotto=0
  if FirstLoad then 
     PaginaReturn     = getCurrentValueFor("PaginaReturn")
     IdProdotto       = "0" & getCurrentValueFor("IdProdotto")
     OperTabella      = Session("swap_OperTabella")
  else
	 IdProdotto       = "0" & getValueOfDic(Pagedic,"IdProdotto")
	 OperTabella      = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdProdotto = cdbl(IdProdotto)
  if Cdbl(IdProdotto)=0 then 
     response.redirect RitornaA(PaginaReturn)
     response.end 
  end if 

  if Oper="CALL_INS" or Oper="CALL_UPD" then 
     xx=RemoveSwap()
     Session("TimeStamp")=TimePage
     KK=trim(Request("ItemToRemove"))
     if kk <> "" then 
        Session("swap_IdProdotto")    = IdProdotto
        Session("swap_IdOpzione")     = KK
        Session("swap_OperTabella")   = Oper
        Session("swap_PaginaReturn")  = "configurazioni/prodotti/ProdottoListaOpzioni.asp"
        response.redirect virtualPath & "configurazioni/prodotti/ProdottoOpzioneModifica.asp"
        response.end 
     end if 
  End if 

  if Oper="CALL_DEL" and CheckTimePageLoad() then 
     KK=Request("ItemToRemove")
     if Cdbl("0" & KK ) > 0 then 
	    q = ""
		q = q & "delete from ProdottoOpzione"
        q = q & " Where IdProdotto = " & IdProdotto
		q = q & " and   IdOpzione = '" & KK & "'"
		'response.write q
        ConnMsde.execute q
     end if 
  
  end if  
  
  'registro i Opzioni della pagina 
  xx=setValueOfDic(Pagedic,"IdProdotto"   ,IdProdotto)
  xx=setValueOfDic(Pagedic,"OperTabella"  ,OperTabella)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  IdAnagServizio = LeggiCampo("select * from Prodotto where idProdotto=" & IdProdotto,"IdAnagServizio")
  DescProdotto   = LeggiCampo("select * from Prodotto where idProdotto=" & IdProdotto,"DescProdotto")
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
				<div class="col-11"><h3>Elenco Opzioni Aggiuntive</b></h3>
				</div>
			</div>
	        <div class="row">
	           <div class="col-1">
	           </div>
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("Prodotto")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescProdotto%>" >
               </div>	
			</div>			

<%
	
Set Rs = Server.CreateObject("ADODB.Recordset")

MySql = "" 
MySql = MySql & " Select A.*,isNull(B.IdOpzione,'') as OpzioneRef,B.FlagObbligatorio,isNull(B.Ordine,9999) as Ordine  "
MySql = MySql & " From Opzione A Left join ProdottoOpzione B "
MySql = MySql & " on  A.IdOpzione = B.IdOpzione"
MySql = MySql & " and B.IdProdotto=" & IdProdotto
MySql = MySql & " Where  A.IdAnagServizio = '" & IdAnagServizio & "'"
MySql = MySql & " and  A.IdTipoOpzione = 'OPZI'"
MySql = MySql & " order By Ordine ,DescInterna"

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
	    <th scope="col">Sel.</th>
		<th scope="col">Opzione Aggiuntiva</th>
		<th scope="col" width="10%">Obbligatorio</th>
		<th scope="col" width="10%">Ordine</th>
		<th scope="col" width="10%">Azioni</th>
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
