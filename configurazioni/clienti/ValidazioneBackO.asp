<%
  NomePagina="ValidazioneBackO.asp"
  titolo="Menu - Gestione validazione Cliente"
  default_check_profile="COLL,BackO"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAffidamento.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionEvento.asp"-->
<%
  livelloPagina="00"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->
<%
'xx=DumpDic(SessionDic,NomePagina)
%>

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include virtual="/gscVirtual/include/head.asp"-->
<!-- Custom styles for this template -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">

</head>
<script>
function localGes(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
	xx=ImpostaValoreDi("Oper","CALL_GES");
	document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%
PaginaReturn = ""
 
   if FirstLoad then 
      PaginaReturn     = getCurrentValueFor("PaginaReturn")
	  TipoRife         = getCurrentValueFor("TipoRife")
	  IdAccountAti     = getCurrentValueFor("IdAccountAti")
	  TipoRichiesta    = getCurrentValueFor("TipoRichiesta")
   else
      PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
	  TipoRife         = getValueOfDic(Pagedic,"TipoRife")
	  IdAccountAti     = Request("IdAccountAti")
	  TipoRichiesta    = getValueOfDic(Pagedic,"TipoRichiesta")
   end if 
   
   IdAccountAti = TestNumeroPos("0" & IdAccountAti)

   'per la gestione di compagnia e collaboratore
   if FirstLoad then 
	  IdAccountCollaboratore = cdbl("0" & getValueOfDic(Pagedic,"IdAccountCollaboratore"))
   'dall'esterno arrivano filtri di ricerca
      TipoRicercaExt   = Session("swap_TipoRicercaExt")
      testo_ricercaExt = Session("swap_testo_ricercaExt")
      Session("swap_TipoRicercaExt") = ""
      Session("swap_testo_ricercaExt") = ""
      if TipoRicercaExt<>"" then 
         v_TipoRicerca = TipoRicercaExt
      end if 
      if testo_ricercaExt<>"" then 
         v_cercatesto = testo_ricercaExt
      end if 	  
   else
	  IdAccountCollaboratore= Cdbl("0" & Request("IdCollaboratore0"))
   end if 

   xx=setValueOfDic(Pagedic,"PaginaReturn"          ,PaginaReturn)
   xx=setValueOfDic(Pagedic,"TipoRife"              ,TipoRife)
   xx=setValueOfDic(Pagedic,"TipoRichiesta"         ,TipoRichiesta)
   xx=setValueOfDic(Pagedic,"IdAccountCollaboratore",IdAccountCollaboratore)
   xx=setCurrent(NomePagina,livelloPagina) 

   tabella=""
   descAzione=""
   if TipoRife="COOB" then 
      tabella="AccountCoobbligato"
	  descAzione="Coobbligato"
   elseif TipoRife="ATI" then 
      tabella="AccountATI"
	  descAzione="A.T.I."
   end if 
   if tabella="" then 
      response.redirect RitornaA(paginaReturn)
   end if 
   'presa in carico 
   
   if Oper="CALL_CAR" and TipoRichiesta<>"STORICO" then 
      idRichiesta  = "0" & request("ItemToRemove")
      if Cdbl(idRichiesta)>0 then 
	     upd=""
		 upd=upd & " Update " & tabella 
		 upd=upd & " set IdAccountBackOffice=" & Session("LoginIdAccount") 
		 upd=upd & "    ,IdStatoValidazione='LAVO'"
		 upd=upd & " Where IdAccountBackOffice=0 and Id" & tabella &"=" & idRichiesta
		 'response.write upd
		 ConnMsde.execute upd 
      end if 
   end if    
   
   'chiamo gestione 
   if Oper="CALL_GES" then 
      idRichiesta  = "0" & request("ItemToRemove")
      if Cdbl(idRichiesta)>0 then 
         xx=RemoveSwap()
         Session("swap_TipoRife")           = TipoRife
         Session("swap_IdRife")             = idRichiesta
         Session("swap_PaginaReturn")       = "configurazioni/Clienti/" & NomePagina
         response.redirect RitornaA("configurazioni/Clienti/ValidazioneBackODettaglio.asp")
         response.end 
      end if 
   end if 
   
   Oggi = Dtos() 
   Set Rs = Server.CreateObject("ADODB.Recordset")
   Set Ds = Server.CreateObject("ADODB.Recordset")

%>

<div class="d-flex" id="wrapper">
	<%
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

            <input type="hidden" name="IdAccountAti" id="IdAccountAti" value="<%=IdAccountAti%>">

			<div class="row">
				<%
				if TipoRichiesta<>"STORICO" then 
				   descEvento = "Gestione Validazione " & descAzione
				else 
				   descEvento = "Storico Validazione " & descAzione
				end if 
				
				if PaginaReturn<>"" then 
				   RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"
				%>
				   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<% end if %>
				<div class="col-11"><h3><%=descEvento%></h3>
				</div>
			</div>
			<%cercaComp="N"%>
			<!--#include virtual="/gscVirtual/configurazioni/Clienti/ricercaBackO.asp"-->
			

    
<%

'caricamento tabella 
   if Condizione<>"" then 
      Condizione = " and " & Condizione

   end if 


   MySql = "" 
   MySql = MySql & " select a.*,B.Denominazione,B.Codicefiscale,B.PartitaIva,C.DescStatoServizio" 
   MySql = MySql & " from " & Tabella & " A, Cliente B, StatoServizio C "
   MySql = MySql & " Where A.IdStatoValidazione = C.IdStatoServizio "
   if cdbl(IdAccountAti)>0 then 
      MySql = MySql & " and   A.IdAccountAti=" & IdAccountAti
   end if 
   if TipoRichiesta<>"STORICO" then 
      MySql = MySql & " and   C.FlagStatoFinale = 0"
   else 
      MySql = MySql & " and   C.FlagStatoFinale = 1"
   end if 
   MySql = MySql & " and   A.IdStatoValidazione <>''"
   MySql = MySql & " and   A.IdAccount = B.IdAccount" 
   tipoGestore=ucase(Session("LoginTipoUtente"))
   MySql = MySql & " and   A.TipoGestore like '%" & tipoGestore & "%'"
   MySql = MySql & Condizione
   if Cdbl(IdAccountCollaboratore)>0 then 
      MySql = MySql & " and   B.IdAccountLivello1 = " & IdAccountCollaboratore
   end if 
   'response.write MySql
   MySql = MySql & " order by DataRichiesta "   

   Rs.CursorLocation = 3 
   Rs.Open MySql, ConnMsde

   DescLoaded=""
   NumCols = numC + 1
   NumRec  = 0
   ShowNew    = true
   ShowUpdate = false
   MsgNoData  = ""
   
  
%>

<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

	
		<div class="table-responsive"><table class="table"><tbody>
		<thead>
		<tr>
			<th scope="col" width="12%">Richiesta Del</th>
			<th scope="col">Cliente</th>
			<th scope="col"><%=descAzione%></th>
		    <th scope="col">Stato</th>
			<th scope="col">Gestita da</th>
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
			Primo=0
			Do While Not rs.EOF and (NumRec<PageSize or Pagesize<=0)
				Primo=Primo+1
				NumRec=NumRec+1

                Id=Rs("Id" & Tabella)
				IdAccBackO=Rs("IdAccountBackOffice")
				
				IdStato=Rs("IdStatoValidazione")
				StatoComp=funDoc_StatoDocum(Id)
				DescLoaded=DescLoaded & Id & ";"
		%>
			<tr scope="col">
				<td>
					<input class="form-control" type="text" readonly value="<%=StoD(Rs("DataRichiesta"))%>">
				</td>
                 <td>
                   <input class="form-control" type="text" readonly value="<%=Rs("Denominazione")%>">
                 </td>				
                 <td>
                   <input class="form-control" type="text" readonly value="<%=Rs("RagSoc")%>">
                 </td>
                 <td>
                   <input class="form-control" type="text" readonly value="<%=Rs("DescStatoServizio")%>">
                 </td>
                 <td>
				   <%
				   DescBackOffice="da assegnare"
				   If Cdbl(IdAccBackO)>0 then 
				      DescBackOffice=LeggiCampo("select * from Account Where IdAccount=" & IdAccBackO,"Nominativo")
				   end if 
				   
				   %>
                   <input class="form-control" type="text" readonly value="<%=DescBackOffice%>">
                 </td>				 

			     <td>
				    <%if (IsBackOffice() or IsCollaboratore()) and Cdbl(IdAccBackO)=0 then
					     RiferimentoA=";#;;2;sele;Prendi in carico;;AttivaFunzione('CALL_CAR','" & Id & "');N"
					%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				    <%end if %>
					
					<%RiferimentoA=";#;;2;dett;Dettaglio;;localGes('" & Id & "');N"
					%>
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
			<!--#include virtual="/gscVirtual/include/paginazione.asp"-->

			</form>
		</div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->

<!--#include virtual="/gscVirtual/include/scripts.asp"-->
<%
if FirstLoad then 
	response.write "<script language=javascript>document.Fdati.submit();</script>" 
	response.end 
end if
%>
  <!-- Menu Toggle Script -->
  <script>
    $("#menu-toggle").click(function(e) {
      e.preventDefault();
      $("#wrapper").toggleClass("toggled");
    });
  </script>
  <script>
    $(document).ready(function(){
      $('[data-toggle="tooltip"]').tooltip();   
    });
  </script>
  <script>
$('.btn').onClick(function(e){
  e.preventDefault();
});  
</script>
</body>

</html>
