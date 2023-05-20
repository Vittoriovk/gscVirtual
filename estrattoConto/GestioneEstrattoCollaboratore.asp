<%
  NomePagina="GestioneEstrattoCollaboratore.asp"
  titolo="Menu - Gestione Estratto Conto"
  default_check_profile="Coll"
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
<script>
function localGes(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
	xx=ImpostaValoreDi("Oper","CALL_GES");
	document.Fdati.submit();
}
function localCauzione()
{
	xx=ImpostaValoreDi("Oper","CALL_NEW");
	document.Fdati.submit();
}

localCauzione
</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%
   PaginaReturn = ""
   DenomCliente = ""
  
   if FirstLoad then 
      PaginaReturn     = getCurrentValueFor("PaginaReturn")
   else
      PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
   end if 
   
   Set Rs = Server.CreateObject("ADODB.Recordset")
   xx=setValueOfDic(Pagedic,"PaginaReturn"      ,PaginaReturn)
   xx=setCurrent(NomePagina,livelloPagina) 

   if Oper=ucase("Removeitem") then 
      IdEstratto = Cdbl("0" & Request("ItemToRemove"))
	  if Cdbl(IdEstratto)>0 then   
	     ConnMsde.execute "Delete From EstrattoConto where IdEstrattoConto = " & IdEstratto
		 ConnMsde.execute "update AccountMovEco set  IdEstrattoConto=0 where IdEstrattoConto = " & IdEstratto
	  end if 
   end if 
   if Oper="CALL_GES" then 
      IdFormazione = Cdbl("0" & Request("ItemToRemove"))
	  if Cdbl(IdFormazione)>0 then   
         xx=RemoveSwap()
		 Session("swap_IdFormazione")  = IdFormazione
         Session("swap_PaginaReturn")  = "Formazione/" & NomePagina
		 if IsBackOffice() then 
		    response.redirect RitornaA("Formazione/ModificaRichiestaFormazioneBackO.asp")
		 else 
            response.redirect RitornaA("Formazione/ModificaRichiestaFormazione.asp")
		 end if 
         response.end 
      end if 
   end if    

   InfoPage="Storico Formazione"
 
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
 
			<div class="row">
				<%
				if PaginaReturn<>"" then 
				   RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"
				%>
				   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<% end if %>
				<div class="col-11"><h3><%=InfoPage%></h3>
				</div>
			</div>
			
			 <!--#include file="FiltroSearchTipo.asp"-->
			 
    
<%

   if condizione<>"" then 
      condizione = " and " & condizione
   end if 
'caricamento tabella 
   NoEnded = true 
   MySql = "" 
   
   MySql = MySql & " select a.*,B.DescStatoCredito,D.Denominazione as DescCliente"
   MySql = MySql & " from EstrattoConto A,StatoCredito B , Cliente D"
   MySql = MySql & " Where A.IdAccount = D.IdAccount" 
   MySql = MySql & " and   A.IdAccountGestore=" & Session("LoginIdAccount")
   MySql = MySql & " and   A.IdStatoEstratto = B.IdStatoCredito"
   MySql = MySql & " and   B.FlagStatoFinale = 0"
   MySql = MySql & " order by DataEstratto Desc "   
   
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

<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

	
		<div class="table-responsive"><table class="table"><tbody>
		<thead>
		<tr><th scope="col">Descrizione </th>
	     <th scope="col">Cliente</th>
		<th scope="col" width="15%">Data Estratto</th>
		<th scope="col" width="15%">Importo &euro;</th>
		    <th scope="col" width="15%">Stato</th>
		    <th scope="col" width="8%">Azioni</th>
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
				
                Id=Rs("IdEstrattoConto")
				DescLoaded=DescLoaded & Id & ";"
				
		%>
			<tr scope="col">
  			    <td>
                   <input class="form-control" type="text" readonly value="<%=Rs("DescEstratto")%>">
                 </td>
			        <td scope="col">
					    <input class="form-control" type="text" readonly value="<%=RS("DescCliente")%>">
					</td>
				<td>
					<input class="form-control" type="text" readonly value="<%=StoD(RS("DataEstratto"))%>">
				</td>
				<td>
					<input class="form-control" type="text" readonly value="<%=insertPoint(RS("ImptEstratto"),2)%>">
				</td>
                 <td>
                   <input class="form-control" type="text" readonly value="<%=rs("DescStatoCredito")%>">
                 </td>

			     <td>
                      <a href='<%=virtualPath%>/pdf/EstrattoConto.asp?IdEstrattoConto=<%=id%>' title="Stampa" target="_new">
	                  <i class="fa fa-2x fa-file-pdf-o"></i></a>  
				 
                     <%RiferimentoA="col-2;#!;;2;dele;Cancella;;RemoveItem('" & Id & "');N" %>
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
