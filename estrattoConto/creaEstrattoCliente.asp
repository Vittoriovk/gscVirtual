<%
  NomePagina="CreaEstrattoCliente.asp"
  titolo="Menu - Gestione Estratto Conto Cliente"
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
	xx=ImpostaValoreDi("Oper","GENERA");
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
	  DataRiferimento  = 0
   else
      PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
      DataRiferimento  = getValueOfDic(Pagedic,"DataRiferimento")
   end if 
   
   Set Rs = Server.CreateObject("ADODB.Recordset")

   xx=setValueOfDic(Pagedic,"PaginaReturn"      ,PaginaReturn)
   xx=setValueOfDic(Pagedic,"DataRiferimento"   ,DataRiferimento)
   xx=setValueOfDic(Pagedic,"DenomCliente"      ,DenomCliente)
   xx=setCurrent(NomePagina,livelloPagina) 

   if Oper="GENERA" then 
      IdAccount = Cdbl("0" & Request("ItemToRemove"))
	  if Cdbl(IdAccount)>0 then   
	     IdTipoCredito = Request("ListaModPag")
		 DataRiferimento=Request("DataRiferimento0")
		 DataRiferimento=DataStringa(Request("DataRiferimento0"))
		 DataEstratto = Dtos()
		 qUpd = "elaboraEstrattoConto " & IdAccount & "," & session("LoginIdAccount") & ",'" & IdTipoCredito & "'," & DataRiferimento  & "," & DataEstratto
		 'response.write qUpd 
		 connMsde.execute qUpd
		 if Err.Number <>0 then 
		    MsgErrore= ErroreDb(err.description)
		 end if 

      end if 
   end if    

   InfoPage="Gestione Estratto Conto"
 
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
		   
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Estratti Al </p>
      </div> 
	  <div class="col-2">
	  	  <%
		  DataRiferimento=Request("DataRiferimento0")
		  DataRiferimento=DataStringa(Request("DataRiferimento0"))
		  	
		  if Cdbl("0" & DataRiferimento) > 20200101 then 
		     DataRiferimento = Cdbl("0" & DataRiferimento)
		  else
		     DataRiferimento = 0
		  end if 
		  DataRicerca=DataRiferimento
          NameLoaded= NameLoaded & ";DataRiferimento,TE"   		  
		  nome="DataRiferimento0"
	      if cdbl("0" & DataRiferimento) = 0 then 
		        DataRiferimento = Dtos()
		  end if 
		  
		  valo=StoD(DataRiferimento)
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
	  </div>

      <div class="col-2">
         <p class="font-weight-bold"> </p>
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
   
   if DataRicerca>0 then
      'cerco tutti i clienti con data da gestire 
      MySql = MySql & " select D.IdAccount,D.Denominazione as DescCliente,sum(ImptMovEco*SegnoSistema)*-1 as Totale "
      MySql = MySql & " from AccountMovEco A , Cliente D"
      MySql = MySql & " Where A.IdAccount = D.IdAccount"  
	  MySql = MySql & " and   A.IdEstrattoConto = 0 "  
	  MySql = MySql & " and   A.DataMovEco <= " & DataRicerca
	  MySql = MySql & " and   A.IdTipoCredito='" & IdtipoCredito & "'"
	  MySql = MySql & " and   A.IdStatoCredito in ('ATTI','UTIL')"
      IdAccColl = Session("LoginIdAccount")
	  filtro = trim(getCondForLevel(session("LivelloAccount"),Session("LoginIdAccount")))
	  if filtro<>"" then 
         MySql = MySql & " and D." & Filtro
      else
         MySql = MySql & " and D.IdAccountLivello1=-1"
      end if 
	  MySql = MySql & Condizione 
      MySql = MySql & " group by D.IdAccount,D.Denominazione"
      MySql = MySql & " order by D.Denominazione " 
  
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
		<tr><th scope="col">Cliente </th>
		<th scope="col" width="15%">Importo Estratto</th>
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
				
                Id=Rs("IdAccount")
				DescLoaded=DescLoaded & Id & ";"
				
				DescStato = Rs("DescStatoServizio") & " " & Rs("NoteFormazioneCliente")
		%>
			<tr scope="col">
  			    <td>
                   <input class="form-control" type="text" readonly value="<%=Rs("DescCliente")%>">
                 </td>
				<td>
					<input class="form-control" type="text" readonly value="<%=InsertPoint(RS("Totale"),2)%>">
				</td>

			     <td>
				    <% if isCliente() or true then %>
                          <%RiferimentoA="col-2;#;;2;crea;Genera;;localGes('" & Id & "');N"
					      %>
					      <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
						  
					<% end if %>

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
  'chiude DataRicerca=0 
  end if %>
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
