<%
  NomePagina="CercaCliente.asp"
  titolo="Ricerca Clienti"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
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

function setClie(id)
{
	
}
function callFun(id)
{
    xx=ImpostaValoreDi("ItemToModify",id);
	xx=ImpostaValoreDi("Oper","ESEGUI");
	document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%

if FirstLoad then 
   CurrentFun     = getCurrentValueFor("CurrentFun")
   OperAmmesse    = getCurrentValueFor("OperAmmesse")   
   PaginaReturn   = getCurrentValueFor("PaginaReturn")
   PageToCall     = getCurrentValueFor("PageToCall")
   opzioneSidebar = getCurrentValueFor("opzioneSidebar")

else
   CurrentFun     = getValueOfDic(Pagedic,"CurrentFun")
   OperAmmesse    = getValueOfDic(Pagedic,"OperAmmesse")
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
   PageToCall     = getValueOfDic(Pagedic,"PageToCall")
   opzioneSidebar = getValueOfDic(Pagedic,"opzioneSidebar")
end if 

'registro i dati della pagina 
xx=setValueOfDic(Pagedic,"CurrentFun"    ,CurrentFun)
xx=setValueOfDic(Pagedic,"OperAmmesse"   ,OperAmmesse)
xx=setValueOfDic(Pagedic,"PaginaReturn"  ,PaginaReturn)
xx=setValueOfDic(Pagedic,"PageToCall"    ,PageToCall)
xx=setValueOfDic(Pagedic,"opzioneSidebar",opzioneSidebar)
xx=setCurrent(NomePagina,livelloPagina) 

if Oper="ESEGUI" then
   xx=RemoveSwap()
   itemId = Cdbl("0" & Request("ItemToModify"))
   Session("swap_IdCliente")         = itemId
   IdAccountCliente = LeggiCampo("select * from Cliente Where idCliente=" & itemId,"IdAccount")
   Session("swap_IdAccountCliente")  = IdAccountCliente
   Session("swap_PaginaReturn")   = "utility/" & NomePagina
   Session("swap_OperAmmesse")    = OperAmmesse
   Session("swap_opzioneSidebar") = opzioneSidebar
   'response.write Session("swap_IdCliente") & " " & Session("swap_IdAccountCliente")
   response.redirect PageToCall
   response.end 
   
end if  

ItemToModify = "0"

%>

<div class="d-flex" id="wrapper">
	<%
	  Session("opzioneSidebar")=opzioneSidebar
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
				<div class="col-11"><h3>Ricerca Cliente per : <%=CurrentFun%></h3>
				</div>
			</div>
			<!-- Mostro i collaboratori diretti se non sono un segnalatore -->
			
			<%if isBackOffice() or (isCollaboratore() and ucase(Session("LoginTipoCollaboratore"))<>"SEGN") then
			%>
			
				<div class="row">
                   <div class="col-1 s1 no-margin font-weight-bold">Collaboratore</div>
	               <div class="col-9 no-margin">
	               <%
	                    stdClass="class='form-control form-control-sm'"
	                    IdAccountCollaboratore=Cdbl("0" & Request("IdCollaboratore0"))
						'il back office prende i livelli 1 
						if IsBackOffice() then 
						   CondRef = "livello = 1"
						else 
						   CondRef = getCondForLevel(session("LivelloAccount"),Session("LoginIdAccount"))
                        end if    
	                    q = "Select * from Collaboratore Where " & CondRef 
		                q = Q & " order by Denominazione "
		                'Where 
	                    response.write ListaDbChangeCompleta(q,"IdCollaboratore0",IdAccountCollaboratore ,"IdAccount","Denominazione" ,1,"","","","","",stdClass)
	  
	              %>
	              </div>
                  <div class="col-2 no-margin"></div>	
                </div>
            <%end if %>
			<%
			AddRow=true
			dim CampoDb(10)
			ElencoOption = ";0;Nominativo;1;Codice Fiscale;2;Partita Iva;2;"
            CampoDB(1)   = "a.Denominazione"
			CampoDb(2)   = "a.CodiceFiscale"
			CampoDb(3)   = "a.PartitaIva"
			
			%>
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
		
<%
			'caricamento tabella 
			if Condizione<>"" then 
				Condizione = " And " & Condizione
			end if 

			Set Rs = Server.CreateObject("ADODB.Recordset")
			MySql = "" 
			MySql = MySql & " Select A.*,c.FlagAttivo,c.DescBlocco From Cliente A "
			MySql = MySql & " left join Account C on a.idAccount = C.IdAccount"
			MySql = MySql & " Where a.IdAzienda = " & Session("IdAziendaWork") 
			
			'per i collaboratori devo vedere solo i propri di livello 
			if isBackOffice()=false then 
		       CondRef = "a." & trim(getCondForLevel(session("LivelloAccount"),Session("LoginIdAccount")))
		       MySql = MySql & " and " & CondRef			
			end if 
			
			
			if Cdbl(IdAccountCollaboratore)>0 then 
			   'se ho selezionato un account per BackOffice sono i diretti
			   if IsBackOffice() then 
			      MySql = MySql & " and   a.IdAccountLivello1 = " & IdAccountCollaboratore
               else 
			   'per gli altri i riferimenti diretti
				  CondRef = "a." & trim(getCondForLevel(session("LivelloAccount")+1,IdAccountCollaboratore))
				  MySql = MySql & " and " & CondRef
			   end if
			end if 
	
            MySql = MySql & " and c.FlagAttivo<>'N' "
			MySql = MySql & Condizione & " order By Denominazione"
            'response.write MySql & err.description 
			Rs.CursorLocation = 3 
			Rs.Open MySql, ConnMsde

			DescLoaded=""
			NumCols = numC + 1
			NumRec  = 0
			MsgNoData  = ""
			'elenco azioni 
%>

<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

			<div class="table-responsive"><table class="table"><tbody>
			<thead>
				<tr>
				<th scope="col">Sel</th>
				<th scope="col">Cliente

				</th>
		        <th scope="col">Codice fiscale</th>
		        <th scope="col">Partita Iva</th>
		        <th scope="col">Riferimento</th>
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
					Id=Rs("IdCliente")
					selected=""
		%>
			<tr scope="col">
			
			<td>
			   <button type="button" class="btn btn-info" onclick="callFun(<%=Id%>)">Sel</button>
			</td>			
				<td>
					<input class="form-control" type="text" readonly value="<%=Rs("Denominazione")%>">
				</td>
				<td>
					<input class="form-control" type="text" readonly value="<%=Rs("Codicefiscale")%>">
				</td>
				<td>
					<input class="form-control" type="text" readonly value="<%=Rs("PartitaIva")%>">
				</td>

				<td>
				    <%
					DescRif=""
					IdRif=0 
				    IdRif=Rs("IdAccountLivello1")
                    if isBackOffice()=false then 
					   if session("LivelloAccount")=2 and Rs("IdAccountLivello3")>0 then 
					      IdRif=Rs("IdAccountLivello3")
					   elseif session("LivelloAccount")=1 and Rs("IdAccountLivello2")>0 then    
					      IdRif=Rs("IdAccountLivello2")
					   end if 
					end if 
					if IdRif>0 then 
					   if Cdbl(IdRif)=Session("LoginIdAccount") then 
					      DescRif = "diretto"
					   else 
					      DescRif = LeggiCampo("Select * from Account Where IdAccount=" & IdRif,"Nominativo")
					   end if 
					end if 
					
					
					%>
					<input class="form-control" type="text" readonly value="<%=DescRif%>">
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
