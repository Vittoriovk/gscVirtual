<%
  NomePagina="ProdottoTemplate.asp"
  titolo="Anagrafica Template Prodotto"
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
		<title><%= titolo %></title>
		<link rel="stylesheet" href="../../vendors/feather/feather.css">
		<link rel="stylesheet" href="../../vendors/ti-icons/css/themify-icons.css">
		<link rel="stylesheet" href="../../vendors/css/vendor.bundle.base.css">
		<link rel="stylesheet" href="../../vendors/select2/select2.min.css">
		<link rel="stylesheet" href="../../vendors/select2-bootstrap-theme/select2-bootstrap.min.css">
		<link rel="stylesheet" href="../../vendors/mdi/css/materialdesignicons.min.css">
		<link rel="stylesheet" href="../../vendors/font-awesome/css/font-awesome.min.css" />
		<link rel="stylesheet" href="../../css/vertical-layout-light/style.css">
	</head>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->


<%

if FirstLoad then 
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
else
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
end if 

on error resume next 
if Oper="CALL_INS" or Oper="CALL_UPD" or Oper="CALL_DEL" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
    KK=Request("ItemToRemove")
    Session("swap_IdProdotto")   = KK
    Session("swap_OperTabella")   = Oper
    Session("swap_PaginaReturn")  = "configurazioni/Prodotti/" & NomePagina
    response.redirect virtualPath & "configurazioni/Prodotti/ProdottoTemplateDettaglio.asp"
    response.end 
End if 
if Oper="CALL_TECN" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
    KK=Request("ItemToRemove")
    if Cdbl("0" & KK ) > 0 then 
        Session("swap_IdProdottoTemplate") = KK
        Session("swap_IdTabellaKeyInt")    = KK
        Session("swap_OperTabella")        = Oper
        Session("swap_PaginaReturn")  = "configurazioni/Prodotti/ProdottoTemplate.asp"
        response.redirect virtualPath & "configurazioni/prodotti/ProdottoTemplateListaDati.asp"
        response.end 
    end if 
End if 
  'registro i dati della pagina 
xx=setValueOfDic(Pagedic,"PaginaReturn"   ,PaginaReturn)
xx=setCurrent(NomePagina,livelloPagina) 
'xx=DumpDic(SessionDic,NomePagina)
  
%>
<div class="container-scroller">
	<%
      callP=VirtualPath & "bar/" & Session("TopBar_" & Session("LoginIdAccount")) 
      Server.Execute(callP) 
	%>
    <!-- Page Content -->
	<div class="container-fluid page-body-wrapper">
		<%
			TitoloNavigazione="Configurazioni"
			Session("opzioneSidebar")="conf"
			callP=VirtualPath & "bar/" & Session("sideBar_" & Session("LoginIdAccount")) 
			Server.Execute(callP) 
		%>	
		<div class="main-panel">          
			<div class="content-wrapper">
				<div class="row">
					<div class="col-lg-12 grid-margin stretch-card">
						<div class="card">
                            <form name="Fdati" Action="<%=NomePagina%>" method="post">
                                <div class="card-body d-flex">	
									<%RiferimentoA=";" & VirtualPath & "SupervisorConfigurazioni.asp;;2;prev;Indietro;;;"%>
									<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
									<h2 style="height: 25px; margin-left: 1rem;">Elenco Template Prodotti</h2>
								</div>
								<div class="card-body">
									<div class="form-group">
										<div class="template-demo d-flex justify-content-between flex-nowrap">
											<div style="float:left; display:block;">
												<button style="font-size: 1.25rem; margin-right: 1rem;" class="add btn btn-primary todo-list-add-btn" type="button" onclick="AttivaFunzione('CALL_INS','0');">Inserisci Template Prodotto</button>
											</div>
                                            <%
                                            AddRow=true
                                            dim CampoDb(10)
                                            CampoDB(1)="DescProdottoTemplate"    
                                            CampoDB(2)="DescRamo"    
                                            ElencoOption=";0;Template Prodotto;1;Ramo;2"
                                            %>  
											<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
										</div>
									</div>
                                    <%
                                    'caricamento tabella 
                                    if Condizione<>"" then 
                                        Condizione=" and " & Condizione
                                    end if 
                                            
                                    Set Rs = Server.CreateObject("ADODB.Recordset")

                                    MySql = "" 
                                    MySql = MySql & " Select a.*,B.DescRamo "
                                    MySql = MySql & " From ProdottoTemplate a,Ramo B"
                                    MySql = MySql & " Where a.IdProdottoTemplate > 0 "
                                    MySql = MySql & " and   a.IdRamo = B.idRamo "
                                    MySql = MySql & Condizione & " order By DescProdottoTemplate"

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

                                    <div class="table-responsive">
										<table class="table">
											<thead>
												<tr>
													<th>
														<h4><b>Template Prodotto</b></h4>															
													</th>
                                                    <th>
														<h4><b>Ramo - Servizio Riferimento</b></h4>															
													</th>
													<th class="text-center">
														<h4><b>Azioni</b></h4>
													</th>
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
                                                    Id=Rs("IdProdottoTemplate")
                                                    err.clear 
                                            %>
        
                                            <tr>
                                                <td>
                                                    <input class="form-control" type="text" readonly value="<%=Rs("DescProdottoTemplate")%>">
                                                </td>
                                                <td>
                                                    <%
                                                    idAnagServ=Rs("idAnagServizio")
                                                    DescServ=Leggicampo("select * from AnagServizio Where IdAnagServizio='" & apici(IdAnagServ) & "'","DescAnagServizio") 
                                                    %>
                                                    <input class="form-control" type="text" readonly value="<%=Rs("DescRamo") & " - " & descServ%>">
                                                </td>            
                                                <td class="text-center">
                                                    <%
                                                    IdAnagCaratteristicaKey=""
                                                    IdAnagCaratteristica = rs("IdAnagCaratteristica")
                                                    if cdbl(IdAnagCaratteristica)>0 then 
                                                        q = ""
                                                        q = q & " select IdAnagCaratteristicaKey From AnagCaratteristica "
                                                        q = q & " where IdAnagCaratteristica =" & IdAnagCaratteristica
                                                        'response.write q
                                                        IdAnagCaratteristicaKey = LeggiCampo(q,"IdAnagCaratteristicaKey")
                                                    end if 
                                    
                                                    IdAnagServizio=rs("IdAnagServizio")
                                                    q = ""
                                                    q = q & " select distinct IdTipoOpzione From Opzione "
                                                    q = q & " where IdAnagServizio='" & IdAnagServizio & "'" 
                                                    if IdAnagCaratteristicaKey="" then 
                                                        q = q & " and IdAnagCaratteristicaKey='" & IdAnagCaratteristicaKey & "'" 
                                                    else
                                                        q = q & " and IdAnagCaratteristicaKey like '%|" & IdAnagCaratteristicaKey & "|%'" 
                                                    end if 
                                                    q = q & " And IdTipoOpzione='"

                                                    ContaOpziTecn=LeggiCampo(q & "TECN'","IdTipoOpzione")
                                                    ContaOpziOpzi=LeggiCampo(q & "OPZI'","IdTipoOpzione")			
                                                    %>

                                                    
                                                    <%RiferimentoA="col-2;#;;2;upda;Aggiorna;;AttivaFunzione('CALL_UPD','" & Id & "');N"%>
                                                    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->&nbsp;            
                                                    <%RiferimentoA="col-2;#;;2;dele;Cancella;;AttivaFunzione('CALL_DEL','" & Id & "');N"%>
                                                    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->&nbsp;   
                                                    <%if contaOpziTecn<>"" or 1=1 then %>
                                                        <%RiferimentoA="col-2;#;;2;manu;Dati Aggiuntivi;;AttivaFunzione( 'CALL_TECN','" & Id & "');N"%>
                                                        <!--#include virtual="/gscVirtual/include/Anchor.asp"-->                 
                                                    <%end if%>
                                                </td>
                                            </tr>
                                            <%    
                                            rs.MoveNext
                                            Loop
                                            end if 
                                            rs.close

                                            %>
                                        </table>
									</div>
									<div class="mt-3">
										<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
										<!--#include virtual="/gscVirtual/include/paginazione.asp"-->
									</div>			
								</div>	
							</form>
						</div>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>

<script TYPE="text/javascript">
	function mouseover(elem) {
		elem.style.color = '#FF0000';
	}
	function mouseout(elem) {
		elem.style.color = '#4B49AC';
	}
</script>
<script>
	(function() {
			var dialog = document.getElementById('myFirstDialog');
			document.getElementById('show').onclick = function() {
				dialog.show();
			};
			document.getElementById('hide').onclick = function() {
				dialog.close();
			};
	})();
</script>

<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->
<script src="../../vendors/js/vendor.bundle.base.js"></script>
<script src="../../vendors/typeahead.js/typeahead.bundle.min.js"></script>
<script src="../../vendors/select2/select2.min.js"></script>
<script src="../../js/off-canvas.js"></script>
<script src="../../js/hoverable-collapse.js"></script>
<script src="../../js/template.js"></script>
<script src="../../js/settings.js"></script>
<script src="../../js/typeahead.js"></script>
<script src="../../js/select2.js"></script>

</body>

</html>
