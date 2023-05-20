<%
  NomePagina="ProdottiAttiva.asp"
  titolo="Attivazione prodotti"
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
NameLoaded=NameLoaded & "ValidoDal,DTO;ValidoAl;DTO"

IdAzienda=1
if FirstLoad then 
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
else
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
end if 

on error resume next 
FlagUpdRiferimento=false 
if Oper="INS" or Oper="UPD" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	IdProdotto =Cdbl("0" & Request("IdProdotto"  & KK))
	IdAccountFornitore=Cdbl("0" & Request("IdAccountFornitore" & KK))
	ValidoDal=DataStringa(Request("Dal" & KK))
	if IsNumeric(ValidoDal)=false then 
	   ValidoDal=Dtos()
	end if 
	Validoal =DataStringa(Request("Al" & KK))
	if IsNumeric(ValidoAl)=false then 
	   ValidoAl=20991231
	end if 
	if IdProdotto>0 and IdAccountFornitore>0 then 
	   qUpd=""
       if Oper="INS" then 
	      qUpd = qUpd & " Insert into ProdottoAttivo "
		  qUpd = qUpd & " (IdAzienda,IdProdotto,IdAccountFornitore,ValidoDal,ValidoAl) values "
		  qUpd = qUpd & " (1," & IdProdotto & "," & IdAccountFornitore & "," & ValidoDal & "," & ValidoAl & ")"
	   else
	      qUpd = qUpd & " update ProdottoAttivo set "
		  qUpd = qUpd & " ValidoDal = " & ValidoDal
		  qUpd = qUpd & ",ValidoAl  = " & ValidoAl
		  qUpd = qUpd & " Where IdProdotto = " & IdProdotto
		  qUpd = qUpd & " and  IdAccountFornitore = " & IdAccountFornitore
		  qUpd = qUpd & " and  IdAzienda = 1 "
	   end if 
	   connMsde.execute qUpd 
    end if 	
End if 
if Oper="DEL" then
   Session("TimeStamp")=TimePage
   KK=Request("ItemToRemove")
   IdProdotto =Cdbl("0" & Request("IdProdotto"  & KK))
   IdAccountFornitore=Cdbl("0" & Request("IdAccountFornitore" & KK))
   if IdProdotto>0 and IdFornitore>0 then 
      qUpd = qUpd & " delete from ProdottoAttivo  "
      qUpd = qUpd & " Where IdProdotto = " & IdProdotto
      qUpd = qUpd & " and  IdAccountFornitore = " & IdAccountFornitore
      qUpd = qUpd & " and  IdAzienda = 1 "	
      ConnMsde.execute qUpd 
      If Err.Number <> 0 Then 
         MsgErrore = ErroreDb(Err.description)
      End If
   END if 
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
									<h2 style="height: 25px; margin-left: 1rem;">Attivazione Prodotti</h2>
								</div>
								<div class="card-body">
									<div class="form-group">
										<div class="template-demo d-flex justify-content-between flex-nowrap">
											<%
											AddRow=true
											dim CampoDb(10)
											CampoDB(1)="DescProdotto"
											CampoDB(2)="DescCompagnia"	
											CampoDB(3)="DescFornitore"
											ElencoOption=";0;Prodotto;1;Fornitore;2;Compagnia;3"
											%>	
											<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
										</div>
										<%
										if Firstload then 
										flagAttivo  = ""
										flagCessato = ""
										flagDaAtt   = ""
										flagTutti  = " checked "
										else 
										flagAttivo  = ""
										flagCessato = ""
										flagDaAtt   = ""
										flagTutti   = ""
										if Request("checkAttivo")="A" then 
											flagAttivo  = " checked "
										end if 
										if Request("checkAttivo")="C" then 
											flagCessato = " checked "
										end if 
										if Request("checkAttivo")="D" then 
											flagDaAtt   = " checked "
										end if 
										if Request("checkAttivo")="T" then 
											flagTutti   = " checked "
										end if 			   
										end if    
										%>
										<div class="row">
											<div class="col-1 font-weight-bold">Servizio</div>
											<div class="col-4">
												<div class="form-group">
													<% 
													IdAnagServizio=Request("IdAnagServizio0")
													if IdAnagServizio="-1" then
														IdAnagServizio=""
													end if 
													stdClass="class='form-control form-control-sm'"
													q = ""
													q = q & " Select * from AnagServizio "
													q = q & " order by DescAnagServizio  "
													response.write ListaDbChangeCompleta(q,"IdAnagServizio0",IdAnagServizio ,"IdAnagServizio","DescAnagServizio" ,1,"Sottometti();","","","","",stdClass)
													%>
												</div>		
											</div>

											<div class="col-1 font-weight-bold">Stato Prod.</div>
											<div class="col-6">
												<div class="form-group" style="font-size: large;">
												<input id="checkAttivo<%=l_Id%>" <%=FlagAttivo%> name="checkAttivo<%=l_Id%>" 
												type="radio" value = "A" class="big-checkbox" onclick="Sottometti();">
													<span class="font-weight-bold">Attivo</span>
												<input id="checkAttivo<%=l_Id%>" <%=FlagCessato%> name="checkAttivo<%=l_Id%>" 
												type="radio" value = "C" class="big-checkbox" onclick="Sottometti();">
													<span class="font-weight-bold">cessato</span>
												<input id="checkAttivo<%=l_Id%>" <%=FlagDaAtt%> name="checkAttivo<%=l_Id%>" 
												type="radio" value = "D" class="big-checkbox" onclick="Sottometti();">
													<span class="font-weight-bold">Da Attivare</span>
												<input id="checkAttivo<%=l_Id%>" <%=FlagTutti%> name="checkAttivo<%=l_Id%>" 
												type="radio" value = "T" class="big-checkbox" onclick="Sottometti();">
													<span class="font-weight-bold">Tutti</span>						
												</div>
											</div>
										</div>
										<div class="row">
											<div class="col-1 font-weight-bold">Ramo</div>
											<div class="col-4">
												<div class="form-group ">
														<% 
														IdRamo=Request("IdRamo0")
														if IdRamo="-1" then
															IdRamo=""
														end if 
														stdClass="class='form-control form-control-sm'"
														q = ""
														q = q & " Select * from Ramo "
														q = q & " order by DescRamo  "
														response.write ListaDbChangeCompleta(q,"IdRamo0",IdRamo ,"IdRamo","DescRamo" ,1,"Sottometti();","","","","",stdClass)
														%>
												</div>		
											</div>

											<div class="col-1 font-weight-bold">Rischio</div>
											<div class="col-4">
												<div class="form-group ">
													<% 
													IdSubRamo=Request("IdSubRamo0")
													if IdSubRamo="-1" then
														IdSubRamo=""
													end if 
													stdClass="class='form-control form-control-sm'"
													q = ""
													q = q & " Select * from SubRamo "
													if cdbl("0" & idRamo)>0 then 
														q = q & " Where IdRamo = " & idRamo
													end if 
													q = q & " order by DescSubRamo  "
													response.write ListaDbChangeCompleta(q,"IdSubRamo0",IdSubRamo ,"IdSubRamo","DescSubRamo" ,1,"Sottometti();","","","","",stdClass)
													%>
												</div>		
											</div>
										</div>
									</div>		
									<%
									'caricamento tabella 
									if FirstLoad = false then
									if Condizione<>"" then 
										Condizione=" and " & Condizione
									end if 

									Set Rs = Server.CreateObject("ADODB.Recordset")

									MySql = "" 
									MySql = MySql & " Select a.*,B.DescCompagnia,D.IdAccount as IdAccountfornitore"
									MySql = MySql & ",D.DescFornitore,isNull(E.ValidoDal,0) as Dal,isNull(E.ValidoAl,0) as Al "
									MySql = MySql & " From Prodotto a "
									MySql = MySql & " inner join Compagnia B on a.IdCompagnia = B.IdCompagnia"
									MySql = MySql & " inner join AccountProdotto C on a.idProdotto = C.IdProdotto"
									MySql = MySql & " inner join Fornitore D on C.IdAccount = D.IdAccount"
									MySql = MySql & " left  join ProdottoAttivo E on A.IdProdotto = e.IdProdotto "
									MySql = MySql & "       and C.IdAccount = E.IdAccountFornitore and E.IdAzienda=1"
									MySql = MySql & " Where A.IdCompagnia > 0 "
									if cdbl("0" & IdRamo)>0 then 
									MySql = MySql & "   and A.IdRamo=" &  IdRamo
									end if 
									if cdbl("0" & IdSubRamo)>0 then 
									MySql = MySql & "   and A.IdSubRamo=" &  IdSubRamo
									end if 

									if IdAnagServizio<>"" then 
									MySql = MySql & "   and A.IdAnagServizio='" &  apici(IdAnagServizio) & "'"
									end if 
									if flagTutti="" then 
									if flagDaAtt<>"" then 
										MySql = MySql & "   and isNull(E.ValidoDal,0) = 0"
									else
										MySql = MySql & "   and isNull(E.ValidoDal,0) > 0"
										if flagCessato<>"" then 
											MySql = MySql & "   and isNull(E.ValidoAl,0) < " & Dtos()
										else 
											MySql = MySql & "   and isNull(E.ValidoAl,0) >= " & Dtos()
										end if 
									end if 
									end if 
									MySql = MySql & Condizione & " order By A.DescProdotto,B.DescCompagnia,D.DescFornitore"

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

									<script>
									function localMod(op,id)
									{
									var xx;
									xx=false;
									if (op=="DEL")
										xx=true;
										
									if (op=="INS" || op=="UPD") {

										var dtf = ValoreDi("Dal" + id).trim();
										var dtt = ValoreDi("Al"  + id).trim();
										
										if (op=="INS" && dtf.length==0)
											yy=ImpostaValoreDi("Dal" + id,ValoreDi("DataDiOggi"));
										if (op=="INS" && dtt.length==0)
											yy=ImpostaValoreDi("Al" + id,"31/12/2099");
										yy=ImpostaValoreDi("DescLoaded",id);
										yy=ImpostaValoreDi("NameLoaded","Dal,DTO;Al,DTO");
										xx=ElaboraControlli();  
											
									}
									
									if (xx==false)
										return false;  

									yy=AttivaFunzione(op,id); 
									
									}
									</script>


									<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

									<div class="table-responsive">
										<table class="table">
											<thead>
												<tr>
													<th>
														<h4><b>Prodotto</b></h4>															
													</th>
													<th>
														<h4><b>Compagnia</b></h4>															
													</th>
													<th>
														<h4><b>Fornitore</b></h4>															
													</th>
													<th>
														<h4><b>Attivo Dal</b></h4>															
													</th>
													<th>
														<h4><b>Attivo Al</b></h4>															
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
													IdP=Rs("IdProdotto")
													IdF=Rs("IdAccountfornitore")
													id = idP & "_" & IdF
													DescLoaded=DescLoaded & Id & ";"
													err.clear 
													if Rs("Dal")=0 then 
													ValidoDal=""
													ValidoAl =""
													else
													ValidoDal=Stod(Rs("Dal"))
													ValidoAl =Stod(Rs("Al"))
													end if 
													
													DescCompleto="" 
													
													'controllo fascia per cauzione provvisoria 
													if rs("IdAnagServizio")="CAUZ_PROV" then 
													MySql = "" 
													MySql = MySql & " Select IdProdotto "
													MySql = MySql & " from AccountProdottoFascia "
													MySql = MySql & " where IdProdotto = " & rs("IdProdotto")
													MySql = MySql & " and   IdAccount = " & rs("IdAccountFornitore")
											
													trovato= cdbl("0" & LeggiCampo(MySql,"IdProdotto"))
													if Cdbl(trovato)=0 then 
														DescCompleto=DescCompleto & "manca fascia di calcolo;" & vbNewLine
													end if 
													MySql = "" 
													MySql = MySql & " Select IdProdotto "
													MySql = MySql & " from AccountProdottoFirma "
													MySql = MySql & " where IdProdotto = " & rs("IdProdotto")
													MySql = MySql & " and   IdAccount = " & rs("IdAccountFornitore")
											
													trovato= cdbl("0" & LeggiCampo(MySql,"IdProdotto"))
													if Cdbl(trovato)=0 then 
														DescCompleto=DescCompleto & "manca configurazione firme;" & vbNewLine
													end if 

													end if 
													
													'controllo documentazione
													if Rs("IdListaAffidamento")=0 and (rs("IdAnagServizio")="CAUZ_PROV" or rs("IdAnagServizio")="CAUZ_DEFI") then
													MySql = "" 
													MySql = MySql & " Select IdProdotto "
													MySql = MySql & " from AccountProdottoDocAff "
													MySql = MySql & " where IdProdotto = " & rs("IdProdotto")
													MySql = MySql & " and   TipoDoc = 'AFFI'"
													trovato= cdbl("0" & LeggiCampo(MySql,"IdProdotto"))
													if Cdbl(trovato)=0 then 
														DescCompleto=DescCompleto & "manca documentazione per affidamento;" & vbNewLine
													end if 
													end if 
											%>
		
											<tr>
												<td style="width: 26%;">
													<input class="form-control" type="text" readonly value="<%=Rs("DescProdotto")%>">
													<input type="hidden" name="IdProdotto<%=id%>" value="<%=Rs("IdProdotto")%>">
													<input type="hidden" name="IdAccountFornitore<%=id%>" value="<%=Rs("IdAccountfornitore")%>">
												</td>
												<td>
													<input class="form-control" type="text" readonly value="<%=Rs("DescCompagnia")%>">
												</td>
												<td>
													<input class="form-control" type="text" readonly value="<%=Rs("DescFornitore")%>">
												</td>
													<%if DescCompleto="" then%> 
													<td style="width: 16%;">
														<input type="text" class="form-control mydatepicker" id="Dal<%=id%>" name="Dal<%=id%>" placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" value="<%=ValidoDal%>"/>
													</td>
													<td style="width: 16%;">
														<input type="text" class="form-control mydatepicker" id="Al<%=id%>" name="Al<%=id%>" placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" value="<%=ValidoAl%>"/>
													</td>
													<%else%>
													<td style="width: 16%;">
														<input type="text" class="form-control  id="Dal<%=id%>" name="Dal<%=id%>" readonly value="<%=ValidoDal%>"/>
													</td>
													<td style="width: 16%;">
														<input type="text" class="form-control  id="Al<%=id%>" name="Al<%=id%>" readonly value="<%=ValidoAl%>"/>
													</td>
													<%end if %>
												<td>
													<%if DescCompleto="" then 
													if ValidoDal="" then%>
													
													<%RiferimentoA="col-2;#;;2;plus;Inserisci;;localMod('INS','" & id & "');N"%>
													<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
													<%else%>				
													<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;localMod('UPD','" & Id & "');N"%>
													<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
													<%RiferimentoA="col-2;#;;2;dele;Cancella;;localMod('DEL','" & Id & "');N"%>
													<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
													<%
													end if 
													else
													xx=ShowLabelAlert("",DescCompleto)
													end if %>
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
									<%end if %>
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
