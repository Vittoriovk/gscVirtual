<%
  NomePagina="ServizioDocumento.asp"
  titolo="Associazione Documenti "
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
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
NameLoaded=NameLoaded & "IdDocumento,LI"

IdAnagServizio=""
if FirstLoad then 
   IdAnagServizio = getCurrentValueFor("IdAnagServizio")
   IdTipoUtenza   = getCurrentValueFor("IdTipoUtenza")
   PaginaReturn   = getCurrentValueFor("PaginaReturn") 
else
   IdAnagServizio = getValueOfDic(Pagedic,"IdAnagServizio")
   IdTipoUtenza   = getValueOfDic(Pagedic,"IdTipoUtenza")
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
   
end if 
if IdAnagServizio="" or IdTipoUtenza="" then 
   response.redirect RitornaA(PaginaReturn)
   response.end 
end if 
IdAnagServizio = trim(IdAnagServizio)
DescLista      = LeggiCampo("Select * from AnagServizio where IdAnagServizio='" & IdAnagServizio & "'","DescAnagServizio") 

on error resume next 
FlagUpdLista=false 
if Oper="INS" then 
    Session("TimeStamp")=TimePage
	KK="0"
	IdDocumento = Request("IdDocumento" & KK)
	checkbox    = Request("checkbox" & KK)
	if checkbox="S" then 
	   FlagObbligatorio=1
	else
	   FlagObbligatorio=0
	end if 
	FlagScadenza = 1

	if Cdbl(IdDocumento)>0 then 
		MyQ = "" 
		MyQ = MyQ & " Insert into ServizioDocumento ("
		MyQ = MyQ & " IdAnagServizio,IdTipoUtenza,IdDocumento,FlagObbligatorio,FlagDataScadenza,DITT,PEGI,PEFI,PEGC"
		MyQ = MyQ & ") values ("			
		MyQ = MyQ & " '" & IdAnagServizio  & "'"
		MyQ = MyQ & ",'" & IdTipoUtenza  & "'"
		MyQ = MyQ & ", " & IdDocumento
		MyQ = MyQ & ", " & FlagObbligatorio	
		MyQ = MyQ & ", " & FlagScadenza
		MyQ = MyQ & ",'" & Request("checkDITT" & KK) & "'"
		MyQ = MyQ & ",'" & Request("checkPEGI" & KK) & "'"
		MyQ = MyQ & ",'" & Request("checkPEFI" & KK) & "'"
		MyQ = MyQ & ",'" & Request("checkPEGC" & KK) & "'"
		MyQ = MyQ & ")"

		ConnMsde.execute MyQ 
		If Err.Number <> 0 Then 
			MsgErrore = ErroreDb(Err.description)
		else
			FlagUpdLista=true
			DescIn=""
		End If
	END if 
End if 
if Oper="UPD" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	checkbox    = Request("checkbox" & KK)
	if checkbox="S" then 
	   FlagObbligatorio=1
	else
	   FlagObbligatorio=0
	end if 	
	MyQ = "" 
	MyQ = MyQ & " update ServizioDocumento set "
	MyQ = MyQ & " FlagObbligatorio= " & FlagObbligatorio	
	MyQ = MyQ & ",DITT='"             & Request("checkDITT" & KK) & "'"
	MyQ = MyQ & ",PEGI='"             & Request("checkPEGI" & KK) & "'"
	MyQ = MyQ & ",PEFI='"             & Request("checkPEFI" & KK) & "'"
	MyQ = MyQ & ",PEGC='"             & Request("checkPEGC" & KK) & "'"
	MyQ = MyQ & " where IdAnagServizio = '" & IdAnagServizio & "'"
	MyQ = MyQ & " and   IdTipoUtenza = '"   & IdTipoUtenza & "'"
	MyQ = MyQ & " and   IdDocumento = " & KK

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else
	    FlagUpdLista=true
	End If
	DescIn=""
End if

if Oper="DEL" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	MyQ = "" 
	MyQ = MyQ & " delete from ServizioDocumento "
	MyQ = MyQ & " where IdAnagServizio = '" & IdAnagServizio & "'"
	MyQ = MyQ & " and   IdTipoUtenza = '"   & IdTipoUtenza & "'"
	MyQ = MyQ & " and   IdDocumento = " & KK

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else
	    FlagUpdLista=true
	End If
	DescIn=""
End if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdAnagServizio" ,IdAnagServizio)
  xx=setValueOfDic(Pagedic,"IdTipoUtenza"   ,IdTipoUtenza)
  xx=setValueOfDic(Pagedic,"PaginaReturn"   ,PaginaReturn)
  
  xx=setCurrent(NomePagina,livelloPagina) 

  DescTipoUtenza = "Coobbligato"
  if IdTipoUtenza="ATI" then 
     DescTipoUtenza = "ATI (Associazione temporanea di impresa )"
  end if 
  if IdAnagServizio="LISTA" then 
     DescLista = "Lista Documenti "
	 DescTipoUtenza = LeggiCampo("Select * from ListaDocumento Where IdListaDocumento=" & IdTipoUtenza,"DescListaDocumento")
  end if   
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
									<h2 style="height: 25px; margin-left: 1rem;">Elenco Documenti</h2>
								</div>
								<div class="row">
									<div class="col-1">
									</div>
									<div class="col-4 form-group ">
										<%xx=ShowLabel("Servizio")%>
										<input type="text" readonly class="form-control input-sm" value="<%=DescLista%>" >
									</div>	
									<div class="col-2">
									</div>
									<div class="col-4 form-group ">
										<%xx=ShowLabel("relativo a ")%>
										<input type="text" readonly class="form-control input-sm" value="<%=DescTipoUtenza%>" >
									</div>	
								</div>
								<div class="card-body">
									<div class="form-group">
										<div class="template-demo d-flex justify-content-between flex-nowrap">
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
									MySql = MySql & " Select a.FlagObbligatorio,a.DITT,a.PEGI,a.PEFI,a.PEGC,b.* "
									MySql = MySql & " From ServizioDocumento a, documento B "
									MySql = MySql & " Where A.IdAnagServizio = '" & IdAnagServizio & "'"
									MySql = MySql & " And A.IdTipoUtenza = '" & IdTipoUtenza & "'"
									MySql = MySql & " And A.IdDocumento = B.IdDocumento "
									MySql = MySql & " And B.DescDocumento LIKE '%" & apici(cerca_testo) & "%' "
									MySql = MySql & " order By B.DescDocumento"

									Rs.CursorLocation = 3 
									Rs.Open MySql, ConnMsde

									DescLoaded=""
									NumCols = 3
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
														<h4><b>Documento</b></h4>															
													</th>
													<th class="text-center" width="10%">
														<h4><b>Obbligatorio</b></h4>															
													</th>
													<th class="text-center" width="10%">
														<h4><b>Ditta</b></h4>															
													</th>
													<th class="text-center" width="10%">
														<h4><b>Pers.Fis.</b></h4>															
													</th>
													<th class="text-center" width="10%">
														<h4><b>Pers.Giu.Cap.</b></h4>															
													</th>
													<th class="text-center" width="10%">
														<h4><b>Pers.Giu.Pers.</b></h4>															
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
											Do While Not rs.EOF and (NumRec<PageSize or Pagesize<=0)
											Primo=Primo+1
											NumRec=NumRec+1
											Id=Rs("IdDocumento")
											DescLoaded=DescLoaded & Id & ";"
											if Rs("FlagObbligatorio")=0 then 
											FlagObbligatorio=""
											else
											FlagObbligatorio=" checked "
											end if
											if Rs("DITT")="" then 
											FlagDITT=""
											else
											FlagDITT=" checked "
											end if				
											if Rs("PEGI")="" then 
											FlagPEGI=""
											else
											FlagPEGI=" checked "
											end if				
											if Rs("PEGC")="" then 
											FlagPEGC=""
											else
											FlagPEGC=" checked "
											end if					
											if Rs("PEFI")="" then 
											FlagPEFI=""
											else
											FlagPEFI=" checked "
											end if								
											%> 
											<tr> 
												<td>
													<input class="form-control" Id="IdDocumento<%=Id%>" type="text" readonly value="<%=Rs("DescDocumento")%>">
												</td>
												<td>
													<div class="form-check text-center">
														<input  id="checkbox<%=Id%>" <%=FlagObbligatorio%> name="checkbox<%=Id%>" type="checkbox" value = "S" class="big-checkbox">
													</div>		
												</td>
												<td>
													<div class="form-check text-center">
														<input  id="checkDITT<%=Id%>" <%=FlagDITT%> name="checkDITT<%=Id%>" type="checkbox" value = "DITT" class="big-checkbox">
													</div>		
												</td>
												<td>
													<div class="form-check text-center">
														<input  id="checkPEFI<%=Id%>" <%=FlagPEFI%> name="checkPEFI<%=Id%>" type="checkbox" value = "PEFI" class="big-checkbox">
													</div>		
												</td>
												<td>
													<div class="form-check text-center">
														<input  id="checkPEGC<%=Id%>" <%=FlagPEGC%> name="checkPEGC<%=Id%>" type="checkbox" value = "PEGI" class="big-checkbox">
													</div>		
												</td>					
												<td>
													<div class="form-check text-center">
														<input  id="checkPEGI<%=Id%>" <%=FlagPEGI%> name="checkPEGI<%=Id%>" type="checkbox" value = "PEGI" class="big-checkbox">
													</div>		
												</td>
												
												<td class="text-center">
													<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;SalvaSingoloEdAttiva('UPD'," & Id & ",true,'','','');N"%>
													<!--#include virtual="/gscVirtual/include/Anchor.asp"-->&nbsp;			
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
											<%
											if ShowNew then 
											Id=0
											%>
											<tr> 
												<td>
													<% 	IdRef="IdDocumento" & Id 	
													inQue = ""
													inQue = inQue & "(select IdDocumento from ServizioDocumento "
													inQue = inQue & " where IdAnagServizio='" & IdAnagServizio & "'"
													inQue = inQue & " and IdTipoUtenza = '" & IdTipoUtenza & "'"
													inQue = inQue & ") "
													query = ""
													query = query & " Select * from documento " 
													query = query & " Where IdDocumento not in " & inQue 
													query = query & " and IdDocumentoInterno='' "  
													query = query & " order By DescDocumento"
													
													response.write ListaDbChangeCompleta (Query,IdRef,"0","IdDocumento","DescDocumento",0,"","IdDocumento","","","dati assenti","class='form-control form-control-sm'")
													
													xx="0" & LeggiCampo(query,"IdDocumento")
													%>
												</td>
												<%if Cdbl(xx)>0 then %>
												<td>
													<div class="form-check text-center">
														<input id="checkbox0" name="checkbox0" type="checkbox" value = "S" class="big-checkbox">
													</div>
													<td>
														<div class="form-check text-center">
															<input id="checkDITT0" checked name="checkDITT0" type="checkbox" value = "DITT" class="big-checkbox">
														</div>		
													</td>
													<td>
														<div class="form-check text-center">
															<input id="checkPEFI0" checked name="checkPEFI0" type="checkbox" value = "PEFI" class="big-checkbox">
														</div>		
													</td>
													<td>
														<div class="form-check text-center">
															<input id="checkPEGC0" checked name="checkPEGC0" type="checkbox" value = "PEGC" class="big-checkbox">
														</div>		
													</td>					
													<td>
														<div class="form-check text-center">
															<input id="checkPEGI0" checked name="checkPEGI0" type="checkbox" value = "PEGI" class="big-checkbox">
														</div>		
													</td>		
												</td>
												<%end if %>
												<td class="text-center">
													<%if Cdbl(xx)>0 then %>
													<%RiferimentoA="col-2;#;;2;insert;Inserisci;;SaveWithOper('INS')"%>
													<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
													<%end if %>
												</td>
											</tr>			
											<%end if%>
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
