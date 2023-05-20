<%
  NomePagina="RamoDettaglio.asp"
  titolo="Gestione Ramo"
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
<script language="JavaScript">

function localFun(Op,Id)
{
	xx=ImpostaValoreDi("DescLoaded","0");
	xx=ElaboraControlli();
	
 	if (xx==false)
	   return false;
	
	ImpostaValoreDi("Oper","update");
	document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

 
 <!-- javascript locale -->
<script>
function localSubmit(Op)
{
var xx;
    xx=false;
	if (Op=="submit")
	   xx=ElaboraControlli();
   	
 	if (xx==false)
	   return false;
		
	ImpostaValoreDi("Oper","update");
	document.Fdati.submit(); 
}
</script>

<%
  NameLoaded= ""
  NameLoaded= NameLoaded & "DescRamo,TE" 
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdRamo=0
  if FirstLoad then 
	 IdRamo   = "0" & Session("swap_IdRamo")
	 if Cdbl(IdRamo)=0 then 
		IdRamo = cdbl("0" & getValueOfDic(Pagedic,"IdRamo"))
	 end if 
	 OperTabella   = Session("swap_OperTabella")
	 PaginaReturn    = getCurrentValueFor("PaginaReturn")
  else
	 IdRamo   = "0" & getValueOfDic(Pagedic,"IdRamo")
	 OperTabella   = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdRamo = cdbl(IdRamo)
  'inizio elaborazione pagina

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
  
   xx=SetDiz(DizDatabase,"IdRamo",0)
   xx=SetDiz(DizDatabase,"DescRamo","")
  
  'recupero i dati 
  if cdbl(IdRamo)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  Ramo "
	  MySql = MySql & " Where IdRamo=" & IdRamo
	  xx=GetInfoRecordset(DizDatabase,MySql)
  end if 
  
  if OperTabella="CALL_DEL" then 
     SoloLettura=true
  end if 
  'inserisco il fornitore 
  descD  = Request("DescRamo0")
  
  if Oper=ucase("update") and OperTabella="CALL_INS" then 
  
    Session("TimeStamp")=TimePage
	KK="0"
	MyQ = "" 
	MyQ = MyQ & " INSERT INTO Ramo (DescRamo,IdAnagRamo) " 
	MyQ = MyQ & " values ('" & apici(descD) & "','" & apici(request("IdAnagRamo0")) & "')" 

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdRamo)>0 then 
	MyQ = "" 
	MyQ = MyQ & " Update Ramo "
	MyQ = MyQ & " Set DescRamo = '" & apici(descD) & "'"
	MyQ = MyQ & ",IdAnagRamo = '" & apici(trim(request("IdAnagRamo0"))) & "'"
	MyQ = MyQ & " Where IdRamo = " & IdRamo

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
  end if

  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdRamo)>0 then 
     MsgErrore = VerificaDel("Ramo",IdRamo) 
	 if MsgErrore = "" then   
		MyQ = "" 
		MyQ = MyQ & " Delete from Ramo "
		MyQ = MyQ & " Where IdRamo = " & IdRamo

		ConnMsde.execute MyQ 
		If Err.Number <> 0 Then 
			MsgErrore = ErroreDb(Err.description)
		else 
		   response.redirect virtualpath & PaginaReturn
		End If	
	end if 
  end if  
  
   DescPageOper="Aggiornamento"
   if OperTabella="V" then 
      DescPageOper = "Consultazione"
   elseIf OperTabella="CALL_INS" then 
      DescPageOper = "Inserimento"
   elseIf OperTabella="CALL_DEL" then 
      DescPageOper = "Cancellazione"	  
   end if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdRamo"  ,IdRamo)
  xx=setValueOfDic(Pagedic,"OperTabella"  ,OperTabella)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescLoaded="0"  
  %>

<% 
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
									<%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
									<!--#include virtual="/gscVirtual/include/Anchor.asp"-->	
									<h2 style="height: 25px">Gestione Ramo: <%=DescPageOper%> </h2>
								</div>
								<div class="card-body d-flex">
									<div class="form-group">
										<%
										kk="DescRamo" 
										NameLoaded= NameLoaded & kk & ",TE;" 
										%>
										<h4 style="height: 25px">Descrizione Ramo</h4>
										<input type="text" style="width: 150%;" class="form-control form-control-lg" Id="<%=KK%>0" name="<%=KK%>0" value="<%=GetDiz(DizDatabase,"DescRamo") %>" >
									
										<%if false then %>
								
										<%
										kk="IdAnagRamo" 
										xx=ShowLabel("Ramo di riferimento")
										NameLoaded= NameLoaded & kk & ",LI;" 
										q = ""
										q = q & " select * from AnagRamo "
										q = q & " where IdAnagRamo not in (select IdAnagRamoPadre from AnagRamo)"
										q = q & " order by descAnagRamo"
										
										stdClass="class='form-control form-control-sm'"
										response.write ListaDbChangeCompleta(q,"IdAnagRamo0",GetDiz(DizDatabase,"IdAnagRamo") ,"IdAnagRamo","DescAnagRamo" ,tt,"","","","","",stdClass)
										%>
										<%end if %>
										<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->		
										<%if SoloLettura=false then%>
										<div class="mt-3">
											<div style="float:left; display:block;">
												<button style="font-size: 1.25rem; margin-right: 1rem;" class="add btn btn-primary todo-list-add-btn" type="button" onclick="localFun('submit','0');">Inserisc</button>		
											</div>
											<%elseif OperTabella="CALL_DEL" then  %>
											<div style="float:left; display:block;">
												<button style="font-size: 1.25rem; margin-right: 1rem;" class="add btn btn-primary todo-list-add-btn" type="button" onclick="localFun('submit','0');">Rimuovi</button>			
											</div>
										</div>
									</div>
								</div> 
								<%end if %>
							<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
							</form>
						</div>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
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
