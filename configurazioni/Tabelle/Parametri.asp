<%
  NomePagina="Parametri.asp"
  titolo="Menu Supervisor - Gestione Parametri"
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
<!--#include virtual="/gscVirtual/js/functionCheckTable.js"-->
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
   NameLoaded= "IdTrattFiscDiritti,LI;IdTrattFiscIntermediazione,LI"
   NameRangeN = ""
  
   ParamVuoto = false 
   tmp=LeggiCampo("select top 1 * from Parametri ","IdTrattFiscaleDirittiEmissione")
   if tmp="" then 
      ParamVuoto = true
   end if 
   err.clear 
   IdTrattFiscDiritti         = TestNumeroPos(Request("IdTrattFiscDiritti0"))
   IdTrattFiscIntermediazione = TestNumeroPos(Request("IdTrattFiscIntermediazione0")) 
   FlagBorsellino = TestNumeroPos(Request("FlagBorsellino0")) 
   FlagFido       = TestNumeroPos(Request("FlagFido0")) 
   FlagEstratto   = TestNumeroPos(Request("FlagEstratto0")) 
   MsgNoData=""

   if Oper=ucase("update") then 
      Session("TimeStamp")=TimePage
      MyQ = "" 

      if ParamVuoto=true then 
         MyQ = MyQ & " INSERT INTO Parametri (IdTrattFiscaleDirittiEmissione,IdTrattFiscaleIntermediazione"
         MyQ = MyQ & ",FlagBorsellino,flagFido,FlagEstratto) values ( " 
         MyQ = MyQ & " " & IdTrattFiscDiritti
         MyQ = MyQ & "," & IdTrattFiscIntermediazione     
		 MyQ = MyQ & "," & FlagBorsellino    
		 MyQ = MyQ & "," & FlagFido    
		 MyQ = MyQ & "," & FlagEstratto    
         MyQ = MyQ & " ) "
      else 
         MyQ = MyQ & " update Parametri set "
         MyQ = MyQ & " IdTrattFiscaleDirittiEmissione = " & NumForDB(IdTrattFiscDiritti)
         MyQ = MyQ & ",IdTrattFiscaleIntermediazione = " & NumForDB(IdTrattFiscIntermediazione)
		 MyQ = MyQ & ",FlagBorsellino = " & NumForDB(FlagBorsellino)
		 MyQ = MyQ & ",FlagFido = " & NumForDB(FlagFido)
		 MyQ = MyQ & ",FlagEstratto = " & NumForDB(FlagEstratto)
      end if 

      ConnMsde.execute MyQ 
      If Err.Number <> 0 Then 
         MsgErrore = ErroreDb(Err.description)
      End If
   end if 

   
   DescPageOper="Aggiornamento"

  'registro i dati della pagina 
   xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
   xx=setCurrent(NomePagina,livelloPagina) 

   DescLoaded="0"  
  
  'recupero i dati 
  if ParamVuoto=false then
      MySql = ""
      MySql = MySql & " Select * From  Parametri "
 
      Set Rs = Server.CreateObject("ADODB.Recordset")

      Rs.CursorLocation = 3 
      Rs.Open MySql, ConnMsde 
      IdTrattFiscDiritti         = RS("IdTrattFiscaleDirittiEmissione")
      IdTrattFiscIntermediazione = RS("IdTrattFiscaleIntermediazione")
 
      rs.close 
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
                        <div class="card-body">
                        <div class="card-body d-flex">	
									<%RiferimentoA=";" & VirtualPath & "SupervisorConfigurazioni.asp;;2;prev;Indietro;;;"%>
									<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
									<h2 style="height: 25px; margin-left: 1rem;">Gestione Parametri : <%=DescPageOper%> </h2>
								</div>
                        <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
                        <div class="row">
                           <div class="col-8">
                              <%xx=ShowLabel("Tratt.Fiscale (imposte) per diritti di emissione")%>
                              
                              <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            
                                 <%
                                 ao_lbd = ""             'descrizione label 
                                 ao_nid = "IdTrattFiscDiritti0"          'nome ed id
                                 ao_val = IdTrattFiscDiritti
                                 ao_Att = "1"                       'indica se deve mettere vuoto 
                                 ao_Tex = "select * from TrattamentoFiscale order By DescTrattamentoFiscale"
                                 'response.write ao_Tex
                                 ao_ids = "IdTrattamentoFiscale"             'valore della select 
                                 ao_des = "DescTrattamentoFiscale"           'valore del testo da mostrare 
                                 ao_cla = ""                        'azzero classe
                                 ao_Eve = ""                        'azzero evento
                                 
                                 ao_Plh = ""                        'indica cosa mettere in caso di vuoto
                                 ao_Cla = "class='form-control form-control-sm'"      
                                 %>
                              <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->   
                           </div>
                        </div>
                        <div class="row">
                           <div class="col-8">
                              <%xx=ShowLabel("Tratt.Fiscale (imposte) per Intermediazione")%>   
                              <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            
                              <%
                                 ao_lbd = ""             'descrizione label 
                                 ao_nid = "IdTrattFiscIntermediazione0"          'nome ed id
                                 ao_val = IdTrattFiscIntermediazione
                                 ao_Att = "1"                       'indica se deve mettere vuoto 
                                 ao_Tex = "select * from TrattamentoFiscale order By DescTrattamentoFiscale"
                                 'response.write ao_Tex
                                 ao_ids = "IdTrattamentoFiscale"             'valore della select 
                                 ao_des = "DescTrattamentoFiscale"           'valore del testo da mostrare 
                                 ao_cla = ""                        'azzero classe
                                 ao_Eve = ""                        'azzero evento
                                 
                                 ao_Plh = ""                        'indica cosa mettere in caso di vuoto
                                 ao_Cla = "class='form-control form-control-sm'"      
                              %>
                              <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->      
                              </div>
                        </div> 
                        <div class="row">
                              <div class="col-5 text-center">
                                 <%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('update','0');S"%>
                                 <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                              </div>
                        </div>
                        
                           <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
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
