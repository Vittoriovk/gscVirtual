<%
  NomePagina="DatoTecnicoDettaglio.asp"
  titolo="Menu Supervisor - Dashboard"
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
function CheckDatiTecnici(Id)

{
	xx=ImpostaColoreFocus("IdElenco" + Id,"","white");
	locName=ValoreDi("DescLoaded");
	yy=ImpostaValoreDi("DescLoaded",Id);

	xx=ElaboraControlli();
	var te = ValoreDi("IdTipoCampo" + Id);
	var el = ValoreDi("IdElenco" + Id);

	te = te.toUpperCase();
	var fl = (te=="MENUDI" || te=="SCELTA" || te=="SPUNTA" );

	if (xx==true && fl==true ) {
		xx=ControllaCampo("IdElenco" + Id,"LI");
		if (xx==false)
			bootbox.alert("Elenco richiesto");
	}
	if (xx==true && fl==false && trim(el)!="-1" ) {
		xx=ImpostaColoreFocus("IdElenco" + Id,"","yellow");
		bootbox.alert("Elenco NON richiesto");
		xx=false;
	}


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
   NameLoaded= "DescDatoTecnico,TE"

   FirstLoad=(Request("CallingPage")<>NomePagina)
   IdDatoTecnico=0
   if FirstLoad then 
      IdDatoTecnico   = "0" & Session("swap_IdDatoTecnico")
      if Cdbl(IdDatoTecnico)=0 then 
         IdDatoTecnico = cdbl("0" & getValueOfDic(Pagedic,"IdDatoTecnico"))
      end if 
      OperTabella   = Session("swap_OperTabella")
      PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
      if PaginaReturn="" then 
        PaginaReturn = Session("swap_PaginaReturn")
      end if 
   else
      IdDatoTecnico   = "0" & getValueOfDic(Pagedic,"IdDatoTecnico")
      OperTabella   = getValueOfDic(Pagedic,"OperTabella")
      PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   end if 
   IdDatoTecnico = cdbl(IdDatoTecnico)
  
  if OperTabella="CALL_DEL" then 
     SoloLettura=true
  end if 

  descD    = Request("DescDatoTecnico0")  
  IdTipo   = Request("IdTipoCampo0")
  Idelenco = Request("Idelenco0")
  if Idelenco="-1" then 
     IdElenco=0
  end if 
  MsgNoData=""

  if Oper=ucase("update") and OperTabella="CALL_INS" then 
  
    Session("TimeStamp")=TimePage
	KK="0"
	MyQ = "" 
	MyQ = MyQ & " INSERT INTO DatoTecnico (DescDatoTecnico,IdTipoCampo,IdElenco) " 
	MyQ = MyQ & " values ('" & apici(descD) & "'" 
    MyQ = MyQ & " ,'" & apici(IdTipo) & "'" 
	MyQ = MyQ & " , " & Idelenco
    MyQ = MyQ & " )" 

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdDatoTecnico)>0 then 
	MyQ = "" 
	MyQ = MyQ & " Update DatoTecnico "
	MyQ = MyQ & " Set DescDatoTecnico = '" & apici(descD) & "'"
	MyQ = MyQ & ",IdTipoCampo='" & apici(IdTipo) & "'"
	MyQ = MyQ & ",IdElenco= " & IdElenco
	MyQ = MyQ & " Where IdDatoTecnico = " & IdDatoTecnico

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
  end if

  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdDatoTecnico)>0 then 
     MsgErrore = VerificaDel("DatoTecnico",IdDatoTecnico) 
	 if MsgErrore = "" then   
		MyQ = "" 
		MyQ = MyQ & " Delete from DatoTecnico "
		MyQ = MyQ & " Where IdDatoTecnico = " & IdDatoTecnico

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
   xx=setValueOfDic(Pagedic,"IdDatoTecnico"  ,IdDatoTecnico)
   xx=setValueOfDic(Pagedic,"OperTabella"  ,OperTabella)
   xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
   xx=setCurrent(NomePagina,livelloPagina) 

   DescLoaded="0"  
  
  'recupero i dati 
  if cdbl(IdDatoTecnico)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  DatoTecnico "
	  MySql = MySql & " Where IdDatoTecnico=" & IdDatoTecnico
 
	  Set Rs = Server.CreateObject("ADODB.Recordset")

      Rs.CursorLocation = 3 
      Rs.Open MySql, ConnMsde 
      DescDatoTecnico = rs("descDatoTecnico")
      IdTipo          = rs("IdTipoCampo")
      IdElenco        = Rs("IdElenco")	  
	  rs.close 
  end if 
   
  'xx=DumpDic(SessionDic,NomePagina)
%>
<div class="d-flex" id="wrapper">
	<%
	  Session("opzioneSidebar")="dash"
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
			<%RiferimentoA="col-1 text-center;" & virtualpath & PaginaReturn & ";;2;prev;Indietro;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h3>Gestione Dato Tecnico : <%=DescPageOper%> </b> </h3>
				</div>
			</div>
   <br>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
   <br>
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   l_Id = "0"
   
   ao_lbd = "Descrizione Dato Tecnico"       'descrizione label 
   ao_nid = "DescDatoTecnico" & l_Id            'nome ed id
   ao_val = "|value=" & DescDatoTecnico       'valore di default
   ao_Plh = "|placeholder=Descrizione Dato Tecnico"              'placeholder
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   	<%
	   ao_lbd = "Tipo Dato"             'descrizione label 
       ao_nid = "IdTipoCampo0"          'nome ed id
       ao_val = IdTipo
	   ao_Tex = "select * from TipoCampo "
	   if SoloLettura=true then
	      ao_Tex = ao_Tex & "Where IdTipoCampo='" & apici(ao_val) & "' "
	   end if 
	   
	   ao_Tex = ao_Tex & " order By DescTipoCampo" 
	   
	   ao_ids = "IdTipoCampo"             'valore della select 
	   ao_des = "DescTipoCampo"           'valore del testo da mostrare 
	   ao_cla = ""                        'azzero classe
	   ao_Eve = ""                        'azzero evento
	   ao_Att = "0"                       'indica se deve mettere vuoto 
	   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
	   ao_Cla = "class='form-control form-control-sm'"	  
	%>
	<!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->   

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   	<%
	   ao_lbd = "Elenco"             'descrizione label 
       ao_nid = "IdElenco0"          'nome ed id
       ao_val = IdElenco
	   ao_Tex = "select * from Elenco order By DescElenco"
	   'response.write ao_Tex
	   ao_ids = "IdElenco"             'valore della select 
	   ao_des = "Descelenco"           'valore del testo da mostrare 
	   ao_cla = ""                        'azzero classe
	   ao_Eve = ""                        'azzero evento
	   ao_Att = "1"                       'indica se deve mettere vuoto 
	   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
	   ao_Cla = "class='form-control form-control-sm'"	  
	%>
	<!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->  	
   <%if SoloLettura=false then%>
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;CheckDatiTecnici('0');S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div></div>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
		</div>
	<%elseif OperTabella="CALL_DEL" then  %>
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Rimuovi; Rimuovi;localFun('submit','0');S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div></div>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
		</div>	
   <%end if %>
   
		<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
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
