<%
  NomePagina="TratFiscDettaglio.asp"
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
  NameLoaded= NameLoaded & "DescTrattamentoFiscale,TE" 
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdTrattamentoFiscale=0
  if FirstLoad then 
	 IdTrattamentoFiscale   = "0" & Session("swap_IdTrattamentoFiscale")
	 if Cdbl(IdTrattamentoFiscale)=0 then 
		IdTrattamentoFiscale = cdbl("0" & getValueOfDic(Pagedic,"IdTrattamentoFiscale"))
	 end if 
	 OperTabella   = Session("swap_OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	 if PaginaReturn="" then 
		PaginaReturn = Session("swap_PaginaReturn")
	 end if 
  else
	 IdTrattamentoFiscale   = "0" & getValueOfDic(Pagedic,"IdTrattamentoFiscale")
	 OperTabella   = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdTrattamentoFiscale = cdbl(IdTrattamentoFiscale)
  'inizio elaborazione pagina

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
  
   xx=SetDiz(DizDatabase,"IdTrattamentoFiscale",0)
   xx=SetDiz(DizDatabase,"DescTrattamentoFiscale","")
  
  'recupero i dati 
  if cdbl(IdTrattamentoFiscale)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  TrattamentoFiscale "
	  MySql = MySql & " Where IdTrattamentoFiscale=" & IdTrattamentoFiscale
	  xx=GetInfoRecordset(DizDatabase,MySql)
  end if 
  
  if OperTabella="CALL_DEL" then 
     SoloLettura=true
  end if 
  'inserisco il fornitore 
  descD  = Request("DescTrattamentoFiscale0")
  
  if Oper=ucase("update") and OperTabella="CALL_INS" then 
  
    Session("TimeStamp")=TimePage
	KK="0"
	MyQ = "" 
	MyQ = MyQ & " INSERT INTO TrattamentoFiscale (DescTrattamentoFiscale) " 
	MyQ = MyQ & " values ('" & apici(descD) & "')" 

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdTrattamentoFiscale)>0 then 
	MyQ = "" 
	MyQ = MyQ & " Update TrattamentoFiscale "
	MyQ = MyQ & " Set DescTrattamentoFiscale = '" & apici(descD) & "'"
	MyQ = MyQ & " Where IdTrattamentoFiscale = " & IdTrattamentoFiscale

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
  end if

  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdTrattamentoFiscale)>0 then 
     MsgErrore = VerificaDel("TrattamentoFiscale",IdTrattamentoFiscale) 
	 if MsgErrore = "" then   
		MyQ = "" 
		MyQ = MyQ & " Delete from TrattamentoFiscale "
		MyQ = MyQ & " Where IdTrattamentoFiscale = " & IdTrattamentoFiscale

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
  xx=setValueOfDic(Pagedic,"IdTrattamentoFiscale"  ,IdTrattamentoFiscale)
  xx=setValueOfDic(Pagedic,"OperTabella"  ,OperTabella)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescLoaded="0"  
  %>

<% 
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
			<%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h3>Gestione Trattamento Fiscale:</b> <%=DescPageOper%> </h3>
				</div>
			</div>

 			<div class="row">
			   <div class="col-2">
			   </div>
			   <div class="col-8">
                  <div class="form-group ">
				     <%xx=ShowLabel("Descrizione Tratt.Fiscale")
					 ao_nid = "DescTrattamentoFiscale0"
					 %>
					 <input type="text" name="<%=ao_nid%>" id="<%=ao_nid%>" class="form-control input-sm" value="<%=GetDiz(DizDatabase,"DescTrattamentoFiscale")%>" >
                  </div>		
			   </div>
			</div>

   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->		
 
   
   <%if SoloLettura=false then%>
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
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
