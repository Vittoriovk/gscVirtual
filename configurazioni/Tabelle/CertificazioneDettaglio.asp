<%
  NomePagina="CertificazioneDettaglio.asp"
  titolo="Menu Supervisor - Dashboard"
  'forzo il controllo al profilo 
  default_check_profile="SUPERV"  
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
  NameLoaded= NameLoaded & "DescBreveCertificazione,TE;DescEstesaCertificazione,TE;PercRiduzioneCauzione,FLQ" 
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  
  IdCertificazione = 0
  if FirstLoad then 
     IdCertificazione = getCurrentValueFor("IdCertificazione")
     OperTabella      = getCurrentValueFor("OperTabella")
     PaginaReturn     = getCurrentValueFor("PaginaReturn") 
  else
     IdCertificazione = getValueOfDic(Pagedic,"IdCertificazione")
     OperTabella      = getValueOfDic(Pagedic,"OperTabella")
     PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  
  IdCertificazione = cdbl(IdCertificazione)
  'inizio elaborazione pagina

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
  
   xx=SetDiz(DizDatabase,"IdCertificazione",0)
   xx=SetDiz(DizDatabase,"DescCertificazione","")
  
   if OperTabella="CALL_DEL" then 
      SoloLettura=true
   end if 
   'inserisco il fornitore 
   descD  = Request("DescBreveCertificazione0")
   descE  = Request("DescEstesaCertificazione0")
   PercR  = Request("PercRiduzioneCauzione0")
   IdDoc  = Request("IdDocumento0")
  
   if Oper=ucase("update") and OperTabella="CALL_INS" then 
  
     Session("TimeStamp")=TimePage
     KK="0"
     MyQ = "" 
     MyQ = MyQ & " INSERT INTO Certificazione (DescBreveCertificazione,DescEstesaCertificazione,PercRiduzioneCauzione,IdDocumento) " 
     MyQ = MyQ & " values ('" & apici(descD) & "','" &  apici(descE) & "'," & NumForDb(PercR) & "," & NumForDb(IdDoc) & ")" 

     ConnMsde.execute MyQ 
     If Err.Number <> 0 Then 
         MsgErrore = ErroreDb(Err.description)
      else 
         response.redirect virtualpath & PaginaReturn
      End If
   end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdCertificazione)>0 then 
	MyQ = "" 
	MyQ = MyQ & " Update Certificazione set "
	MyQ = MyQ & "  DescBreveCertificazione = '" & apici(descD) & "'"
	MyQ = MyQ & " ,DescEstesaCertificazione = '" & apici(descE) & "'"
	MyQ = MyQ & " ,PercRiduzioneCauzione = " & NumforDb(PercR)
	MyQ = MyQ & " ,IdDocumento = " & NumforDb(IdDoc)
	MyQ = MyQ & " Where IdCertificazione = " & IdCertificazione

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
  end if

  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdCertificazione)>0 then 
     MsgErrore = VerificaDel("Certificazione",IdCertificazione) 
	 if MsgErrore = "" then   
		MyQ = "" 
		MyQ = MyQ & " Delete from Certificazione "
		MyQ = MyQ & " Where IdCertificazione = " & IdCertificazione

		ConnMsde.execute MyQ 
		If Err.Number <> 0 Then 
			MsgErrore = ErroreDb(Err.description)
		else 
		   response.redirect virtualpath & PaginaReturn
		End If	
	end if 
  end if  

   'recupero i dati 
  if cdbl(IdCertificazione)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  Certificazione "
	  MySql = MySql & " Where IdCertificazione=" & IdCertificazione
	  xx=GetInfoRecordset(DizDatabase,MySql)
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
  xx=setValueOfDic(Pagedic,"IdCertificazione"  ,IdCertificazione)
  xx=setValueOfDic(Pagedic,"OperTabella"       ,OperTabella)
  xx=setValueOfDic(Pagedic,"PaginaReturn"      ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescLoaded="0"  
  flagReadonly=""
  if SoloLettura then 
     flagReadonly=" readonly "
  end if 
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
				<div class="col-11"><h3>Gestione Certificazione:</b> <%=DescPageOper%> </h3>
				</div>
			</div>

			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-9">
				  <div class="form-group ">
					 <%
					 kk="DescBreveCertificazione" 
					 xx=ShowLabel("Descrizione")
					 
					 %>
                     <input type="text" rows="3" <%=flagReadonly%> class="form-control" Id="<%=KK%>0" name="<%=KK%>0"
					 value ="<%=GetDiz(DizDatabase,"DescBreveCertificazione")%>" >
				  </div>		
			   </div>
			</div>	 
			
			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-9">
				  <div class="form-group ">
					 <%
					 kk="DescEstesaCertificazione" 
					 xx=ShowLabel("Descrizione Dettagliata")
					 
					 %>
                     <textarea rows="3" <%=flagReadonly%> class="form-control" Id="<%=KK%>0" name="<%=KK%>0"
					 ><%=GetDiz(DizDatabase,"DescEstesaCertificazione") %></textarea>
				  </div>		
			   </div>
			</div>	 

			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-1">
				  <div class="form-group ">
					 <%
					 kk="PercRiduzioneCauzione" 
					 xx=ShowLabel("% riduzione")
					 
					 %>
                     <input type="text" rows="3" <%=flagReadonly%> class="form-control" Id="<%=KK%>0" name="<%=KK%>0"
					 value ="<%=GetDiz(DizDatabase,"PercRiduzioneCauzione")%>" >
				  </div>		
			   </div>
			</div>	 

			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-6">
				  <div class="form-group ">
					 <%
					 kk="IdDocumento" 
					 IdDocumento = GetDiz(DizDatabase,"IdDocumento")
					 xx=ShowLabel("Documento Da Caricare")
					 stdClass="class='form-control form-control-sm'"
					 q = ""
	                 q = q & " Select * from Documento "
					 tt=1
					 if flagReadonly<>"" then 
					    q = q & " and IdDocumento =  " & IdDocumento
					    tt=0
					 end if 
	                 q = q & " order by DescDocumento  "					 
					 response.write ListaDbChangeCompleta(q,"IdDocumento0",IdDocumento ,"IdDocumento","DescDocumento" ,tt,"","","","","",stdClass)
					 %>
				  </div>		
			   </div>
			</div>
<br>			
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
