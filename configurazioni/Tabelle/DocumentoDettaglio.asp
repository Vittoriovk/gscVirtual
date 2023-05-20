<%
  NomePagina="DocumentoDettaglio.asp"
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
  NameLoaded= NameLoaded & "DescDocumento,TE" 
  NameLoaded= NameLoaded & ";GiorniPreavvisoScadenza,IN" 
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdDocumento=0
  if FirstLoad then 
	 IdDocumento   = "0" & Session("swap_IdDocumento")
	 if Cdbl(IdDocumento)=0 then 
		IdDocumento = cdbl("0" & getValueOfDic(Pagedic,"IdDocumento"))
	 end if 
	 OperTabella   = Session("swap_OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	 if PaginaReturn="" then 
		PaginaReturn = Session("swap_PaginaReturn")
	 end if 
  else
	 IdDocumento   = "0" & getValueOfDic(Pagedic,"IdDocumento")
	 OperTabella   = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdDocumento = cdbl(IdDocumento)
  'inizio elaborazione pagina

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
  
   xx=SetDiz(DizDatabase,"IdDocumento",0)
   xx=SetDiz(DizDatabase,"DescDocumento","")
   xx=SetDiz(DizDatabase,"FlagScadenza","1")
   xx=SetDiz(DizDatabase,"GiorniPreavvisoScadenza",30)
  
  
  'recupero i dati 
  if cdbl(IdDocumento)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  Documento "
	  MySql = MySql & " Where IdDocumento=" & IdDocumento
	  xx=GetInfoRecordset(DizDatabase,MySql)
  end if 
  
  if OperTabella="CALL_DEL" then 
     SoloLettura=true
  end if 
  'inserisco il fornitore 
  descD  = Request("DescDocumento0")
  flagS  = 1
  giorni = 0
  if Request("FlagScadenza")="NO" then 
     flagS  = 0
	 giorni = 0
  else 
     giorni = Request("GiorniPreavvisoScadenza0")
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_INS" then 
  
    Session("TimeStamp")=TimePage
	KK="0"
	MyQ = "" 
	MyQ = MyQ & " INSERT INTO Documento (DescDocumento,FlagScadenza,GiorniPreavvisoScadenza) " 
	MyQ = MyQ & " values ('" & apici(descD) & "'," & flagS & "," & giorni & ")" 

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdDocumento)>0 then 
	MyQ = "" 
	MyQ = MyQ & " Update Documento "
	MyQ = MyQ & " Set DescDocumento = '" & apici(descD) & "'"
	MyQ = MyQ & ",FlagScadenza = " & flagS
	MyQ = MyQ & ",GiorniPreavvisoScadenza = " & giorni 
	MyQ = MyQ & " Where IdDocumento = " & IdDocumento
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
  end if
  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdDocumento)>0 then 
     MsgErrore = VerificaDel("Documento",IdDocumento) 
	 if MsgErrore = "" then   
		MyQ = "" 
		MyQ = MyQ & " Delete from Documento "
		MyQ = MyQ & " Where IdDocumento = " & IdDocumento

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
  xx=setValueOfDic(Pagedic,"IdDocumento"  ,IdDocumento)
  xx=setValueOfDic(Pagedic,"OperTabella"  ,OperTabella)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  checkedSi=""
  checkedNo=""
  if GetDiz(DizDatabase,"FlagScadenza") = "1" then 
     checkedSi = " checked "
  else
     checkedNo = " checked "
  end if 
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
				<div class="col-11"><h3>Gestione Documento:</b> <%=DescPageOper%> </h3>
				</div>
			</div>
			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-6">
                  <div class="form-group ">
				     <%
					 kk="DescDocumento" 
					 xx=ShowLabel("Descrizione Documento")
					 NameLoaded= NameLoaded & kk & ",TE;" 
					 
					 %>
					 <input type="text" class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
					 value="<%=GetDiz(DizDatabase,"DescDocumento") %>" >
                  </div>
				</div>  
            </div>

   <div class="row">
      <div class="col-1">
      </div> 
      <div class="col-1">
         <%xx=ShowLabel("Scadenza")%>
      </div> 
   </div>
   <div class="row">
      <div class="col-1">
      </div> 
	  <% if SoloLettura=false or GetDiz(DizDatabase,"FlagScadenza") = "1" then %> 
	  <div class="col-1">
         <div class="form-check-inline">
              <input name="FlagScadenza" value="SI" type="radio" id="FlagScadenzaSI" <%=checkedSi%>>
         </div>
		 <span class="font-weight-bold">SI</span>
	  </div>
	  <%end if %>
  
	<% if SoloLettura=false or GetDiz(DizDatabase,"FlagScadenza") = "0" then %> 
	  <div class="col-1">
         <div class="form-check-inline">
              <input name="FlagScadenza" value="NO" type="radio" id="FlagScadenzaNO" <%=checkedNo%>>
         </div>
         <span class="font-weight-bold">NO</span>		 
	  </div>	  
	   <%end if %>
   </div>   
			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-2">
                  <div class="form-group ">
				     <%
					 kk="GiorniPreavvisoScadenza" 
					 xx=ShowLabel("Giorni Preavviso scadenza")
					 NameLoaded= NameLoaded & kk & ",TE;" 
					 
					 %>
					 <input type="text" class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
					 value="<%=GetDiz(DizDatabase,"GiorniPreavvisoScadenza") %>" >
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
