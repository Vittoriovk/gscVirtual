<%
  NomePagina="SubRamoDettaglio.asp"
  titolo="Caratteristica Ramo"
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
  NameLoaded= NameLoaded & "DescRamo,TE" 
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdSubRamo=0
  if FirstLoad then 
	 IdSubRamo   = "0" & Session("swap_IdSubRamo")
	 if Cdbl(IdSubRamo)=0 then 
		IdSubRamo = cdbl("0" & getValueOfDic(Pagedic,"IdSubRamo"))
	 end if 
	 OperTabella   = Session("swap_OperTabella")
	 PaginaReturn    = getCurrentValueFor("PaginaReturn")
  else
	 IdSubRamo   = "0" & getValueOfDic(Pagedic,"IdSubRamo")
	 OperTabella   = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdSubRamo = cdbl(IdSubRamo)
  'inizio elaborazione pagina

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
  
   xx=SetDiz(DizDatabase,"IdSubRamo",0)
   xx=SetDiz(DizDatabase,"DescSubRamo","")
  
  'recupero i dati 
  if cdbl(IdSubRamo)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  SubRamo "
	  MySql = MySql & " Where IdSubRamo=" & IdSubRamo
	  xx=GetInfoRecordset(DizDatabase,MySql)
  end if 
  
  if OperTabella="CALL_DEL" then 
     SoloLettura=true
  end if 
 
  descD  = Request("DescSubRamo0")
  IdRamo = Request("IdRamo0")
  
  if Oper=ucase("update") and OperTabella="CALL_INS" then 
  
    Session("TimeStamp")=TimePage
	KK="0"
	MyQ = "" 
	MyQ = MyQ & " INSERT INTO SubRamo (DescSubRamo,IdRamo) " 
	MyQ = MyQ & " values ('" & apici(descD) & "'," & IdRamo & ")" 

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdSubRamo)>0 then 
	MyQ = "" 
	MyQ = MyQ & " Update SubRamo "
	MyQ = MyQ & " Set DescSubRamo = '" & apici(descD) & "'"
	MyQ = MyQ & ",IdRamo = " & IdRamo
	MyQ = MyQ & " Where IdSubRamo = " & IdSubRamo

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
  end if

  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdSubRamo)>0 then 
     MsgErrore = VerificaDel("SubRamo",IdSubRamo) 
	 if MsgErrore = "" then   
		MyQ = "" 
		MyQ = MyQ & " Delete from SubRamo "
		MyQ = MyQ & " Where IdSubRamo = " & IdSubRamo

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
  xx=setValueOfDic(Pagedic,"IdSubRamo"  ,IdSubRamo)
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
				<div class="col-11"><h3>Gestione Caratteristica Ramo:</b> <%=DescPageOper%> </h3>
				</div>
			</div>

			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-6">
                  <div class="form-group ">
				     <%
					 kk="DescSubRamo" 
					 xx=ShowLabel("Descrizione Caratteristica Ramo")
					 NameLoaded= NameLoaded & kk & ",TE;" 
					 
					 %>
					 <input type="text" class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
					 value="<%=GetDiz(DizDatabase,"DescSubRamo") %>" >
                  </div>
				</div>  
            </div>
			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-6">
                  <div class="form-group ">
				     <%
					 kk="IdRamo" 
					 xx=ShowLabel("Ramo di riferimento")
					 NameLoaded= NameLoaded & kk & ",LI;" 
					 q = ""
					 q = q & " select * from Ramo "
					 q = q & " order by descRamo"
					 
					 stdClass="class='form-control form-control-sm'"
					 response.write ListaDbChangeCompleta(q,"IdRamo0",GetDiz(DizDatabase,"IdRamo") ,"IdRamo","DescRamo" ,tt,"","","","","",stdClass)
					 
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
