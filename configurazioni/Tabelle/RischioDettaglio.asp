<%
  NomePagina="RischioDettaglio.asp"
  titolo="Gestione Del Rischio"
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
function reload()
{
	ImpostaValoreDi("Oper","reload");
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
  NameLoaded= NameLoaded & "DescRischio,TE;IdRamo,LI" 
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdRischio=0
  if FirstLoad then 
	 IdRischio   = "0" & Session("swap_IdRischio")
	 if Cdbl(IdRischio)=0 then 
		IdRischio = cdbl("0" & getValueOfDic(Pagedic,"IdRischio"))
	 end if 
	 OperTabella   = Session("swap_OperTabella")
	 PaginaReturn    = getCurrentValueFor("PaginaReturn")
  else
	 IdRischio   = "0" & getValueOfDic(Pagedic,"IdRischio")
	 OperTabella   = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdRischio = cdbl(IdRischio)
  'inizio elaborazione pagina

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
  
   xx=SetDiz(DizDatabase,"IdRischio",0)
   xx=SetDiz(DizDatabase,"DescRischio","")
   xx=SetDiz(DizDatabase,"IdRamo",0)
   xx=SetDiz(DizDatabase,"IdAnagCaratteristica","")
  
  'recupero i dati 
  if cdbl(IdRischio)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  Rischio "
	  MySql = MySql & " Where IdRischio=" & IdRischio
	  xx=GetInfoRecordset(DizDatabase,MySql)
  end if 
  
  if OperTabella="CALL_DEL" then 
     SoloLettura=true
  end if 

  if FirstLoad=false then 
     descRischio          = GetDiz(DizDatabase,"DescRischio")
     IdRamo               = GetDiz(DizDatabase,"IdRamo")
     IdAnagCaratteristica = GetDiz(DizDatabase,"IdAnagCaratteristica")
  
     if oper<>"" then   
        descRischio          = Request("DescRischio0")
        IdRamo               = cdbl("0" & Cdbl(request("IdRamo0")))
        IdAnagCaratteristica = Request("IdAnagCaratteristica0")
	 end if 
  end if 
  if Oper=ucase("update") and OperTabella="CALL_INS" then 
  
    Session("TimeStamp")=TimePage
	KK="0"
	MyQ = "" 
	MyQ = MyQ & " INSERT INTO Rischio (DescRischio,IdRamo,IdAnagCaratteristica) " 
	MyQ = MyQ & " values ('" & apici(descRischio) & "'," & IdRamo & ",'" & apici(IdAnagCaratteristica) & "')" 

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdRischio)>0 then 
	MyQ = "" 
	MyQ = MyQ & " Update Rischio "
	MyQ = MyQ & " Set DescRischio = '" & apici(descRischio) & "'"
	MyQ = MyQ & ",IdRamo = " & IdRamo
	MyQ = MyQ & ",IdAnagCaratteristica='" & apici(IdAnagCaratteristica) & "'"
	MyQ = MyQ & " Where IdRischio = " & IdRischio

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
  end if

  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdRischio)>0 then 
     MsgErrore = VerificaDel("Rischio",IdRischio) 
	 if MsgErrore = "" then   
		MyQ = "" 
		MyQ = MyQ & " Delete from Rischio "
		MyQ = MyQ & " Where IdRischio = " & IdRischio

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
  xx=setValueOfDic(Pagedic,"IdRischio"    ,IdRischio)
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
				<div class="col-11"><h3>Gestione Rischio:</b> <%=DescPageOper%> </h3>
				</div>
			</div>

			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-6">
                  <div class="form-group ">
				     <%
					 kk="DescRischio" 
					 xx=ShowLabel("Descrizione Rischio")
					 NameLoaded= NameLoaded & kk & ",TE;" 
					 
					 %>
					 <input type="text" class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
					 value="<%=DescRischio %>" >
                  </div>
				</div>  
            </div>
			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-6">
                  <div class="form-group ">
				     <%
					 tt=1
					 kk="IdRamo" 
					 xx=ShowLabel("Ramo di riferimento")
					 NameLoaded= NameLoaded & kk & ",LI;" 
					 q = ""
					 q = q & " select * from Ramo "
					 q = q & " order by descRamo"
					 stdClass="class='form-control form-control-sm'"
					 response.write ListaDbChangeCompleta(q,"IdRamo0",IdRamo ,"IdRamo","DescRamo" ,tt,"reload()","","","","",stdClass)
					 
					 %>
                  </div>
				</div>  
            </div>
			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-6">
                  <div class="form-group ">
				     <%
					 qAs=""
					 if cdbl(IdRamo)>0 then 
					    IdAnagRamo=LeggiCampo("select * from Ramo Where IdRamo=" & IdRamo,"IdAnagRamo")
						if IdAnagRamo<>"" then 
						qAS =  "SELECT IdAnagServizio From AnagServizio where IdAnagRamo='" & IdAnagRamo & "'"
						end if 
						
                     end if  					 
					 kk="IdAnagCaratteristica" 
					 xx=ShowLabel("Template di riferimento")
					 NameLoaded= NameLoaded & kk & ",LI;" 
					 q = ""
					 q = q & " select * from AnagCaratteristica "
					 if qAS<>"" then 
					    q = q & " Where IdAnagServizio in (" & qAS & ")"  
					 end if 
					 q = q & " order by descAnagCaratteristica"
					 stdClass="class='form-control form-control-sm'"
					 response.write ListaDbChangeCompleta(q,"IdAnagCaratteristica0",IdAnagCaratteristica ,"IdAnagCaratteristica","DescAnagCaratteristica" ,tt,"","","","","",stdClass)
					 
					 cs=LeggiCampo(q,"IdAnagCaratteristica")
					 if cs="" and OperTabella<>"CALL_DEL" then 
					    MsgErrore="Template non definito : non e' possibile procedere"
					 end if
					 %>
                  </div>
				</div>  
            </div>			
<br>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->		
 
   
   <%if SoloLettura=false and cs<>"" then%>
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
