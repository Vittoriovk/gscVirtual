<%
  NomePagina="CaratteristicaDettaglioModifica.asp"
  titolo="Template"
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

<%
  NameLoaded= ""
  NameLoaded= NameLoaded & "DescAnagCaratteristica,TE" 
 
   FirstLoad=(Request("CallingPage")<>NomePagina)
   IdAnagCaratteristica=0
   if FirstLoad then 
      PaginaReturn          = getCurrentValueFor("PaginaReturn")
      IdAnagServizio        = getCurrentValueFor("IdAnagServizio")
      IdAnagCaratteristica  = "0" & getCurrentValueFor("IdAnagCaratteristica")
      OperTabella           = Session("swap_OperTabella")
   else
      PaginaReturn          = getValueOfDic(Pagedic,"PaginaReturn")
      OperTabella           = getValueOfDic(Pagedic,"OperTabella")
      IdAnagServizio        = getValueOfDic(Pagedic,"IdAnagServizio")
      IdAnagCaratteristica  = "0" & getValueOfDic(Pagedic,"IdAnagCaratteristica")
   end if 
  IdAnagCaratteristica = cdbl(IdAnagCaratteristica)
  FlagFormazione = false 
  if IdAnagServizio="FORMAZ" then 
     FlagFormazione = true
  end if 
  
  'inizio elaborazione pagina

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
  
   xx=SetDiz(DizDatabase,"IdAnagCaratteristica",0)
   xx=SetDiz(DizDatabase,"IdAnagServizio","")
   xx=SetDiz(DizDatabase,"DescAnagCaratteristica","")
   xx=SetDiz(DizDatabase,"FlagModificabile",0)
   xx=SetDiz(DizDatabase,"KeyRicerca","")
   xx=SetDiz(DizDatabase,"FlagFAD",0)
   xx=SetDiz(DizDatabase,"FlagAula",0)
   xx=SetDiz(DizDatabase,"FlagPratica",0)
   xx=SetDiz(DizDatabase,"DurataFAD",0)
   xx=SetDiz(DizDatabase,"DurataAula",0)
   xx=SetDiz(DizDatabase,"DurataPratica",0)
   
  'recupero i dati 
  if cdbl(IdAnagCaratteristica)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  AnagCaratteristica "
	  MySql = MySql & " Where IdAnagCaratteristica=" & IdAnagCaratteristica
	  xx=GetInfoRecordset(DizDatabase,MySql)
  end if 
  
  if OperTabella="CALL_DEL" then 
     SoloLettura=true
  end if 
  'inserisco il fornitore 
  descD      = Request("DescAnagCaratteristica0")
  KeyRicerca = Request("KeyRicerca0")
  
  if CheckTimePageLoad()=false then 
     Oper=""
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_INS" then 
    if FlagFormazione = true then 
      FlagFAD       = cdbl("0" & Request("FlagFAD0"))
      FlagAula      = cdbl("0" & Request("FlagAula0"))
	  FlagPratica   = cdbl("0" & Request("FlagPratica0"))
      DurataFad     = cdbl("0" & Request("DurataFad0"))
	  DurataAula    = cdbl("0" & Request("DurataAula0"))
	  DurataPratica = cdbl("0" & Request("DurataPratica0"))
	else 
      FlagFAD       = 0
      FlagAula      = 0
	  FlagPratica   = 0
      DurataFad     = 0
	  DurataAula    = 0
	  DurataPratica = 0
	end if
	
    Session("TimeStamp")=TimePage
	KK="0"
	MyQ = "" 
	MyQ = MyQ & " INSERT INTO AnagCaratteristica (IdAnagServizio,DescAnagCaratteristica,FlagModificabile,KeyRicerca"
	MyQ = MyQ & " ,FlagAula,FlagFAD,FlagPratica"
	MyQ = MyQ & " ,DurataAula,durataFAD,durataPratica) " 
	MyQ = MyQ & " values ('" & IdAnagServizio & "','" & apici(descD) & "',1,'" & apici(KeyRicerca) & "'"
	MyQ = MyQ & " ," & NumForDb(FlagAula) & "," & NumForDb(FlagFAD) & "," &  NumForDb(FlagPratica)  
	MyQ = MyQ & " ," & NumForDb(DurataAula) & "," & NumForDb(DurataFAD) & "," &  NumForDb(DurataPratica) & ")" 

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdAnagCaratteristica)>0 then 
    if FlagFormazione = true then 
      FlagFAD       = cdbl("0" & Request("FlagFAD0"))
      FlagAula      = cdbl("0" & Request("FlagAula0"))
	  FlagPratica   = cdbl("0" & Request("FlagPratica0"))
      DurataFad     = cdbl("0" & Request("DurataFad0"))
	  DurataAula    = cdbl("0" & Request("DurataAula0"))
	  DurataPratica = cdbl("0" & Request("DurataPratica0"))
	end if   
	MyQ = "" 
	MyQ = MyQ & " Update AnagCaratteristica Set "
	MyQ = MyQ & " DescAnagCaratteristica = '" & apici(descD) & "'"
	MyQ = MyQ & ",KeyRicerca ='" & apici(KeyRicerca) & "'"
	if FlagFormazione = true then 
	   MyQ = MyQ & ",FlagFAD = "       & NumForDb(FlagFad)
       MyQ = MyQ & ",FlagAula = "      & NumForDb(FlagAula)
	   MyQ = MyQ & ",FlagPratica = "   & NumForDb(FlagPratica)
	   MyQ = MyQ & ",DurataFad = "     & NumForDb(DurataFad)
	   MyQ = MyQ & ",DurataAula = "    & NumForDb(DurataAula)
	   MyQ = MyQ & ",DurataPratica = " & NumForDb(DurataPratica)
	end if  
	MyQ = MyQ & " Where IdAnagCaratteristica = " & IdAnagCaratteristica

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
  end if

  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdAnagCaratteristica)>0 then 
     MsgErrore = VerificaDel("AnagCaratteristica",IdAnagCaratteristica) 
	 if MsgErrore = "" then   
		MyQ = "" 
		MyQ = MyQ & " Delete from AnagCaratteristica "
		MyQ = MyQ & " Where IdAnagCaratteristica = " & IdAnagCaratteristica

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
  
  xx=setValueOfDic(Pagedic,"IdAnagServizio"        ,IdAnagServizio)
  xx=setValueOfDic(Pagedic,"IdAnagCaratteristica"  ,IdAnagCaratteristica)
  xx=setValueOfDic(Pagedic,"OperTabella"           ,OperTabella)
  xx=setValueOfDic(Pagedic,"PaginaReturn"          ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescLoaded="0" 
  DescAnagServizio = LeggiCampo("select * from AnagServizio where IdAnagServizio='" & IdAnagServizio & "'","DescAnagServizio")  
  
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
				<div class="col-11"><h3>Gestione Template:</b> <%=DescPageOper%> </h3>
				</div>
			</div>
	        <div class="row">
	           <div class="col-1">
	           </div>
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("Servizio")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescAnagServizio%>" >
               </div>	
			</div>
			
			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-6">
                  <div class="form-group ">
				     <%
					 kk="DescAnagCaratteristica" 
					 xx=ShowLabel("Descrizione Template")
					 NameLoaded= NameLoaded & kk & ",TE;" 
					 
					 %>
					 <input type="text" class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
					 value="<%=GetDiz(DizDatabase,"DescAnagCaratteristica") %>" >
                  </div>
				</div>  
            </div>
			<div class="row">
			   <div class="col-1">
			   </div>
			   <div class="col-6">
                  <div class="form-group ">
				     <%
					 kk="KeyRicerca" 
					 xx=ShowLabel("Chiavi di ricerca")
					 
					 %>
					 <input type="text" class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
					 value="<%=GetDiz(DizDatabase,"KeyRicerca") %>" >
                  </div>
				</div>  
            </div>			
			<%if flagFormazione = true then %>
		
			<div class="row">
			   <div class="col-1">
			   </div>				
			   <div class="col-1">
			   <%xx=ShowLabel("FAD")
			   response.write "<br>"
			   checked=""
			   If Cdbl(GetDiz(DizDatabase,"FlagFad"))=1 then 
			      checked = " checked "
			   end if 
			   %>
			   <input id="FlagFAD0" <%=checked%> name="FlagFAD0" type="checkbox" value = "1" class="big-checkbox">
			   </div>
			    <div class="col-1">
                  <div class="form-group ">
				     <%
					 kk="DurataFad" 
					 xx=ShowLabel("Numero Ore")
					 NameLoaded= NameLoaded & kk & ",INZ;" 
					 
					 %>
					 <input type="text" class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
					 value="<%=GetDiz(DizDatabase,"DurataFad") %>" >
                  </div>
				</div> 
			</div>
			<div class="row">
			   <div class="col-1">
			   </div>							
			   <div class="col-1">
			   <%xx=ShowLabel("Aula")
			   response.write "<br>"
			   checked=""
			   If Cdbl(GetDiz(DizDatabase,"FlagAula"))=1 then 
			      checked = " checked "
			   end if 
			   %>
			   <input id="FlagAula0" <%=checked%> name="FlagAula0" type="checkbox" value = "1" class="big-checkbox">
			   </div>
			    <div class="col-1">
                  <div class="form-group ">
				     <%
					 kk="DurataAula" 
					 xx=ShowLabel("Numero Ore")
					 NameLoaded= NameLoaded & kk & ",INZ;" 
					 
					 %>
					 <input type="text" class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
					 value="<%=GetDiz(DizDatabase,"DurataAula") %>" >
                  </div>
				</div>  
			</DIV>
			<div class="row">
			   <div class="col-1">
			   </div>							
			   <div class="col-1">
			   <%xx=ShowLabel("Pratica")
			   response.write "<br>"
			   checked=""
			   If Cdbl(GetDiz(DizDatabase,"FlagPratica"))=1 then 
			      checked = " checked "
			   end if 
			   %>
			   <input id="FlagPratica0" <%=checked%> name="FlagPratica0" type="checkbox" value = "1" class="big-checkbox">
			   </div>
			    <div class="col-1">
                  <div class="form-group ">
				     <%
					 kk="DurataPratica" 
					 xx=ShowLabel("Numero Ore")
					 NameLoaded= NameLoaded & kk & ",INZ;" 
					 
					 %>
					 <input type="text" class="form-control" Id="<%=KK%>0" name="<%=KK%>0" 
					 value="<%=GetDiz(DizDatabase,"DurataPratica") %>" >
                  </div>
				</div>  
			</DIV>			
			<%end if %>
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
