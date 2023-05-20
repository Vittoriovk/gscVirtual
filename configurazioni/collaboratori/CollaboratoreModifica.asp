<%
  NomePagina="CollaboratoreModifica.asp"
  titolo="Collaboratori per Azienda"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/functionCf.asp"-->
<!--#include virtual="/gscVirtual/common/functionCfScript.asp"-->
<!--#include virtual="/gscVirtual/common/functionDataList.asp"-->
<!--#include virtual="/gscVirtual/common/functionDataListScript.asp"-->
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
<!--#include virtual="/gscVirtual/js/functionTable.js"-->
<!--#include virtual="/gscVirtual/js/functionLocalita.js"-->
<script language="JavaScript">

function calcolaCFColl()
{
   var nome=$("#Cognome0").val();
   var cogn=$("#Nome0").val();
   var sess=$("#IdSesso0").val();;
  
   var dtna=$("#DataNascita0").val();
   var stat=$("#StatoNascita0").val();
   var comu=$("#ComuneNascita0").val();
   var prov=$("#ProvinciaNascita0").val();
   var cf = calcolaCF(nome,cogn,sess,dtna,stat,prov,comu);
   if (cf.length>0)
      $("#CodiceFiscale0").val(cf);
}
function calcolaCFAmmiPrep(id)
{
   var nome=$("#Cognome" + id + "0").val();
   var cogn=$("#Nome" + id + "0").val();
   var sess=$("#IdSesso" + id + "0").val();;
  
   var dtna=$("#DataNascita" + id + "0").val();
   var stat=$("#StatoNascita" + id + "0").val();
   var comu=$("#ComuneNascita" + id + "0").val();
   var prov=$("#ProvinciaNascita" + id + "0").val();
   var cf = calcolaCF(nome,cogn,sess,dtna,stat,prov,comu);
   if (cf.length>0)
      $("#CodiceFiscale" + id + "0").val(cf);
}
function copiaDaAmm()
{
   $("#CognomePreposto0").val($("#CognomeAmministratore0").val());
   $("#NomePreposto0").val($("#NomeAmministratore0").val());
   $("#CodiceFiscalePreposto0").val($("#CodiceFiscaleAmministratore0").val());
   $("#IndirizzoPreposto0").val($("#IndirizzoAmministratore0").val());
   $("#CittaPreposto0").val($("#CittaAmministratore0").val());
   $("#ProvinciaPreposto0").val($("#ProvinciaAmministratore0").val());
   $("#CapPreposto0").val($("#CapAmministratore0").val());
   $("#StatoNascitaPreposto0").val($("#StatoNascitaAmministratore0").val());
   $("#ComuneNascitaPreposto0").val($("#ComuneNascitaAmministratore0").val());
   $("#ProvinciaNascitaPreposto0").val($("#ProvinciaNascitaAmministratore0").val());
   $("#DataNascitaPreposto0").val($("#DataNascitaAmministratore0").val());

}

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

<!--#include virtual="/gscVirtual/modelli/FunctionAccount.asp"-->
  
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

  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdAccount=0
  IdCollaboratore=0
  if FirstLoad then 
	 IdCollaboratore   = "0" & Session("swap_IdCollaboratore")
	 if Cdbl(IdCollaboratore)=0 then 
		IdCollaboratore = cdbl("0" & getValueOfDic(Pagedic,"IdCollaboratore"))
	 end if   
	 IdAccount   = "0" & Session("swap_IdAccount")
	 if Cdbl(IdAccount)=0 then 
		IdAccount = cdbl("0" & getValueOfDic(Pagedic,"IdAccount"))
	 end if 
	 IdAzienda   = "0" & Session("swap_IdAzienda")
	 if Cdbl(IdAzienda)=0 then 
		IdAzienda = cdbl("0" & getValueOfDic(Pagedic,"IdAzienda"))
	 end if	 
	 OperAmmesse   = Session("swap_OperAmmesse")
	 OperTabella   = Session("swap_OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	 if PaginaReturn="" then 
		PaginaReturn = Session("swap_PaginaReturn")
	 end if 
	 if Cdbl(IdCollaboratore)=0 then 
	    IdTipoColl=Session("swap_IdTipoCollaboratore")
	    IdTipoPers=Session("swap_IdPersCollaboratore")
	 end if 

  else
	 IdAccount       = "0" & getValueOfDic(Pagedic,"IdAccount")
	 IdCollaboratore = "0" & getValueOfDic(Pagedic,"IdCollaboratore")
	 OperAmmesse     = getValueOfDic(Pagedic,"OperAmmesse")
	 OperTabella     = getValueOfDic(Pagedic,"OperTabella") 
	 PaginaReturn    = getValueOfDic(Pagedic,"PaginaReturn")
	 if Cdbl(IdCollaboratore)=0 then 
	    IdTipoColl=getValueOfDic(Pagedic,"IdTipoColl")
	    IdTipoPers=getValueOfDic(Pagedic,"IdTipoPers")
	 end if	 
   end if 
   IdCollaboratore = cdbl(IdCollaboratore)
   IdAccount = cdbl(IdAccount)
   if OperAmmesse="" then 
      if IdCollaboratore = 0 then 
         OperAmmesse="CRUD"
      end if 
   end if 

  'sono in inserimento : creo un account fittizio 
  if cdbl(IdAccount)=0 and OperTabella="CALL_INS" then 
     IdAccount=GetTempAccount()
  end if 
  
  'inserisco account 
   if Oper=ucase("update") then 
      Ritorna=false 
	  OperAmmesse="U"
      Session("TimeStamp")=TimePage
      MsgErrore=""
      Cognome    = Request("Cognome0")
      Nome       = Request("Nome0")
	  Nominativo = Request("Denominazione0")
	  if Nominativo = "" then 
	     Nominativo =  trim(trim(Cognome) & " " & Trim(Nome)) 
	  end if 

	  StatoNascita     = Request("StatoNascita0")
	  ComuneNascita    = Request("ComuneNascita0")
	  ProvinciaNascita = Request("ProvinciaNascita0")
	  DataNascita      = Request("DataNascita0")
	  if len(DataNascita)<>10 then 
	     DataNascita=0
	  else 
	     DataNascita=DataStringa(DataNascita)
	  end if 
	  IdSesso       = Request("IdSesso0")
      CodiceFiscale = Request("CodiceFiscale0")
      PartititaIva  = Request("PartitaIva0")
      SezioneRui    = Request("SezioneRui0")
      NumeroRui     = Request("NumeroRui0") 	  
      IdTipoColl    = Request("IdTipoCollaboratore0")
	  if request("checkCrea0")="S" then 
	     FlagGeneraCollaboratore=1
	  else
	     FlagGeneraCollaboratore=0
	  end if 
	  if request("checkTeam0")="S" then 
	     FlagTeamLeader=1
	  else
	     FlagTeamLeader=0
	  end if 
	  IdTeam                = TestNumeroPos("0" & Request("IdTeam0"))
      CognomePreposto       = Request("CognomePreposto0")
      NomePreposto          = Request("NomePreposto0")
      CodiceFiscalePreposto = Request("CodiceFiscalePreposto0")
      PartitaIvaPreposto    = Request("PartitaIvaPreposto0")
      IndirizzoPreposto     = Request("IndirizzoPreposto0")
      CittaPreposto         = Request("CittaPreposto0")
      ProvinciaPreposto     = Request("ProvinciaPreposto0")
      CapPreposto           = Request("CapPreposto0")
	  StatoNascitaPreposto     = Request("StatoNascitaPreposto0")
	  ComuneNascitaPreposto    = Request("ComuneNascitaPreposto0")
	  ProvinciaNascitaPreposto = Request("ProvinciaNascitaPreposto0")
	  DataNascitaPreposto      = Request("DataNascitaPreposto0")
	  if len(DataNascitaPreposto)<>10 then 
	     DataNascitaPreposto=0
	  else 
	     DataNascitaPreposto=DataStringa(DataNascitaPreposto)
	  end if 
	  IdSessoPreposto       = Request("IdSessoPreposto0")
      SezioneRuiPreposto    = Request("SezioneRuiPreposto0")
      NumeroRuiPreposto     = Request("NumeroRuiPreposto0")
      DataIscrizioneRuiPreposto = Request("DataIscrizioneRuiPreposto0")
	  if len(DataIscrizioneRuiPreposto)<>10 then 
	     DataIscrizioneRuiPreposto=0
      else
	     DataIscrizioneRuiPreposto=DataStringa(DataIscrizioneRuiPreposto)
	  end if 

      CognomeAmministratore       = Request("CognomeAmministratore0")
      NomeAmministratore          = Request("NomeAmministratore0")
      CodiceFiscaleAmministratore = Request("CodiceFiscaleAmministratore0")
      PartitaIvaAmministratore    = Request("PartitaIvaAmministratore0")
      IndirizzoAmministratore     = Request("IndirizzoAmministratore0")
      CittaAmministratore         = Request("CittaAmministratore0")
      ProvinciaAmministratore     = Request("ProvinciaAmministratore0")
      CapAmministratore           = Request("CapAmministratore0")
	  StatoNascitaAmministratore     = Request("StatoNascitaAmministratore0")
	  ComuneNascitaAmministratore    = Request("ComuneNascitaAmministratore0")
	  ProvinciaNascitaAmministratore = Request("ProvinciaNascitaAmministratore0")
	  DataNascitaAmministratore      = Request("DataNascitaAmministratore0")
	  if len(DataNascitaAmministratore)<>10 then 
	     DataNascitaAmministratore=0
	  else 
	     DataNascitaAmministratore=DataStringa(DataNascitaAmministratore)
	  end if 

      IdSessoAmministratore = Request("IdSessoAmministratore0")
  
      DataIscrizioneRui     = Request("DataIscrizioneRui0")
	  if len(DataIscrizioneRui)<>10 then 
	     DataIscrizioneRui=0
	  else 
	     DataIscrizioneRui=DataStringa(DataIscrizioneRui)
	  end if 

      if Cdbl(IdCollaboratore)=0 then 
		 if Cdbl(IdAccount)=0 then 
		    MsgErrore="errore di sistema : contattare assistenza"
         else 
		    myUpd = ""
		    myUpd = myUpd & " Update Account Set IdAzienda=" & Session("IdAziendaWork") 
			myUpd = myUpd & ",IdTipoAccount='Coll'"
			myUpd = myUpd & ",FlagAttivo='S',Abilitato=0"
			myUpd = myUpd & ",Nominativo='" & apici(Nominativo) & "'"
			myUpd = myUpd & " Where IdAccount=" & IdAccount
			ConnMsde.execute = MyUpd
			NextL=cdbl(session("LivelloAccount") + 1)
            MyQ = "" 
            MyQ = MyQ & " Insert into Collaboratore (IdAccount,IdAzienda,Denominazione,IdTipoDitta,Livello)"
            MyQ = MyQ & " values (" & IdAccount & "," & Session("IdAziendaWork") & ",'" & Apici(Nominativo) & "'"
            MyQ = MyQ & ",'" & apici(IdTipoPers) & "'," & NextL & ")"
            ConnMsde.execute MyQ 
			'response.write MyQ
            If Err.Number <> 0 Then 
               MsgErrore = ErroreDb(Err.description)
			   IdAccount=0
            else
			   Ritorna=true 
               IdCollaboratore = GetTableIdentity("Collaboratore")    
            end if 
         end if 
      end if 
      'aggiorno Collaboratore 
	  'devo gestire il cambio profilo 
	  IdProfiloProdotto    = getRequestAsNum("IdProfiloProdotto0") 

	  IdProfiloProdottoOld = LeggiCampo("select * from Collaboratore Where IdCollaboratore=" & IdCollaboratore ,"IdProfiloProdotto")
	  IdProfiloProdottoNew = IdProfiloProdotto
	  'cambiato profilo
	  if Cdbl(IdProfiloProdottoNew)<>Cdbl(IdProfiloProdottoOld) then 
	     'devo rimuovere il vecchio 
	     if Cdbl(IdProfiloProdottoOld)<>0 then 
		    v = ""
			v = v & " delete from AccountProfiloProdotto"
			v = v & " Where IdProfiloProdotto = " & IdProfiloProdottoOld 
			v = v & " and IdAccount = " & IdAccount
			connMsde.execute v
		 end if 
		 'devo aggiungere il nuovo
	     if Cdbl(IdProfiloProdottoNew)<>0 then 
		    v = ""
			v = v & " insert into AccountProfiloProdotto ("
			v = v & " IdAccount,IdProfiloProdotto,ValidoDal,ValidoAl)"
			v = v & " values ("
			v = v & IdAccount & "," & IdProfiloProdotto & "," & Dtos() & ",99991231"
			v = v & " )"
		    ConnMsde.execute V 
		 end if 
		 
	  end if 
	  
	  
      MyQ = "" 
      MyQ = MyQ & " update Collaboratore set "
      MyQ = MyQ & " Denominazione = '" & apici(Nominativo) & "'"
	  MyQ = MyQ & ",FlagGeneraCollaboratore=1"
      MyQ = MyQ & ",Cognome = '" & apici(Cognome) & "'"
      MyQ = MyQ & ",Nome = '"    & apici(Nome)    & "'"
	  MyQ = MyQ & ",StatoNascita = '" & apici(StatoNascita) & "'"
	  MyQ = MyQ & ",ComuneNascita = '" & apici(ComuneNascita) & "'"
	  MyQ = MyQ & ",ProvinciaNascita = '" & apici(ProvinciaNascita) & "'"
	  MyQ = MyQ & ",DataNascita = " & DataNascita 
      MyQ = MyQ & ",CodiceFiscale = '" & apici(CodiceFiscale) & "'"
      MyQ = MyQ & ",PartitaIva = '"  & apici(PartititaIva) & "'"
      MyQ = MyQ & ",SezioneRui = '"  & apici(SezioneRui) & "'"
      MyQ = MyQ & ",NumeroRui = '"   & apici(NumeroRui) & "'"
      MyQ = MyQ & ",DataIscrizioneRui = "      & DataIscrizioneRui
      MyQ = MyQ & ",FlagTeamLeader = " & FlagTeamLeader
      MyQ = MyQ & ",IdTeam = "         & IdTeam
      MyQ = MyQ & ",IdTipoCollaboratore = '"   & apici(IdTipoColl) & "'"
      MyQ = MyQ & ",CognomePreposto = '"       & apici(CognomePreposto) & "'"
      MyQ = MyQ & ",NomePreposto = '"          & apici(NomePreposto) & "'"
      MyQ = MyQ & ",CodiceFiscalePreposto = '" & apici(CodiceFiscalePreposto) & "'"
      MyQ = MyQ & ",PartitaIvaPreposto = '"    & apici(PartitaIvaPreposto) & "'"
      MyQ = MyQ & ",IndirizzoPreposto = '"     & apici(IndirizzoPreposto) & "'"
      MyQ = MyQ & ",CittaPreposto = '"         & apici(CittaPreposto) & "'"
      MyQ = MyQ & ",ProvinciaPreposto = '"     & apici(ProvinciaPreposto) & "'"
      MyQ = MyQ & ",CapPreposto = '"           & apici(CapPreposto) & "'"
	  MyQ = MyQ & ",StatoNascitaPreposto = '"  & apici(StatoNascitaPreposto) & "'"
	  MyQ = MyQ & ",ComuneNascitaPreposto = '" & apici(ComuneNascitaPreposto) & "'"
	  MyQ = MyQ & ",ProvinciaNascitaPreposto = '" & apici(ProvinciaNascitaPreposto) & "'"
	  MyQ = MyQ & ",DataNascitaPreposto = "    & DataNascitaPreposto 
      MyQ = MyQ & ",SezioneRuiPreposto = '"    & apici(SezioneRuiPreposto) & "'"
      MyQ = MyQ & ",NumeroRuiPreposto ='"      & apici(NumeroRuiPreposto) & "'"
      MyQ = MyQ & ",DataIscrizioneRuiPreposto = " & DataIscrizioneRuiPreposto
      MyQ = MyQ & ",CognomeAmministratore = '"       & apici(CognomeAmministratore) & "'"
      MyQ = MyQ & ",NomeAmministratore = '"          & apici(NomeAmministratore) & "'"
      MyQ = MyQ & ",CodiceFiscaleAmministratore = '" & apici(CodiceFiscaleAmministratore) & "'"
      MyQ = MyQ & ",PartitaIvaAmministratore = '"    & apici(PartitaIvaAmministratore) & "'"
      MyQ = MyQ & ",IndirizzoAmministratore = '"     & apici(IndirizzoAmministratore) & "'"
      MyQ = MyQ & ",CittaAmministratore = '"         & apici(CittaAmministratore) & "'"
      MyQ = MyQ & ",ProvinciaAmministratore = '"     & apici(ProvinciaAmministratore) & "'"
      MyQ = MyQ & ",CapAmministratore = '"           & apici(CapAmministratore) & "'"
	  MyQ = MyQ & ",StatoNascitaAmministratore = '"  & apici(StatoNascitaAmministratore) & "'"
	  MyQ = MyQ & ",ComuneNascitaAmministratore = '" & apici(ComuneNascitaAmministratore) & "'"
	  MyQ = MyQ & ",ProvinciaNascitaAmministratore = '" & apici(ProvinciaNascitaAmministratore) & "'"
	  MyQ = MyQ & ",DataNascitaAmministratore = "    & DataNascitaAmministratore 
	  MyQ = MyQ & ",IdSesso = '"                     & apici(IdSesso) & "'"
	  MyQ = MyQ & ",IdSessoPreposto = '"             & apici(IdSessoPreposto) & "'"
	  MyQ = MyQ & ",IdSessoAmministratore = '"       & apici(IdSessoAmministratore) & "'"
	  MyQ = MyQ & ",IdProfiloProdotto = "            & numForDb(IdProfiloProdotto)

	  
      MyQ = MyQ & " Where IdCollaboratore = " & IdCollaboratore
	  'response.write MyQ
	  ConnMsde.execute MyQ 
      If Err.Number <> 0 Then 
	     Ritorna=false 
         MsgErrore = ErroreDb(Err.description)
	  else
	    myUpd = ""
	    myUpd = myUpd & " Update Account Set Nominativo='" & apici(Nominativo) & "'"
		myUpd = myUpd & " Where IdAccount=" & IdAccount
		ConnMsde.execute MyUpd	  
      end if 
   end if 

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
	 
   'recupero i dati 
  if cdbl(IdCollaboratore)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  Collaboratore "
	  MySql = MySql & " Where IdCollaboratore=" & IdCollaboratore
	  xx=GetInfoRecordset(DizDatabase,MySql)
	  IdTipoPers=Getdiz(DizDatabase,"IdTipoDitta")
	  IdTipoColl=Getdiz(DizDatabase,"IdTipoCollaboratore")
	  IdAccount =Cdbl(Getdiz(DizDatabase,"IdAccount"))
  end if 
     
   DescPageOper="Aggiornamento"
   if OperAmmesse="R" then 
      DescPageOper = "Consultazione"
   elseIf cdbl(IdCollaboratore)=0 then 
      DescPageOper = "Inserimento"
   end if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdCollaboratore"  ,IdCollaboratore)
  xx=setValueOfDic(Pagedic,"IdAccount"        ,IdAccount)
  xx=setValueOfDic(Pagedic,"OperAmmesse"      ,OperAmmesse)
  xx=setValueOfDic(Pagedic,"IdTipoColl"       ,IdTipoColl)
  xx=setValueOfDic(Pagedic,"IdTipoPers"       ,IdTipoPers)  
  xx=setValueOfDic(Pagedic,"PaginaReturn"     ,PaginaReturn)
  xx=setValueOfDic(Pagedic,"OperTabella"      ,OperTabella)
 
  xx=setCurrent(NomePagina,livelloPagina) 
  DescLoaded="0"  
  
  showElabCf = false 
  %>
<% 
  'xx=DumpDic(SessionDic,NomePagina)
%>
<div class="d-flex" id="wrapper">

	<%
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
				<div class="col-11"><h3>Gestione Collaboratore :</b> <%=DescPageOper%> </h3>
				</div>
			</div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

    <%
	  stdClass="class='form-control form-control-sm'"
      l_Id = "0"
	  err.clear
      ReadOnly=""
	  SoloLettura=false
      if instr(OperAmmesse,"U")=0 or (instr(OperAmmesse,"I")>0 and cdbl("0" & IdCollaboratore)>0) then 
         SoloLettura=true
         ReadOnly=" readonly "
      end if 
   
   %>
   <div class="row">
      <div class="col-2"><p class="font-weight-bold">Tipo Collaboratore</p></div>   
      <div class = "col-3">
	     <%
		 q = ""
		 q = q & "Select * from TipoCollaboratore "
		 if Cdbl(IdCollaboratore)>0 then 
            IdTipoColl=Getdiz(DizDatabase,"IdTipoCollaboratore") 
		 end if 
		 if IdTipoColl<>"" then 
		    q = q & " where IdTipoCollaboratore='" & IdTipoColl & "'"
		 else 
		    NextL=cdbl(session("LivelloAccount") + 1)
		    q = q & " where LivelloMinimo>=" & NextL & " and LivelloMassimo<= " & NextL 
		 end if 
		 q = q & " order By DescTipoCollaboratore"
	     response.write ListaDbChangeCompleta(q,"IdTipoCollaboratore" & l_Id,IdTipoColl ,"IdTipoCollaboratore","DescTipoCollaboratore" ,0,"","","","","",stdClass)
	     %>
      </div>

      <div class="col-2"><p class="font-weight-bold">Tipo Utenza</p></div>   
      <div class = "col-3">
	     <%
		 q = "select * from TipoDitta where IdTipoditta='" & apici(IdTipoPers) & "'"
		 valo = LeggiCampo(q,"DescTipoDitta")
	     %>
		 <input type="text" readonly class="form-control" value="<%=valo%>" >	 
      </div>
      <div class="col-2"><p class="font-weight-bold"> </p>
      </div> 
   </div>
   <%
   NameLoaded= ""
   if IdTipoPers="PEGI" then
      NameLoaded= NameLoaded & "Denominazione,TE"   		  
   elseif  IdTipoPers="DITT" then 
      NameLoaded= NameLoaded & "Denominazione,TE"   		  
      NameLoaded= NameLoaded & ";Cognome,TE"   		  
	  NameLoaded= NameLoaded & ";Nome,TE"   		  
   else 
      NameLoaded= NameLoaded & "Cognome,TE"   		  
	  NameLoaded= NameLoaded & ";Nome,TE"   		  
   end if 
   %>
   
   
   <%if IdTipoPers="PEGI" or IdTipoPers="DITT" then%>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Denominazione</p>
      </div> 
	  <div class="col-8">
	  	  <%
 		  nome="Denominazione" & l_id
		  valo=Getdiz(DizDatabase,"Denominazione")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>    
   <%end if%> 
   
   <%if IdTipoPers="PEFI" or IdTipoPers="DITT" then
      showElabCf = true 
      lblC0 = "Cognome" 
	  lblC1 = "Nome"
      colss = "col-3"
   %>

   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold"><%=lblC0%></p>
      </div> 
	  <div class="<%=colss%>">
	  	  <%
		  nome="Cognome" & l_id
		  valo=Getdiz(DizDatabase,"Cognome")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div>
	  <%if IdTipoPers<>"PEGI" then%>
      <div class="col-2">
         <p class="font-weight-bold"><%=lblC1%></p>
      </div> 
	  <%end if %>
	  <div class="<%=colss%>">
	  	  <%
		  nome="Nome" & l_id
		  valo=Getdiz(DizDatabase,"Nome")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>   
   <%end if %>
   <%if IdTipoPers="PEFI" or IdTipoPers="DITT" then%>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Stato Nascita</p>
      </div>
      <div class = "col-3">
	     <%
		 valo=Getdiz(DizDatabase,"StatoNascita")
		 if valo="" then 
		    valo="IT"
		 end if
         IdStato=Valo 
		 q = ""
		 q = q & "Select * from Stato "
		 if readonly<>"" then 
            q = q & " Where IdStato='" & valo & "'" 
		 end if 
		 q = q & " order By DescStato"
		 onC = "absoluteChangeStato('StatoNascita0','ProvinciaNascita0');"
	     response.write ListaDbChangeCompleta(q,"StatoNascita" & l_Id,valo ,"IdStato","DescStato" ,0,onC,"","","","",stdClass)
	     %>
       </div>
      <div class="col-2">
         <p class="font-weight-bold">Provincia Nascita</p>
      </div> 
	  <div class="col-3">
	  	  <%
          NameLoaded= NameLoaded & ";ProvinciaNascita,TE"   		  
		  nome="ProvinciaNascita" & l_id
		  valo=Getdiz(DizDatabase,"ProvinciaNascita")
		  idProvincia=valo
		  onC = "absoluteChangeProvincia('StatoNascita0','ProvinciaNascita0','ComuneNascita0');"
		  listDataProv=""
		  listDataComu=""
		  if IdStato="IT" then 
		     listDataProv="absoluteProvinciaIT"
			 if IdProvincia<>"" then 
			    IdProvincia = getSiglaProvinciaDaProvincia(IdProvincia)
				if IdProvincia<>"" then 
			       listDataComu="absoluteComune" & IdProvincia
				end if 
			 end if 
		  end if 
		  %>	  
	      <input type="text" list="<%=listDataProv%>" onchange="<%=onC%>" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div>   
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Comune Nascita</p>
      </div> 
	  <div class="col-3">
	  	  <%
          NameLoaded= NameLoaded & ";ComuneNascita,TE"   		  
		  nome="ComuneNascita" & l_id
		  valo=Getdiz(DizDatabase,"ComuneNascita")
		  %>	  
	      <input type="text" list="<%=listDataComu%>" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div> 
      <div class="col-1">
         <p class="font-weight-bold">Data Nascita</p>
      </div> 
	  <div class="col-2">
	  	  <%
          NameLoaded= NameLoaded & ";DataNascita,TE"   		  
		  nome="DataNascita" & l_id
		  valo=StoD(Getdiz(DizDatabase,"DataNascita"))
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
	  </div>
	  <div class="col-1">
         <p class="font-weight-bold">Sesso</p>
      </div>
	  <div class="col-1">
	  	     <%
		 valo=Getdiz(DizDatabase,"IdSesso")
		 if valo="" then 
		    valo="M"
		 end if
		 q = ""
		 q = q & "Select * from Sesso "
		 if readonly<>"" then 
            q = q & " Where IdSesso='" & valo & "'" 
		 end if 
		 q = q & " order By DescSesso"
	     response.write ListaDbChangeCompleta(q,"IdSesso" & l_Id,valo ,"IdSesso","DescSesso" ,0,"","","","","",stdClass)
	     %>
		 </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>	  
   </div>

	  
   <%end if %>
   
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Codice fiscale</p>
      </div> 
	  <div class="col-2">
		  <%
		  NameLoaded= NameLoaded & ";CodiceFiscale,CF" 
		  nome="CodiceFiscale" & l_id
		  valo=Getdiz(DizDatabase,"CodiceFiscale")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
      <div class="col-1">
		  <%if Readonly="" and showElabCf then %>
            <a href="#" title="Traduci" onclick="reverseCF('CodiceFiscale0','','','ComuneNascita0','ProvinciaNascita0','','IdSesso0','DataNascita0','','StatoNascita0')">  
               <i class="fa fa-2x fa-retweet"></i>
			</a>
            <a href="#" title="Calcola" onclick="calcolaCFColl()">  
               <i class="fa fa-2x fa-id-card-o"></i>
			</a>			
		  <%end if %>
         
      </div>
	  <%if IdTipoPers="PEFI" then%>
      <div class="col-7">
         <p class="font-weight-bold"> </p>
      </div>
	  
	  <% else %>
      <div class="col-2">
         <p class="font-weight-bold">Partita Iva</p>
      </div> 
	  	  <div class="col-3">
		  <%
		  NameLoaded= NameLoaded & ";PartitaIva,PI" 
		  nome="PartitaIva" & l_id
		  valo=Getdiz(DizDatabase,"PartitaIva")
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
	  <% end if  %>
   </div> 

   <div class="row">
      <div class="col-2"><p class="font-weight-bold">Sezione Rui</p></div>   
      <div class = "col-2">
	     <%
		 valo=Getdiz(DizDatabase,"SezioneRui")
		 NextL=cdbl(session("LivelloAccount") + 1)
		 q = ""
		 q = q & " SELECT * From TipoRui where LivelloMinimo>=" & NextL & " and LivelloMassimo<= " & NextL & " order By DescTipoRui  "
	     response.write ListaDbChangeCompleta(q,"SezioneRui" & l_Id,valo ,"IdTipoRui","DescTipoRui" ,0,"","","","","",stdClass)
	     %>
      </div>

      <div class="col-1"><p class="font-weight-bold">Num. RUI</p></div>   
      <div class = "col-2">
		  <%
		  NameLoaded= NameLoaded & ";NumeroRui,TE" 
		  nome="NumeroRui" & l_id
		  valo=Getdiz(DizDatabase,"NumeroRui")
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
      </div>
	  <div class="col-1"><p class="font-weight-bold">Iscritto il</p></div>
      <div class = "col-2">
		  <%
		  NameLoaded= NameLoaded & ";DataIscrizioneRui,DTO" 
		  nome="DataIscrizioneRui" & l_id
		  valo=StoD(Getdiz(DizDatabase,"DataIscrizioneRui"))
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >		  
      </div>
   </div>
   

	<div class="row">
     
       <div class="col-2">
		  <p class="font-weight-bold">Profilo Prodotti Base</p>
	   </div>
      <div class = "col-4">
	     <%
		 valo=Getdiz(DizDatabase,"IdProfiloProdotto")
		 IdProfiloProdotto = Valo
		 'controllo se associato ed esistente 
		 if Cdbl(IdProfiloProdotto)>0 then 
		    v = ""
			v = v & " select top 1 A.IdProfiloProdotto "
			v = v & " from ProfiloProdotto A, AccountProfiloProdotto B"
			v = v & " Where A.IdProfiloProdotto = " & IdProfiloProdotto
			v = v & " and   A.IdProfiloProdotto = B.IdProfiloProdotto"
			v = v & " and   B.IdAccount = " & NumForDb(Getdiz(DizDatabase,"IdAccount"))
			t = cdbl("0" & LeggiCampo(v,"IdProfiloProdotto"))
			if Cdbl(t)=0 then 
			   IdProfiloProdotto = 0
			   valo = 0
			end if 
		 end if 
		 q = ""
		 q = q & " SELECT * From ProfiloProdotto "
		 q = q & " order By DescProfiloProdotto"
		 'response.write q
	     response.write ListaDbChangeCompleta(q,"IdProfiloProdotto" & l_Id,valo ,"IdProfiloProdotto","DescProfiloProdotto" ,1,"","","","","",stdClass)
	     %>
      </div>	   
	   <div class="col-2">
		  <p class="font-weight-bold"> </p>
	   </div>

	</div> 


	<div class="row">
     
       <div class="col-2">
		  <p class="font-weight-bold">Team Leader</p>
	   </div>
	   <div class ="col-3"> 
	   <%
	   TeamLeader=""
	   if Getdiz(DizDatabase,"FlagTeamLeader")=1 then 
	      TeamLeader=" Checked "
	   end if 
	   
	   %>	   
	      <input id="checkTeam<%=l_Id%>" <%=TeamLeader%> name="checkTeam<%=l_Id%>" 
				type="checkbox" value = "S" class="big-checkbox" >
                <span class="font-weight-bold">Si</span>
	   </div> 
	   <div class="col-2">
		  <p class="font-weight-bold">Team</p>
	   </div> 

      <div class = "col-3">
	     <%
		 valo=Getdiz(DizDatabase,"IdTeam")
		 q = ""
		 q = q & " SELECT * From Collaboratore Where IdAzienda=1 and FlagTeamLeader=1 and IdCollaboratore<> " & IdCollaboratore
		 q = q & " order By Denominazione"
		 'response.write q
	     response.write ListaDbChangeCompleta(q,"IdTeam" & l_Id,valo ,"IdCollaboratore","Denominazione" ,1,"","","","","",stdClass)
	     %>
      </div>	   
	   <div class="col-2">
		  <p class="font-weight-bold"> </p>
	   </div>

	</div> 
    <%if IdTipoPers="PEGI" then%>
	
	  <a class="btn btn-info" data-toggle="collapse" href="#collapseAmministratore" role="button" 
		 aria-expanded="false" aria-controls="collapseAmministratore">
		 <span Id="Amministratore_Plus"><i class="fa fa-1x fa-plus-circle"></i></span>
		 <span Id="Amministratore_Minus" style= "display:none"><i class="fa fa-1x fa-minus-circle"></i></span>
		 <input type="hidden" id="Amministratore_plusMinus" value = "+">
		 </a>
		 <B> Amministratore </B>
	  </p> 
	  

		<div class="collapse" id="collapseAmministratore">
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Cognome</p></div> 
			  <div class="col-3">
				  <%
				  nome="CognomeAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"CognomeAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Nome</p></div> 
			  <div class="col-3">
				  <%
				  nome="NomeAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"NomeAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p>
			  </div>
		   </div> 

  
		   <div class="row">
			  <div class="col-2">
				 <p class="font-weight-bold">Stato Nascita</p>
			  </div>
			  
			  <div class = "col-3">
				 <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "StatoNascitaAmministratore,LI"  				 
				 valo=Getdiz(DizDatabase,"StatoNascitaAmministratore")
				 if valo="" then 
					valo="IT"
				 end if
				 IdStatoAmministratore=Valo 
				 q = ""
				 q = q & "Select * from Stato "
				 if readonly<>"" then 
					q = q & " Where IdStato='" & valo & "'" 
				 end if 
				 q = q & " order By DescStato"
				 onC = "absoluteChangeStato('StatoNascitaAmministratore0','ProvinciaNascitaAmministratore0');"
				 response.write ListaDbChangeCompleta(q,"StatoNascitaAmministratore" & l_Id,valo ,"IdStato","DescStato" ,0,onC,"","","","",stdClass)
				 %>
			   </div>			  
			  <div class="col-2">
				 <p class="font-weight-bold">Provincia Nascita</p>
			  </div> 
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "ProvinciaNascitaAmministratore,TE"   		  
				  nome="ProvinciaNascitaAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"ProvinciaNascitaAmministratore")
				  IdProvinciaAmmi=Valo
		          onC = "absoluteChangeProvincia('StatoNascitaAmministratore0','ProvinciaNascitaAmministratore0','ComuneNascitaAmministratore0');"
                  listDataProvAmmi=""
		          listDataComuAmmi=""
				  if IdStatoAmministratore="IT" then 
					 listDataProvAmmi="absoluteProvinciaIT"
					 if IdProvinciaAmmi<>"" then 
						IdProvinciaAmmi = getSiglaProvinciaDaProvincia(IdProvinciaAmmi)
						if IdProvinciaAmmi<>"" then 
						   listDataComuAmmi="absoluteComune" & IdProvinciaAmmi
						end if 
					 end if 
				  end if 				  
				  %>	  
				  <input type="text" list="<%=listDataProvAmmi%>" onchange="<%=onC%>" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div> 	  

			  <div class="col-2">
				 <p class="font-weight-bold"> </p>
			  </div>
		   </div>
  
		   <div class="row">
			  <div class="col-2">
				 <p class="font-weight-bold">Comune Nascita</p>
			  </div> 
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "ComuneNascitaAmministratore,TE"   		  
				  nome="ComuneNascitaAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"ComuneNascitaAmministratore")
				  
				  %>	  
				  <input type="text"  list="<%=listDataComuAmmi%>" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div> 		   
			  <div class="col-1">
				 <p class="font-weight-bold">Data Nascita</p>
			  </div> 
			  <div class="col-2">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "DataNascitaAmministratore,TE"   		  
				  nome="DataNascitaAmministratore" & l_id
				  valo=StoD(Getdiz(DizDatabase,"DataNascitaAmministratore"))
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
						 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
			  </div> 
		  	  <div class="col-1">
                 <p class="font-weight-bold">Sesso</p>
             </div>
	  <div class="col-1">
	  	     <%
		 valo=Getdiz(DizDatabase,"IdSessoAmministratore")
		 if valo="" then 
		    valo="M"
		 end if
		 q = ""
		 q = q & "Select * from Sesso "
		 if readonly<>"" then 
            q = q & " Where IdSesso='" & valo & "'" 
		 end if 
		 q = q & " order By DescSesso"
	     response.write ListaDbChangeCompleta(q,"IdSessoAmministratore" & l_Id,valo ,"IdSesso","DescSesso" ,0,"","","","","",stdClass)
	     %>
		 </div>
			  <div class="col-2">
				 <p class="font-weight-bold"> </p>
			  </div>	  
		   </div>
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Codice fiscale</p></div> 
			  <div class="col-2">
				  <%
				  nome="CodiceFiscaleAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"CodiceFiscaleAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-1">
				  <%if Readonly=""  then %>
					<a href="#!" title="Traduci" onclick="reverseCF('CodiceFiscaleAmministratore0','','','ComuneNascitaAmministratore0','ProvinciaNascitaAmministratore0','','IdSessoAmministratore0','DataNascitaAmministratore0','','StatoNascitaAmministratore0')">  
					   <i class="fa fa-2x fa-retweet"></i>
					</a>
					<a href="#!" title="Calcola" onclick="calcolaCFAmmiPrep('Amministratore')">  
					   <i class="fa fa-2x fa-id-card-o"></i>
					</a>			
				  <%end if %>
				 
			  </div> 	 			  
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
		   
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Indirizzo</p></div> 
			  <div class="col-3">
				  <%
				  nome="IndirizzoAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"IndirizzoAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Citta'</p></div> 
				  <div class="col-3">
				  <%
				  nome="CittaAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"CittaAmministratore")
				  %>
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Provincia</p></div> 
			  <div class="col-3">
				  <%
				  nome="ProvinciaAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"ProvinciaAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Cap</p></div> 
				  <div class="col-3">
				  <%
				  nome="CapAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"CapAmministratore")
				  %>
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
		
		</div>  <!-- fine sezione amministratore -->
	  <%end if  %>
	  
      <%if IdTipoPers="PEGI" then%>
	  <a class="btn btn-info" data-toggle="collapse" href="#collapsePreposto" role="button" 
		 aria-expanded="false" aria-controls="collapsePreposto">
		 <span Id="Preposto_Plus"><i class="fa fa-1x fa-plus-circle"></i></span>
		 <span Id="Preposto_Minus" style= "display:none"><i class="fa fa-1x fa-minus-circle"></i></span>
		 <input type="hidden" id="Preposto_plusMinus" value = "+">
		 </a>
		 <B> Preposto </B>
		 <%
		 'puo' copiare da amministratore 
		 if readonly="" then 
			%>
			<button type="button" class="btn btn-success" onclick="copiaDaAmm();">Copia da amministratore</button>
		 <%
		 end if 
		 %>
	  </p> 
	  

		<div class="collapse" id="collapsePreposto">
   <div class="row">
      <div class="col-2"><p class="font-weight-bold">Cognome</p></div> 
	  <div class="col-3">
	  	  <%
		  nome="CognomePreposto" & l_id
		  valo=Getdiz(DizDatabase,"CognomePreposto")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div>
      <div class="col-2"><p class="font-weight-bold">Nome</p></div> 
	  <div class="col-3">
	  	  <%
		  nome="NomePreposto" & l_id
		  valo=Getdiz(DizDatabase,"NomePreposto")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
      <div class="col-2"><p class="font-weight-bold"> </p>
      </div>
   </div> 
		   <div class="row">
			  <div class="col-2">
				 <p class="font-weight-bold">Stato Nascita</p>
			  </div>
			  
			  <div class = "col-3">
				 <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "StatoNascitaPreposto,LI"  				 
				 valo=Getdiz(DizDatabase,"StatoNascitaPreposto")
				 if valo="" then 
					valo="IT"
				 end if
				 IdStatoPreposto=Valo 
				 q = ""
				 q = q & "Select * from Stato "
				 if readonly<>"" then 
					q = q & " Where IdStato='" & valo & "'" 
				 end if 
				 q = q & " order By DescStato"
				 onC = "absoluteChangeStato('StatoNascitaPreposto0','ProvinciaNascitaPreposto0');"
				 response.write ListaDbChangeCompleta(q,"StatoNascitaPreposto" & l_Id,valo ,"IdStato","DescStato" ,0,onC,"","","","",stdClass)
				 %>
			   </div>			  
			  <div class="col-2">
				 <p class="font-weight-bold">Provincia Nascita</p>
			  </div>
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "ProvinciaNascitaPreposto,TE"   		  
				  nome="ProvinciaNascitaPreposto" & l_id
				  valo=Getdiz(DizDatabase,"ProvinciaNascitaPreposto")
				  IdProvinciaPrep=Valo
		          onC = "absoluteChangeProvincia('StatoNascitaPreposto0','ProvinciaNascitaPreposto0','ComuneNascitaPreposto0');"
                  listDataProvPrep=""
		          listDataComuPrep=""
				  if IdStatoPreposto="IT" then 
					 listDataProvPrep="absoluteProvinciaIT"
					 if IdProvinciaPrep<>"" then 
						IdProvinciaPrep = getSiglaProvinciaDaProvincia(IdProvinciaPrep)
						if IdProvinciaPrep<>"" then 
						   listDataComuPrep="absoluteComune" & IdProvinciaPrep
						end if 
					 end if 
				  end if 				  
				  %>	  
				  <input type="text" list="<%=listDataProvPrep%>"  <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div> 			  
          </div>
   
		   <div class="row">
			  <div class="col-2">
				 <p class="font-weight-bold">Comune Nascita</p>
			  </div> 
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "ComuneNascitaPreposto,TE"   		  
				  nome="ComuneNascitaPreposto" & l_id
				  valo=Getdiz(DizDatabase,"ComuneNascitaPreposto")
				  
				  %>	  
				  <input type="text"  list="<%=listDataComuPrep%>" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div> 		   
			  <div class="col-1">
				 <p class="font-weight-bold">Data Nascita</p>
			  </div> 
			  <div class="col-2">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "DataNascitaPreposto,DT"   		  
				  nome="DataNascitaPreposto" & l_id
				  valo=StoD(Getdiz(DizDatabase,"DataNascitaPreposto"))
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
						 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
			  </div> 
		  	  <div class="col-1">
                 <p class="font-weight-bold">Sesso</p>
             </div>
	  <div class="col-1">
	  	     <%
		 valo=Getdiz(DizDatabase,"IdSessoPreposto")
		 if valo="" then 
		    valo="M"
		 end if
		 q = ""
		 q = q & "Select * from Sesso "
		 if readonly<>"" then 
            q = q & " Where IdSesso='" & valo & "'" 
		 end if 
		 q = q & " order By DescSesso"
	     response.write ListaDbChangeCompleta(q,"IdSessoPreposto" & l_Id,valo ,"IdSesso","DescSesso" ,0,"","","","","",stdClass)
	     %>
		 </div>
   </div>
   
 		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Codice fiscale</p></div> 
			  <div class="col-2">
				  <%
				  nome="CodiceFiscalePreposto" & l_id
				  valo=Getdiz(DizDatabase,"CodiceFiscalePreposto")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-1">
				  <%if Readonly=""  then %>
					<a href="#!" title="Traduci" onclick="reverseCF('CodiceFiscalePreposto0','','','ComuneNascitaPreposto0','ProvinciaNascitaPreposto0','','IdSessoPreposto0','DataNascitaPreposto0','','StatoNascitaPreposto0')">  
					   <i class="fa fa-2x fa-retweet"></i>
					</a>
					<a href="#!" title="Calcola" onclick="calcolaCFAmmiPrep('Preposto')">  
					   <i class="fa fa-2x fa-id-card-o"></i>
					</a>			
				  <%end if %>
				 
			  </div> 	 			  
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
   
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Indirizzo Preposto</p></div> 
			  <div class="col-3">
				  <%
				  nome="IndirizzoPreposto" & l_id
				  valo=Getdiz(DizDatabase,"IndirizzoPreposto")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Citta' Preposto</p></div> 
				  <div class="col-3">
				  <%
				  nome="CittaPreposto" & l_id
				  valo=Getdiz(DizDatabase,"CittaPreposto")
				  %>
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Provincia Preposto</p></div> 
			  <div class="col-3">
				  <%
				  nome="ProvinciaPreposto" & l_id
				  valo=Getdiz(DizDatabase,"ProvinciaPreposto")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Cap Preposto</p></div> 
				  <div class="col-3">
				  <%
				  nome="CapPreposto" & l_id
				  valo=Getdiz(DizDatabase,"CapPreposto")
				  %>
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
		   
   <div class="row">
      <div class="col-2"><p class="font-weight-bold">Sezione Rui</p></div>   
      <div class = "col-2">
	     <%
		 valo=Getdiz(DizDatabase,"SezioneRuiPreposto")
		 q = ""
		 q = q & " SELECT * From TipoRui where LivelloMinimo>=0 and LivelloMassimo<=99 order By DescTipoRui  "
	     response.write ListaDbChangeCompleta(q,"SezioneRuiPreposto" & l_Id,valo ,"IdTipoRui","DescTipoRui" ,0,"","","","","",stdClass)
	     %>
      </div>

      <div class="col-1"><p class="font-weight-bold">Num. RUI</p></div>   
      <div class = "col-2">
		  <%
		  nome="NumeroRuiPreposto" & l_id
		  valo=Getdiz(DizDatabase,"NumeroRuiPreposto")
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
      </div>
	  <div class="col-1"><p class="font-weight-bold">Iscritto il</p></div>
      <div class = "col-2">
		  <%
		  NameLoaded= NameLoaded & ";DataIscrizioneRuiPreposto,DT" 
		  nome="DataIscrizioneRuiPreposto" & l_id
		  valo=StoD(Getdiz(DizDatabase,"DataIscrizioneRuiPreposto"))
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >		  
      </div>
   </div>
</div>
      

   <%end if %>

   <!--#include virtual="/gscVirtual/include/setDataForCall.asp"--> 
   
   <%
   NomeStruttura     = "SEDI_COLLABORATORE"
   DescStruttura     = "Sedi Collaboratore"
   flagOperStruttura = "CUD"
   ProfiloAccount    = "COLL"
   %>
   <!--#include virtual="/gscVirtual/configurazioni/sedi/StrutturaSede.asp"-->    

   <%
   NomeStruttura     = "CONTATTI_COLLABORATORE"
   DescStruttura     = "Contatti Collaboratore"
   flagOperStruttura = "CUD"
   ProfiloAccount    = "COLL"
   %>
   <!--#include virtual="/gscVirtual/configurazioni/contatti/StrutturaContatto.asp"--> 
   
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
   <%end if %>
   
			<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
<%
  xx=createDataList("STATO","dataList_Stato","")
  xx=createDataList("PROVINCIA_IT","absoluteProvinciaIT","")
  
  if listDataComu<>"" then 
     'response.write "creo0:" & listDataComu & " per:" & idProvincia 
     xx=createDataList("COMUNE_BYSIGLAPROV_IT",listDataComu,idProvincia)
  end if 
  
  if listDataComuAmmi<>"" and listDataComuAmmi<>listDataComu then 
     'response.write "creo1:" & listDataComuAmmi & " per:" & idProvinciaAmmi
     xx=createDataList("COMUNE_BYSIGLAPROV_IT",listDataComuAmmi,idProvinciaAmmi)
  end if 
  if listDataComuPrep<>"" and listDataComuPrep<>listDataComuAmmi and listDataComuPrep<>listDataComu then 
     'response.write "creo2:" & listDataComuPrep & " per:" & idProvinciaPrep  
     xx=createDataList("COMUNE_BYSIGLAPROV_IT",listDataComuPrep,idProvinciaPrep)
  end if 
  
%>
			
			</form>
<!--#include virtual="/gscVirtual/include/FormSoggetti.asp"-->
		</div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

</body>

</html>
