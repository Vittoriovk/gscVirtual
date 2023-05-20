<%
  NomePagina="ModificaProvvigione.asp"
  titolo="Regola Provvigione"
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
<!--#include virtual="/gscVirtual/js/functionTable.js"-->
<script language="JavaScript">

function localFun(Op,Id)
{
	xx=ResetCampo("IdTipoCalcolo0","","white");
	xx=ImpostaValoreDi("DescLoaded","0");
	xx=ElaboraControlli();
	
 	if (xx==false)
	   return false;

	var vProf = ValoreDi("IdProfiloProdotto0");
	var vProd = ValoreDi("IdProdotto0");
 	if (!(vProf=='-1') && !(vProd=='-1')) {
	   alert('selezionare solo uno fra gruppo e prodotto');
	   return false;
    }
	vv = ValoreDi("IdTipoCalcolo0");
	zz = ValoreDi("PercProvRete0");
	if ((vv=="-1" && zz!="0") || (vv!="-1" && zz=="0")) {
	   xx=ImpostaColoreFocus("IdTipoCalcolo0","","yellow");
	   xx=ImpostaColoreFocus("PercProvRete0","","yellow");
	   alert("Dati non validi");
	   return false;
	}
	ImpostaValoreDi("Oper","update");
	document.Fdati.submit();
}

 
function refresh()
{
	ImpostaValoreDi("Oper","refresh");
	document.Fdati.submit();
}
</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<!--#include virtual="/gscVirtual/modelli/FunctionAccount.asp"-->
  
 <!-- javascript locale -->

<%

  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdRegolaProvvigione=0
  TipoProvvigione="RETE" 
  
  if FirstLoad then 
	 IdRegolaProvvigione   = "0" & Session("swap_IdRegolaProvvigione")
	 if Cdbl(IdRegolaProvvigione)=0 then 
		IdRegolaProvvigione = cdbl("0" & getValueOfDic(Pagedic,"IdRegolaProvvigione"))
	 end if   
	 TipoProvvigione = Session("swap_TipoProvvigione")
	 if TipoProvvigione = "" then 
	    TipoProvvigione="RETE"
     end if 

	 OperAmmesse   = Session("swap_OperAmmesse")
	 OperTabella   = Session("swap_OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	 if PaginaReturn="" then 
		PaginaReturn = Session("swap_PaginaReturn")
	 end if 
  else
	 IdRegolaProvvigione  = "0" & getValueOfDic(Pagedic,"IdRegolaProvvigione")
	 TipoProvvigione      = getValueOfDic(Pagedic,"TipoProvvigione")
	 OperTabella          = getValueOfDic(Pagedic,"OperTabella") 
	 PaginaReturn         = getValueOfDic(Pagedic,"PaginaReturn")
   end if 
   IdRegolaProvvigione = cdbl(IdRegolaProvvigione)
   if isAdmin() then 
      IdAccountRegistratore = 0
   else
      IdAccountRegistratore = Session("LoginIdAccount") 
   end if 
   
   if OperAmmesse="" then 
      if IdRegolaProvvigione = 0 then 
         OperAmmesse="CRUD"
      end if 
   end if 

  'inserisco account 
   if Oper=ucase("update") then 
      Ritorna=false 
	  OperAmmesse="U"
      Session("TimeStamp")=TimePage
      MsgErrore=""
      DescRegolaProvvigione = Request("DescRegolaProvvigione0")
      IdProfiloProdotto     = TestNumeroPos("0" & Request("IdProfiloProdotto0"))
      IdProdotto            = TestNumeroPos("0" & Request("IdProdotto0"))
      PercProvRete          = TestNumeroPos("0" & Request("PercProvRete0"))
      CompensoFisso         = TestNumeroPos("0" & Request("CompensoFisso0"))
      PercOvertax           = 0 
      PercDiritti           = TestNumeroPos("0" & Request("PercDiritti0"))
      PercInterm            = TestNumeroPos("0" & Request("PercInterm0"))
	  
      IdTipoCalcolo     = Request("IdTipoCalcolo0")
	  if IdTipoCalcolo = "-1" then 
	     IdTipoCalcolo = ""
	     PercProvRete=0
      end if 

      IdCollaboratore   = TestNumeroPos("0" & Request("IdCollaboratore0"))
	  IdFornitore       = TestNumeroPos("0" & Request("IdFornitore0"))
	  ValidoDal         = TestNumeroPos("0" & DataStringa(Request("ValidoDal0")))

	  err.clear
      FlagAdded=false 
      if Cdbl(IdRegolaProvvigione)=0 then 

	     FlagAdded=true 
         MyQ = "" 
         MyQ = MyQ & " Insert into RegolaProvvigione (TipoRegola,DescRegolaProvvigione,IdFornitore,IdCollaboratore)"
         MyQ = MyQ & " values ('RETE','" & Apici(DescRegolaProvvigione) & "'," & TimeToS() & "," & TimeToS() & ")"
 
         ConnMsde.execute MyQ 
         If Err.Number <> 0 Then 
            MsgErrore = ErroreDb(Err.description)
         else
            IdRegolaProvvigione = GetTableIdentity("RegolaProvvigione")    
         end if  
	  end if 
      'aggiorno RegolaProvvigione 
	  if cdbl(IdRegolaProvvigione) > 0 then 
         MyQ = "" 
         MyQ = MyQ & " update RegolaProvvigione set "
         MyQ = MyQ & " DescRegolaProvvigione = '" & apici(DescRegolaProvvigione) & "'"
         MyQ = MyQ & ",IdProfiloProdotto = "      & NumForDb(IdProfiloProdotto)  
		 MyQ = MyQ & ",IdFornitore = "            & NumForDb(IdFornitore)  
		 MyQ = MyQ & ",IdCollaboratore = "        & NumForDb(IdCollaboratore) 
         MyQ = MyQ & ",IdProdotto = "             & NumForDb(IdProdotto)  
         MyQ = MyQ & ",PercProvRete = "           & NumForDb(PercProvRete)  
         MyQ = MyQ & ",CompensoFisso = "          & NumForDb(CompensoFisso)
         MyQ = MyQ & ",PercOvertax = "            & NumForDb(PercOvertax)
         MyQ = MyQ & ",PercDiritti = "            & NumForDb(PercDiritti)
		 MyQ = MyQ & ",PercInterm = "             & NumForDb(PercInterm)
         MyQ = MyQ & ",IdTipoCalcolo = '"         & apici(IdTipoCalcolo) & "'"
		 MyQ = MyQ & ",ValidoDal = "              & NumForDb(ValidoDal)
         MyQ = MyQ & " Where IdRegolaProvvigione = " & IdRegolaProvvigione
		 'response.write MyQ
         ConnMsde.execute MyQ 
      end if 
      If Err.Number <> 0 Then 
	     Ritorna=false 
         MsgErrore = ErroreDb(Err.description)
		 if FlagAdded=true then 
		    ConnMsde.execute "Delete From RegolaProvvigione Where IdRegolaProvvigione = " & IdRegolaProvvigione
			IdRegolaProvvigione=0
			Oper="REFRESH"
		 end if 
      end if 
 
   end if 

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
	 
   'recupero i dati 
   if cdbl(IdRegolaProvvigione)>0 then
      MySql = ""
      MySql = MySql & " Select * From  RegolaProvvigione "
      MySql = MySql & " Where IdRegolaProvvigione=" & IdRegolaProvvigione
      xx=GetInfoRecordset(DizDatabase,MySql)
	  PercProvRete        = GetDiz(DizDatabase,"PercProvRete")
      CompensoFisso       = GetDiz(DizDatabase,"CompensoFisso")
      PercOvertax         = GetDiz(DizDatabase,"PercOvertax")
      PercDiritti         = GetDiz(DizDatabase,"PercDiritti")
	  PercInterm          = GetDiz(DizDatabase,"PercInterm")
  else
      PercProvRete  = 0
      CompensoFisso = 0
      PercOvertax   = 0
      PercDiritti   = 0
	  PercInterm    = 0
  end if 

   DescPageOper="Aggiornamento"
   if OperAmmesse="R" then 
      DescPageOper = "Consultazione"
   elseIf cdbl(IdRegolaProvvigione)=0 then 
      DescPageOper = "Inserimento"
   end if
  'registro i dati della pagina 
   xx=setValueOfDic(Pagedic,"IdRegolaProvvigione" ,IdRegolaProvvigione)
   xx=setValueOfDic(Pagedic,"TipoProvvigione"     ,TipoProvvigione)
   xx=setValueOfDic(Pagedic,"OperAmmesse"         ,OperAmmesse)
   xx=setValueOfDic(Pagedic,"PaginaReturn"        ,PaginaReturn)
   xx=setValueOfDic(Pagedic,"OperTabella"         ,OperTabella)
  
   xx=setCurrent(NomePagina,livelloPagina) 
   DescLoaded="0"  
   
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
				<div class="col-11"><h3>Gestione Regola Provvigione Rete :</b> <%=DescPageOper%> </h3>
				</div>
			</div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

    <%
	  stdClass="class='form-control form-control-sm'"
      l_Id = "0"
	  err.clear
      ReadOnly=""
	  SoloLettura=false
	  OperAmmesse="U"
      if instr(OperAmmesse,"U")=0 or (instr(OperAmmesse,"I")>0 and cdbl("0" & IdRegolaProvvigione)>0) then 
         SoloLettura=true
         ReadOnly=" readonly "
      end if 
   
   %>
  
   
   <%
   FlagSelPoss=true
   
   NameLoaded= ""
   NameLoaded= NameLoaded & "DescRegolaProvvigione,TE"   
   NameLoaded= NameLoaded & ";PercProvRete,FLQ" 
   NameLoaded= NameLoaded & ";CompensoFisso,FLZ"    
   NameLoaded= NameLoaded & ";PercDiritti,FLQ" 
   NameLoaded= NameLoaded & ";PercInterm,FLQ" 
   NameLoaded= NameLoaded & ";ValidoDal,DTO"
   %>
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Descrizione Regola"                 'descrizione label 
   ao_nid = "DescRegolaProvvigione" & l_Id            'nome ed id
   if Oper=ucase("refresh") then 
      ao_val = "|value=" & Request("DescRegolaProvvigione0")
   else 
      ao_val = "|value=" & GetDiz(DizDatabase,"DescRegolaProvvigione")
   end if       
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		    
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   ao_lbd = "Collaboratore"                       'descrizione label 
   ao_nid = "IdCollaboratore" & l_Id              'nome ed id
   
   if Oper=ucase("refresh") then 
      ao_val = Request("IdCollaboratore0") 'valore di default
   else 
      ao_val = GetDiz(DizDatabase,"IdCollaboratore") 'valore di default
   end if     

   ao_Tex = "SELECT * From Collaboratore "
   if IsAdmin() then 
      ao_Tex = ao_Tex & " Where IdAccountLivello1=0 "
   else
      ao_Tex = ao_Tex & " Where IdAccountLivello1=0 "
   end if 
   ao_Tex = ao_Tex & " order By Denominazione"
   ao_ids = "IdCollaboratore"             'valore della select 
   ao_des = "Denominazione"           'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->

   
      <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   entity="Fornitore"
   ao_lbd = entity                       'descrizione label 
   ao_nid = "Id" & entity & l_Id              'nome ed id
   
   if Oper=ucase("refresh") then 
      IdFornitore = "0" & request("Id" & entity & l_id)     'valore di default
   else 
      idFornitore = "0" & GetDiz(DizDatabase,"Id" & entity) 'valore di default
   end if 
   
   QueryCompForm = ""
   QueryRamiForm = ""
   IdFornitore=TestNumeroPos(IdFornitore)
   'response.write "idf=" & IdFornitore
   if Cdbl(idFornitore)>0 then 
      IdAccountforn = cdbl("0" & LeggiCampo("select * from Fornitore Where IdFornitore=" & IdFornitore,"IdAccount"))
	  if Cdbl(IdAccountforn) > 0 then 
	     QueryCompForm = "Select IdCompagnia From AccountCompagnia Where IdAccount=" & IdAccountforn
         QueryRamiForm = "Select distinct IdRamo From AccountProdotto A, Prodotto B Where IdAccount=" & IdAccountforn
         QueryRamiForm = QueryRamiForm & " and A.IdProdotto = B.IdProdotto "	  
	  end if 
   else
      IdAccountforn = 0
   end if 
   
   ao_val = IdFornitore 
   ao_Tex = "SELECT * From " & entity
   ao_Tex = ao_Tex & " order By Desc" & entity
   ao_ids = "Id" & entity             'valore della select 
   ao_des = "Desc" & entity           'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = "refresh()"               'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->
<div class="row"><div class="col-2"><p class="font-weight-bold"></p></div></div>

  <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   entity="ProfiloProdotto"
   ao_lbd = "Profilo Prodotti"        
   ao_nid = "Id" & entity & l_Id             
   if Oper=ucase("refresh") then 
      IdComp = request("Id" & entity & l_id)     'valore di default
   else 
      idComp = GetDiz(DizDatabase,"Id" & entity) 'valore di default
   end if 
   ao_val = IdComp
   ao_Tex = "SELECT * From " & entity
   ao_Tex = ao_Tex & " Where IdTipoProfilo = 'PROFILO' " 
   ao_Tex = ao_Tex & " order By Desc" & entity
   ao_ids = "Id" & entity             'valore della select 
   ao_des = "Desc" & entity           'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = "refresh()"               'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   
  
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->
<div class="row   " >
   <div class="col-2">
      <p class="font-weight-bold"></p>
   </div>
   <div class = "col-8">
         <b>Oppure</B>
   </div>
   <div class="col-2">
      <p class="font-weight-bold"> </p>
   </div>

</div>   

   <%
   entity="Prodotto"
   ao_lbd = entity                       'descrizione label 
   ao_nid = "Id" & entity & l_Id              'nome ed id
   if Oper=ucase("refresh") then 
      ao_val = Request("Id" & entity & l_id)
   else 
      ao_val = GetDiz(DizDatabase,"Id" & entity)
   end if      
   QueryProdForm = ""
   QueryProdForm = QueryProdForm & "Select a.IdProdotto From Prodotto a,AccountProdotto B "
   QueryProdForm = QueryProdForm & "Where  A.IdProdotto = B.IdProdotto "
   if Cdbl(IdAccountforn) > 0 then 
      QueryProdForm = QueryProdForm & " and B.IdAccount = " & IdAccountforn
   end if 
   
   ao_Tex = "SELECT * From " & entity
   if QueryProdForm<>"" then 
      ao_Tex = ao_Tex & " Where IdProdotto in (" & QueryProdForm &  ") "   
   end if 
   
   ao_Tex = ao_Tex & " order By Desc" & entity
   ao_ids = "Id" & entity             'valore della select 
   ao_des = "Desc" & entity           'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   
   ii=LeggiCampo(ao_tex,"IdProdotto")
   if Cdbl("0" & ii)=0 then 
      FlagSelPoss=false
   end if  
   
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"--> 

   
   
   <Div class="row">
   <div class="col-2">
       <p class="font-weight-bold">Valida dal</p>
   </div> 
	  <div class="col-2">
	  	  <%
		  nome="ValidoDal" & l_id
          if Oper=ucase("refresh") then 
             ao_val = Request(nome)
          else 
             ao_val = StoD(GetDiz(DizDatabase,"ValidoDal"))
          end if    		  
		  
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=ao_val%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
	  </div>
	</div>  
      <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   entity="TipoCalcolo"
   ao_lbd = entity                       'descrizione label 
   ao_nid = "Id" & entity & l_Id              'nome ed id
   ao_val = GetDiz(DizDatabase,"Id" & entity) 'valore di default
   ao_Tex = "SELECT * From " & entity
   ao_Tex = ao_Tex & " Where FlagOperatore like '%RETE%' "
   ao_Tex = ao_Tex & " order By Desc" & entity
   ao_ids = "Id" & entity             'valore della select 
   ao_des = "Desc" & entity           'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   
   <div class="row" >
      <div class="col-2">
          <p class="font-weight-bold">Provvigioni</p>
      </div>
      <div class = "col-3">
<%
'              ListaDbChangeCompleta(Query ,Name  ,CodValue,ColCod,ColText,FlagVuoto,Change,Campo,Larghezza,DescVuoto,DescNoData,Classe)
       response.write ListaDbChangeCompleta(ao_Tex,ao_nid,ao_val  ,ao_Ids,ao_Des ,ao_Att   ,ao_Eve,""   ,""       ,ao_Plh   ,ao_NoD    ,ao_cla)
%>
      </div>
      <div class="col-1">
          <p class="font-weight-bold">Provv. %</p>
      </div>	  
      <div class = "col-1">
        <input value="<%=PercProvRete%>" type="text" name="PercProvRete0" id="PercProvRete0" class="form-control"  >
      </div>	  
      <div class="col-1">
          <p class="font-weight-bold">Fisso &euro;</p>
      </div>		  
      <div class = "col-1">
        <input value="<%=CompensoFisso%>" type="text" name="CompensoFisso0" id="CompensoFisso0" class="form-control"  >
      </div>	  

   </div>

   <div class="row" >
      <div class="col-2">
          <p class="font-weight-bold">% Su Intermediazione</p>
      </div>

      <div class = "col-1">
        <input value="<%=PercInterm%>" type="text" name="PercInterm0" id="PercInterm0" class="form-control"  >
      </div>
	  

      <div class="col-2">
         <p class="font-weight-bold">% Su Diritti</p>
      </div>
      <div class = "col-1">
        <input value="<%=PercDiritti%>" type="text" name="PercDiritti0" id="PercDiritti0" class="form-control"  >
      </div>  
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div> 

  
     <%
	 if FlagSelPoss=false and MsgErrore="" then 
	    SoloLettura=true
		MsgErrore="Dati Non Disponibili per selezione "
	 %>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
	 <%
     end if 
	 if SoloLettura=false then%>
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
