<%
  NomePagina="UtenteModifica.asp"
  titolo="Utenti per Azienda"
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
  IdUtente=0
  if FirstLoad then 
	 IdUtente   = "0" & Session("swap_IdUtente")
	 if Cdbl(IdUtente)=0 then 
		IdUtente = cdbl("0" & getValueOfDic(Pagedic,"IdUtente"))
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
  else
	 IdAccount       = "0" & getValueOfDic(Pagedic,"IdAccount")
	 IdUtente        = "0" & getValueOfDic(Pagedic,"IdUtente")
	 OperAmmesse     = getValueOfDic(Pagedic,"OperAmmesse")
	 OperTabella     = getValueOfDic(Pagedic,"OperTabella") 
	 PaginaReturn    = getValueOfDic(Pagedic,"PaginaReturn")
   end if 
   IdUtente = cdbl(IdUtente)
   IdAccount = cdbl(IdAccount)
   if OperAmmesse="" then 
      if IdUtente = 0 then 
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
      Cognome       = Request("Cognome0")
      Nome          = Request("Nome0")
      Nominativo    = trim(trim(Cognome) & " " & Trim(Nome)) 
      CodiceFiscale = Request("CodiceFiscale0")
      PartititaIva  = Request("PartitaIva0")
      Email         = Request("Email0")
      Password      = Request("Password0")
	  CheckAttivo  =Request("CheckAttivo0")
	  if CheckAttivo<>"S" then  
	     CheckAttivo="N"
	  end if 
	  DescBlocco=Request("DescBlocco0")


      if Cdbl(IdUtente)=0 then 
		 if Cdbl(IdAccount)=0 then 
		    MsgErrore="errore di sistema : contattare assistenza"
         else 
		    ConnMsde.execute "Update Account Set IdAzienda=" & Session("IdAziendaWork") & ",IdTipoAccount='BackO',FlagAttivo='N',Abilitato=1 Where IdAccount=" & IdAccount
            MyQ = "" 
            MyQ = MyQ & " Insert into Utente (IdAccount,IdAzienda,DescUtente,Cognome,Nome,Email)"
            MyQ = MyQ & " values (" & IdAccount & "," & Session("IdAziendaWork") & ",'" & Apici(Nominativo) & "'"
            MyQ = MyQ & ",'" & apici(Cognome) & "','" & Apici(Nome) & "','" & apici(Email) & "')"
            ConnMsde.execute MyQ 
			'response.write MyQ
            If Err.Number <> 0 Then 
               MsgErrore = ErroreDb(Err.description)
			   IdAccount=0
            else
			   Ritorna=true 
               IdUtente = GetTableIdentity("Utente")    
            end if 
         end if 
      end if 
      'aggiorno Utente 
      MyQ = "" 
      MyQ = MyQ & " update Utente set "
      MyQ = MyQ & " Cognome = '"       & apici(Cognome) & "'"
      MyQ = MyQ & ",Nome = '"          & apici(Nome)    & "'"
	  MyQ = MyQ & ",DescUtente = '"    & apici(Nominativo) & "'"
	  MyQ = MyQ & ",Email = '"         & apici(Email) & "'"
	  MyQ = MyQ & ",Password = '"      & apici(cripta(Password)) & "'"
      MyQ = MyQ & ",CodiceFiscale = '" & apici(CodiceFiscale) & "'"
      MyQ = MyQ & ",PartitaIva = '"    & apici(PartititaIva) & "'"
      MyQ = MyQ & " Where IdUtente = " & IdUtente
	 
	  ConnMsde.execute MyQ 
      If Err.Number <> 0 Then 
	     Ritorna=false 
         MsgErrore = ErroreDb(Err.description)
      else 
	     ConnMsde.execute "Update Account Set Nominativo='" & apici(Nominativo) & "' Where IdAccount=" & IdAccount
	     MsgErrore = UpdateLoginAccount(IdAccount,Email,cripta(Password),CheckAttivo,DescBlocco)
      end if 
   end if 

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
	 
   'recupero i dati 
   CheckAttivo = "N"
   DescBlocco  = ""
   if cdbl(IdUtente)>0 then
      MySql = ""
      MySql = MySql & " Select * From  Utente "
      MySql = MySql & " Where IdUtente=" & IdUtente
      xx=GetInfoRecordset(DizDatabase,MySql)
      IdAccount = Cdbl(Getdiz(DizDatabase,"IdAccount"))
	  CheckAttivo = LeggiCampo("Select * from Account where idAccount=" & IdAccount,"FlagAttivo")
	  DescBlocco  = LeggiCampo("Select * from Account where idAccount=" & IdAccount,"DescBlocco")
  end if 

     
   DescPageOper="Aggiornamento"
   if OperAmmesse="R" then 
      DescPageOper = "Consultazione"
   elseIf cdbl(IdUtente)=0 then 
      DescPageOper = "Inserimento"
   end if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdUtente"     ,IdUtente)
  xx=setValueOfDic(Pagedic,"IdAccount"    ,IdAccount)
  xx=setValueOfDic(Pagedic,"OperAmmesse"  ,OperAmmesse)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setValueOfDic(Pagedic,"OperTabella"  ,OperTabella)
  
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
				<div class="col-11"><h3>Gestione Utente :</b> <%=DescPageOper%> </h3>
				</div>
			</div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

    <%
	  stdClass="class='form-control form-control-sm'"
      l_Id = "0"
	  err.clear
      ReadOnly=""
	  SoloLettura=false
      if instr(OperAmmesse,"U")=0 or (instr(OperAmmesse,"I")>0 and cdbl("0" & IdUtente)>0) then 
         SoloLettura=true
         ReadOnly=" readonly "
      end if 
   
   %>
  
   
   <%
   NameLoaded= ""
   NameLoaded= NameLoaded & "Cognome,TE"   
   NameLoaded= NameLoaded & ";Nome,TE" 
   %>

   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Cognome</p>
      </div> 
	  <div class="col-3">
	  	  <%
		  nome="Cognome" & l_id
		  valo=Getdiz(DizDatabase,"Cognome")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div>
      <div class="col-2">
         <p class="font-weight-bold">Nome</p>
      </div> 

	  <div class="col-3">
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
   
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Codice fiscale</p>
      </div> 
	  <div class="col-3">
		  <%
		  NameLoaded= NameLoaded & ";CodiceFiscale,CF" 
		  nome="CodiceFiscale" & l_id
		  valo=Getdiz(DizDatabase,"CodiceFiscale")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
	  <%if false then %>
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
	  <%end if %>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div> 

   <div class="row">
	  <div class="col-2"><p class="font-weight-bold">E-mail </p></div>  
      <div class = "col-3">
		  <%
		  NameLoaded= NameLoaded & ";Email,EMO" 
		  nome="Email" & l_id
		  valo=Getdiz(DizDatabase,"EMail")
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                 class="form-control  " >		  
      </div>

      <div class="col-2">
         <p class="font-weight-bold">Password</p>
      </div> 
	  	  <div class="col-3">
		  <%
		  valo=decripta(Getdiz(DizDatabase,"password"))
		  %>
		  <input value="<%=valo%>" type="text" name="Password0" id="Password0" class="form-control" <%=readonly%> >
	  </div>
         <div class="col-2">
		     <%if SoloLettura=false then 
		          RiferimentoA="center;#;;2;lucc;Genera;;creaPassword('Password0');S"%>
		         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->						 
 			 <%else %>
             <p class="font-weight-bold"> </p>
			 <%end if %>
         </div>

   </div>     
    <%
   	FlagAttivo=""
	if CheckAttivo="S" then 
	   FlagAttivo=" checked "
	end if 
	%> 
	<div class="row">
     
	   <div class="col-2">
		  <p class="font-weight-bold">Attivo</p>
	   </div>

	   <div class ="col-8"> 
	   <input id="checkAttivo<%=l_Id%>" <%=FlagAttivo%> name="checkAttivo<%=l_Id%>" 
				type="checkbox" value = "S" class="big-checkbox" >
                <span class="font-weight-bold">Abilitato</span>
	   </div>

	   <div class="col-2">
		  <p class="font-weight-bold"> </p>
	   </div>

	</div>
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Blocco Utente"                 'descrizione label 
   ao_nid = "DescBlocco" & l_Id            'nome ed id
   ao_val = "|value=" & DescBlocco
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		 
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
