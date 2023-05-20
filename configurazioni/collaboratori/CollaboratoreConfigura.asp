<%
  NomePagina="CollaboratoreConfigura.asp"
  titolo="Utenti per Azienda"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/FunMailWithAttach.asp"-->
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
	if (Op=="submit")
	   ImpostaValoreDi("Oper","update");
	if (Op=="send")
	   ImpostaValoreDi("Oper","update_send");
	   
	document.Fdati.submit();

}

</script>

<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

 
<%

  Set RsRec = Server.CreateObject("ADODB.Recordset")
  NameLoaded="UserId,TE"
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdCollaboratore=0
  if FirstLoad then 
	 IdCollaboratore   = "0" & Session("swap_IdCollaboratore")
	 if Cdbl(IdCollaboratore)=0 then 
		IdCollaboratore = cdbl("0" & getValueOfDic(Pagedic,"IdCollaboratore"))
	 end if   
	 OperTabella   = Session("swap_OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	 if PaginaReturn="" then 
		PaginaReturn = Session("swap_PaginaReturn")
	 end if 
  else
	 IdCollaboratore     = "0" & getValueOfDic(Pagedic,"IdCollaboratore")
	 OperTabella   = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   end if 

   IdCollaboratore = cdbl(IdCollaboratore)
   if Cdbl(IdCollaboratore)=0 then 
      response.redirect RitornaA(PaginaReturn)
	  response.end 
   end if 
   
   idAccount         = LeggiCampo("Select * from Collaboratore Where IdCollaboratore=" & IdCollaboratore,"IdAccount")
   DescCollaboratore = LeggiCampo("select * from Collaboratore Where IdCollaboratore=" & IdCollaboratore,"Denominazione")
   
  'inizio elaborazione pagina

  'inserisco account 
   Ritorna=false 
   SendMail=false 
   DescClie=""
   if Oper=ucase("update_send") then 
      SendMail=true 
      Oper=ucase("update")
   end if 
   'response.write Oper
   if Oper=ucase("update") then 
      CheckAttivo  =Request("CheckAttivo0")
      if CheckAttivo<>"S" then  
         CheckAttivo="N"
         Abilitato=0
	  else 
	     Abilitato=1
	  end if
      CheckGenera  =Request("CheckGenera0")
      if CheckGenera<>"S" then  
         CheckGenera="N"
         Genera=0
	  else 
	     Genera=1
	  end if
	  if IsAdmin() then 
	     IdAccountParametro = IdAccount
	     %>
         <!--#include virtual="/gscVirtual/configurazioni/collaboratori/ListaParametriUpdate.asp"--> 
		 <%
	     IdProcessoElaborativo=Request("IdProcessoElaborativo0")
		 IdEsterno            =Request("IdEsterno0")
	     if IdProcessoElaborativo="-1" then 
	        IdProcessoElaborativo=""
	     end if 
		 if IdEsterno<>"" then 
		    qEsiste = "select * from Collaboratore where IdEsterno='" & IdEsterno & "' and IdAccount<>" & IdAccount
			trovato = LeggiCampo(qEsiste,"Denominazione")
			if Trovato<>"" then 
			   MsgErrore = "Id Univoco gia' assegnato a " & trovato
			end if 
	     end if
	  end if 
	   
	  if MsgErrore="" then 
         MySql = ""
         MySql = MySql & " update Collaboratore set "
         MySql = MySql & " FlagGeneraCollaboratore =" & Genera
	     if IsAdmin() then 
	        MySql = MySql & ",IdProcessoElaborativo ='" & IdProcessoElaborativo & "'"
		    MySql = MySql & ",IdEsterno ='" & IdEsterno & "'"
	     end if 
         MySql = MySql & " where IdAccount=" & IdAccount 
         ConnMsde.execute MySql 
	  end if 
	  
	  DescBlocco=Request("DescBlocco0")
      
      RsRec.CursorLocation = 3
      RsRec.Open "Select * from Collaboratore Where IdCollaboratore=" & IdCollaboratore, ConnMsde
      IdAccount=RsRec("IDAccount")
      IdAzienda=RsRec("IdAzienda")
      DescClie =RsRec("DescCollaboratore")
      RsRec.close 
	  'controllo duplicazione se abilitato 
	  if Abilitato=1 then 
		 MySql = ""
		 MySql = MySql & " select top 1 idAccount From Account "
         MySql = MySql & " where IdAccount<>" & IdAccount 
         MySql = MySql & " and   IdAzienda = " & IdAzienda 
         MySql = MySql & " and   FlagAttivo='S'"  
         MySql = MySql & " and   UserId='" & apici(Request("UserId0")) & "'" 

         v_ret = Cdbl("0" & LeggiCampo(MySql,"IdAccount"))
      else
         v_ret = 0
	  end if 
	  if Cdbl(v_ret)=0 then 
         PassWord=cripta(Request("PassWord0"))
		 MySql = ""
         MySql = MySql & " update Account set "
         MySql = MySql & " UserId   ='" & apici(Request("UserId0")) & "'"
         MySql = MySql & ",PassWord ='" & apici(PassWord) & "'"
         MySql = MySql & ",Abilitato =" & Abilitato
		 MySql = MySql & ",DescBlocco='" & apici(DescBlocco) & "'"
         MySql = MySql & " where IdAccount=" & IdAccount 
         ConnMsde.execute MySql 
      else
         MsgErrore="Utenza esistente"
      end if 
	  IdAccountModPag = IdAccount
	  %>
   <!--#include virtual="/gscVirtual/configurazioni/pagamenti/UpdateListaModPag.asp"-->	  
	  <%
   end if 
   if SendMail then 
      TextHtml=""
	  TextText=""
	  addInfo = "<br>UserId=" & Request("UserId0") & "<br> Password=" & Request("PassWord0")
	  toAddress=Request("UserId0")
      xx=CreaTestoMail(DescClie,"le inviamo con la presente i sui dati di accesso alla piattaforma " & addInfo,TextHtml,TextText)
      xx=SendMailMessageHTMLWithAttach("", ToAddress, "","Accesso alla piattaforma", TextText, TextHtml, "", false) 
	  if xx<>"" then 
	     Msgerrore="Impossibile recapitare il messaggio"
      else
	     Ritorna=true
      end if 
   end if 
   if Ritorna=true then 
      response.redirect virtualPath & PaginaReturn
	  response.end
   end if 

   xx=setValueOfDic(Pagedic,"IdCollaboratore" ,IdCollaboratore)
   xx=setValueOfDic(Pagedic,"OperTabella"     ,OperTabella)
   xx=setValueOfDic(Pagedic,"PaginaReturn"    ,PaginaReturn)
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
				<div class="col-11"><h3>Gestione Configurazione Collaboratore :</b> <%=DescCollaboratore%> </h3>
				</div>
			</div>

   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
   
   <%
   l_id = "0"
   LeggiDati=false
   IdProcessoElaborativo=""
   IdEsterno=""
   if Cdbl(IdCollaboratore)>0 then
      err.clear 
      LeggiDati=true
      Set RsRec = Server.CreateObject("ADODB.Recordset")
      MySql = "" 
      MySql = MySql & " Select a.*,isnull(b.UserId,'') as UserId,isnull(b.Password,'') as Password,isnull(B.Abilitato,0) as Abilitato "
	  MySql = MySql & " ,IsNull(B.DescBlocco,'') as DescBlocco"
	  MySql = MySql & " from Collaboratore A left join Account B "
	  MySql = MySql & " On a.idAccount = b.idAccount "
	  MySql = MySql & " Where a.IdCollaboratore = " & IdCollaboratore

      RsRec.CursorLocation = 3
      RsRec.Open MySql, ConnMsde 

      If Err.number<>0 then	
       	 LeggiDati=false
      elseIf RsRec.EOF then	
         LeggiDati=false
		 RsRec.close 
	  else 
	     IdProcessoElaborativo = RsRec("IdProcessoElaborativo")
		 IdEsterno             = RsRec("IdEsterno")
      End if
   end if    

   ao_val = "" 
   if LeggiDati then 
      ao_val = ao_val & RsRec("UserId")
   end if 
   readonly=""
   if SoloLettura=true then
      readonly=" readonly "
   else 
      contattoIdAccount=RsRec("IdAccount")
	  contattoTipo     ="'Mail'"
	  campoxValore     ="UserId0"
   %>
      <!--#include virtual="/gscVirtual/utility/SelezioneContatto.asp"-->
   <%
   end if 
   %>
   <div class="row">
        <div class="col-1"><p class="font-weight-bold">UserId (e-mail)</p></div>

         <div class = "col-3">
             <input value="<%=ao_val%>" type="text" name="UserId0" id="UserId0" class="form-control" <%=readonly%> >
         </div>

         <div class="col-1">
		     <%if SoloLettura=false then 
		          RiferimentoA="center;#;;2;mail;Seleziona ;; $('#contattoModal').modal('toggle');S"%>
		         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->						 
 			 <%else %>
             <p class="font-weight-bold"> </p>
			 <%end if %>
         </div>

   <%
   ao_val = "" 
   if LeggiDati then 
      ao_val = ao_val & decripta(RsRec("Password"))
   end if 
   %>
   

        <div class="col-1"><p class="font-weight-bold">Password</p></div>

         <div class = "col-3">
             <input value="<%=ao_val%>" type="text" name="Password0" id="Password0" class="form-control" <%=readonly%> >
         </div>

         <div class="col-1">
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
	if RsRec("Abilitato")= "1" then 
	   FlagAttivo=" checked "
	end if 
   	FlagGenera=""
	if RsRec("FlagGeneraCollaboratore")= "1" then 
	   FlagGenera=" checked "
	end if 
	
	
	%> 
	<div class="row">
     
	   <div class="col-2">
		  <p class="font-weight-bold">Accesso alla piattaforma</p>
	   </div>

	   <div class ="col-2"> 
	   <input id="checkAttivo<%=l_Id%>" <%=FlagAttivo%> name="checkAttivo<%=l_Id%>" 
				type="checkbox" value = "S" class="big-checkbox" >
                <span class="font-weight-bold">Abilitato</span>
	   </div>
	   <div class="col-2">
		  <p class="font-weight-bold">Creazione Collaboratori</p>
	   </div>
	   
	   <div class ="col-2"> 
	   <input id="checkGenera<%=l_Id%>" <%=FlagGenera%> name="checkGenera<%=l_Id%>" 
				type="checkbox" value = "S" class="big-checkbox" >
                <span class="font-weight-bold">Abilitato</span>
	   </div>
	</div>
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Blocco Utente"                 'descrizione label 
   ao_nid = "DescBlocco" & l_Id            'nome ed id
   ao_val = "|value=" 
   if LeggiDati then 
      ao_val = ao_val & RsRec("DescBlocco")
   end if 
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		 
   
	<%if IsAdmin() then %>
	<div class="row">   
	   <div class="col-2">
	      <p class="font-weight-bold">Processo di Pagamento</p>
	   </div>
       <div class="col-4">
       <%
         query = ""
         query = query & " Select * from ProcessoElaborativo " 
         query = query & " Where TipoProcesso = 'PAGAMENTO_SERVIZIO' " 
         query = query & " order By DescProcessoElaborativo"
         response.write ListaDbChangeCompleta(Query,"IdProcessoElaborativo0",IdProcessoElaborativo,"IdProcessoElaborativo","DescProcessoElaborativo",1,"","","","","dati assenti","class='form-control form-control-sm'")					 
 
       %>
       </div>
	   <div class="col-2">
	      <p class="font-weight-bold">Codice Univoco</p>
	   </div>
         <div class = "col-3">
             <input value="<%=IdEsterno%>" type="text" name="IdEsterno0" id="IdEsterno0" class="form-control" <%=readonly%> >
         </div>	   
	   
	</div>
	<%end if %>

   
   <%IdAccountModPag=IdAccount
     OpDocAmm="U"
   %>
   <!--#include virtual="/gscVirtual/configurazioni/pagamenti/ListaModPag.asp"-->
   
   
   
   
   
   <%
   if IsAdmin() then
      tmpIdAccountRequest=IdAccount
      listaParametri="'VAL_COB','VAL_ATI','ASS_PRO'"
   %>
      <!--#include virtual="/gscVirtual/configurazioni/collaboratori/ListaParametri.asp"-->
   <%
   end if 
      if leggiDati then 
	     Rs.close
      end if 
   
   
      if SoloLettura=false then%>
		<div class="row">
		    <div class="mx-auto">
		       <%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
		       <!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		     </div>
		    <div class="mx-auto">
		       <%RiferimentoA="center;#;;2;mail;Registra ed Invia; Invia mail;localFun('send','0');S"%>
		       <!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		     </div>			 
		</div>
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
