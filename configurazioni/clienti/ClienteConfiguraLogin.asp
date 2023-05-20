<%
  NomePagina="ClienteConfiguraLogin.asp"
  titolo="Utenti per Azienda"
  default_check_profile="Coll"
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

  NameLoaded="UserId,EM"
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdCliente=0
  if FirstLoad then 
	 IdCliente   = "0" & Session("swap_IdCliente")
	 if Cdbl(IdCliente)=0 then 
		IdCliente = cdbl("0" & getValueOfDic(Pagedic,"IdCliente"))
	 end if   
	 OperTabella   = Session("swap_OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	 if PaginaReturn="" then 
		PaginaReturn = Session("swap_PaginaReturn")
	 end if 
  else
	 IdCliente     = "0" & getValueOfDic(Pagedic,"IdCliente")
	 OperTabella   = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   end if 

   IdCliente = cdbl(IdCliente)
   if Cdbl(IdCliente)=0 then 
      response.redirect RitornaA(PaginaReturn)
	  response.end 
   end if 
  'inizio elaborazione pagina
   DescCliente=LeggiCampo("select * from Cliente Where IdCliente=" & IdCliente,"Denominazione")
   IdAccount  =LeggiCampo("select * from Cliente Where IdCliente=" & IdCliente,"IdAccount")   
  'inserisco account 
   Ritorna=false 
   SendMail=false 
   DescClie=""
   if Oper=ucase("update_send") then 
      SendMail=true 
      Oper=ucase("update")
   end if 
   if Oper=ucase("update") then 
      CheckAttivo  =Request("CheckAttivo0")
      if CheckAttivo<>"S" then  
         CheckAttivo="N"
         Abilitato=0
	  else 
	     Abilitato=1
	  end if 
	  DescBlocco=Request("DescBlocco0")
      Set RsRec = Server.CreateObject("ADODB.Recordset")
      RsRec.CursorLocation = 3
      RsRec.Open "Select * from Cliente Where IdCliente=" & IdCliente, ConnMsde
      IdAccount=RsRec("IDAccount")
      IdAzienda=RsRec("IdAzienda")
      DescClie =RsRec("DescCliente")
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
         MySql = MySql & " update Account set "
         MySql = MySql & " UserId   ='" & apici(Request("UserId0")) & "'"
         MySql = MySql & ",PassWord ='" & apici(PassWord) & "'"
         MySql = MySql & ",Abilitato =" & Abilitato
		 MySql = MySql & ",DescBlocco ='" & apici(DescBlocco) & "'"
         MySql = MySql & " where IdAccount=" & IdAccount 
         ConnMsde.execute MySql 
      else
         MsgErrore="Utenza esistente"
      end if
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
  
   
   DescPageOper=DescCliente

   xx=setValueOfDic(Pagedic,"IdCliente" ,IdCliente)
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
				<div class="col-11"><h3>Configurazione Accesso Cliente</b></h3>
				</div>
			</div>

   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
   
   <%
   l_id = "0"
   LeggiDati=false
   if Cdbl(IdCliente)>0 then
      err.clear 
      LeggiDati=true
      Set RsRec = Server.CreateObject("ADODB.Recordset")
      MySql = "" 
      MySql = MySql & " Select a.*,isnull(b.UserId,'') as UserId,isnull(b.Password,'') as Password"
	  MySql = MySql & " ,isnull(B.Abilitato,0) as Abilitato,isnull(B.DescBlocco,0) as DescBlocco "
	  MySql = MySql & " from Cliente A left join Account B "
	  MySql = MySql & " On a.idAccount = b.idAccount "
	  MySql = MySql & " Where a.IdCliente = " & IdCliente

      RsRec.CursorLocation = 3
      RsRec.Open MySql, ConnMsde 

      If Err.number<>0 then	
       	 LeggiDati=false
      elseIf RsRec.EOF then	
         LeggiDati=false
		 RsRec.close 
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
        <div class="col-2"><p class="font-weight-bold">Cliente</p></div>

         <div class = "col-8">
             <input value="<%=DescCliente%>" type="text" class="form-control" readonly >
         </div>

   </div> 
   
   <div class="row">
        <div class="col-2"><p class="font-weight-bold">UserId (e-mail)</p></div>

         <div class = "col-8">
             <input value="<%=ao_val%>" type="text" name="UserId0" id="UserId0" class="form-control" <%=readonly%> >
         </div>

         <div class="col-2">
		     <%if SoloLettura=false then 
		          RiferimentoA="center;#;;2;mail;Seleziona ;; $('#contattoModal').modal('toggle');S"%>
		         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->						 
 			 <%else %>
             <p class="font-weight-bold"> </p>
			 <%end if %>
         </div>
   </div> 
   <%
   ao_val = "" 
   if LeggiDati then 
      ao_val = ao_val & decripta(RsRec("Password"))
   end if 
   %>
   
   <div class="row">
        <div class="col-2"><p class="font-weight-bold">Password</p></div>

         <div class = "col-8">
             <input value="<%=ao_val%>" type="text" name="Password0" id="Password0" class="form-control" <%=readonly%> >
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
	if RsRec("Abilitato")= "1" then 
	   FlagAttivo=" checked "
	end if 
	%> 
	<div class="row">
     
	   <div class="col-2">
		  <p class="font-weight-bold">Accesso alla piattaforma</p>
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
   <%
   ao_val = "" 
   if LeggiDati then 
      ao_val = RsRec("DescBlocco")
   end if 
   %>

   
   <div class="row">
        <div class="col-2"><p class="font-weight-bold">Blocco Utente</p></div>

         <div class = "col-8">
             <input value="<%=ao_val%>" type="text" name="DescBlocco0" id="DescBlocco0" class="form-control" <%=readonly%> >
         </div>

   </div>
   
    <%
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
