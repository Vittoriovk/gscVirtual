<%
  NomePagina="ClienteConfigura.asp"
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

  NameLoaded=""
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
   if Oper=ucase("update") then 
   IdAccountModPag = IdAccount
%>
        <!--#include virtual="/gscVirtual/configurazioni/pagamenti/UpdateListaModPag.asp"-->	  
<%
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
				<div class="col-11"><h3>Configurazione Cliente </b></h3>
				</div>
			</div>
   <div class="row">
        <div class="col-2"><p class="font-weight-bold">Cliente</p></div>

         <div class = "col-8">
             <input value="<%=DescCliente%>" type="text" class="form-control" readonly >
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
      MySql = MySql & " Select a.*,isnull(b.UserId,'') as UserId,isnull(b.Password,'') as Password,isnull(B.Abilitato,0) as Abilitato "
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

   %>
   
  
 
    <%
	if isSegnalatore()=false then  
	   IdAccountModPag=IdAccount
       OpDocAmm="U"
   %>
   <!--#include virtual="/gscVirtual/configurazioni/pagamenti/ListaModPag.asp"-->
 
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
