<%
  NomePagina="CambioPassword.asp"
  titolo="Cambio Password"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
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
  if FirstLoad then 
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	 if PaginaReturn="" then 
		PaginaReturn = Session("swap_PaginaReturn")
	 end if 
  else
     PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
  end if   
  IdAccount=Session("LoginIdAccount")
  
  'inserisco account 
   if Oper=ucase("update") then 
      Password      = Request("Password0")
      MyQ = "" 
      MyQ = MyQ & " update Account set "
      MyQ = MyQ & " Password = '"      & apici(cripta(Password)) & "'"
      MyQ = MyQ & " Where IdAccount = " & IdAccount
      ConnMsde.execute MyQ 

   end if 

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
	 
   if cdbl(IdAccount)>0 then
      MySql = ""
      MySql = MySql & " Select * From  Account "
      MySql = MySql & " Where IdAccount=" & IdAccount
      xx=GetInfoRecordset(DizDatabase,MySql)
  end if 

     
   DescPageOper="Cambio Password"

  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
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

  
   
   <%
   NameLoaded= ""
   NameLoaded= NameLoaded & "Password,TE"   
   %>
 

   <br>
   <div class="row">
      <div class="col-1">
      </div> 
      <div class="col-1">
         <p class="font-weight-bold">Password</p>
      </div> 
	  	  <div class="col-3">
		  <%
		  valo=decripta(Getdiz(DizDatabase,"password"))
		  %>
		  <input value="<%=valo%>" type="text" name="Password0" id="Password0" class="form-control" <%=readonly%> >
	  </div>
       <div class="col-2">
	          <%
	          RiferimentoA="center;#;;2;lucc;Genera;;creaPassword('Password0');S"%>
	         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->						 
       </div>

   </div>     

		<div class="row"><div class="col-2"></div>
		<%RiferimentoA="col-2 text-center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div>
  
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
