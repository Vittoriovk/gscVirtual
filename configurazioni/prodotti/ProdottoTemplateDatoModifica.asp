<%
  NomePagina="ProdottoTemplateDatoModifica.asp"
  titolo="Template Prodotto : dati aggiuntivi"
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

  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdProdottoTemplate=0
  if FirstLoad then 
     PaginaReturn     = getCurrentValueFor("PaginaReturn")
     IdProdottoTemplate       = "0" & getCurrentValueFor("IdProdottoTemplate")
	 IdOpzione        = getCurrentValueFor("IdOpzione")
     OperTabella      = Session("swap_OperTabella")
  else
	 IdProdottoTemplate       = "0" & getValueOfDic(Pagedic,"IdProdottoTemplate")
	 IdOpzione        = getValueOfDic(Pagedic,"IdOpzione")
	 OperTabella      = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdProdottoTemplate = cdbl(IdProdottoTemplate)
  if Cdbl(IdProdottoTemplate)=0 or trim(IdOpzione)="" then 
     response.redirect RitornaA(PaginaReturn)
     response.end 
  end if 
  MySql = "" 
  MySql = MySql & " Select * From ProdottoTemplateOpzione "
  MySql = MySql & " Where IdOpzione = '" & IdOpzione & "'"
  MySql = MySql & " and IdProdottoTemplate=" & IdProdottoTemplate
 
  if Oper=ucase("Update") then 
     if Request("checkObbligatorio0")="S" then 
	    FlagObbligatorio = 1
     else
	    FlagObbligatorio = 0
     end if 
	 Ordine = Cdbl("0" & Request("Ordine0"))
	 Rigo   = Cdbl("0" & Request("Rigo0"))
	 qUpd=""
     if LeggiCampo(MySql,"IdOpzione")="" then 
	    qUpd = qUpd & " insert into ProdottoTemplateOpzione (IdProdottoTemplate,IdOpzione,FlagObbligatorio,Ordine,rigo)"
		qUpd = qUpd & " values("
		qUpd = qUpd & "  " & IdProdottoTemplate
		qUpd = qUpd & ",'" & IdOpzione & "'" 
		qUpd = qUpd & ", " & FlagObbligatorio
		qUpd = qUpd & ", " & Ordine 
		qUpd = qUpd & ", " & Rigo
		qUpd = qUpd & " )"
	 else
	    qUpd = qUpd & " update ProdottoTemplateOpzione set "
		qUpd = qUpd & " FlagObbligatorio = " & FlagObbligatorio
		qUpd = qUpd & ",Ordine = " & Ordine
		qUpd = qUpd & ",rigo = " & Rigo
        qUpd = qUpd & " Where IdOpzione = '" & IdOpzione & "'"
        qUpd = qUpd & " and IdProdottoTemplate=" & IdProdottoTemplate
	 end if 
	 if qUpd<>"" then 
	    'response.write qUpd 
	    ConnMsde.execute qUpd 
		if Err.number=0 then
		   response.redirect RitornaA(PaginaReturn)
           response.end 
		else
		   MsgErrore=err.description 
		end if 
	 end if 

  End if 
  
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdProdottoTemplate"   ,IdProdottoTemplate)
  xx=setValueOfDic(Pagedic,"IdOpzione"    ,IdOpzione)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescProdottoTemplate   = LeggiCampo("select * from ProdottoTemplate where idProdottoTemplate=" & IdProdottoTemplate,"DescProdottoTemplate")
  q = "select * from Opzione Where IdOpzione = '" & IdOpzione & "'"
  DescOpzione    = LeggiCampo(q,"DescInterna")
  %>

<div class="d-flex" id="wrapper">
	<%
	  TitoloNavigazione="Configurazioni"
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
			<%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h3>Elenco Dati Standard</b></h3>
				</div>
			</div>
	        <div class="row">
	           <div class="col-1">
	           </div>
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("ProdottoTemplate")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescProdottoTemplate%>" >
               </div>	
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("Dato Aggiuntivo")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescOpzione%>" >
               </div>				   
			</div>			

<%
set Rs = Server.CreateObject("ADODB.Recordset")

'response.write MySql 

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

CheckObbligatorio=""
rigo  =1
Ordine=1
if Rs.eof=false then 
   if Rs("FlagObbligatorio")=1 then 
      CheckObbligatorio = " checked "
	  Ordine=Rs("Ordine")
	  rigo  =Rs("Rigo") 
   end if 
end if 
Rs.close 
err.clear 

DescLoaded=""
NumCols = numC + 1
NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
nameLoaded = "Ordine,INO;Rigo,INO"
l_Id = "0"
%>
<br>
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Presenza Dato</p>
	   </div>

	   <div class ="col-8"> 
	   <input id="checkObbligatorio<%=l_Id%>" <%=CheckObbligatorio%> name="checkObbligatorio<%=l_Id%>" 
				type="checkbox" value = "S" class="big-checkbox" >
                <span class="font-weight-bold">Obbligatorio</span>
	   </div>

	</div>
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Rigo Esposizione</p>
	   </div>

	   <div class ="col-1"> 
	   <input id="Rigo<%=l_Id%>" name="Rigo<%=l_Id%>" type="text" value = "<%=rigo%>" class="form-control input-sm" >
	   </div>

	</div>	
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Ordine Esposizione</p>
	   </div>

	   <div class ="col-1"> 
	   <input id="Ordine<%=l_Id%>" name="Ordine<%=l_Id%>" type="text" value = "<%=ordine%>" class="form-control input-sm" >
	   </div>

	</div>

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
