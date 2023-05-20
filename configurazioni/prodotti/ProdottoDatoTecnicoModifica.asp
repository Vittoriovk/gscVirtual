<%
  NomePagina="ProdottoDatoTecnicoModifica.asp"
  titolo="Prodotto compagnia: dati tecnici"
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
  IdProdotto=0
  if FirstLoad then 
     PaginaReturn    = getCurrentValueFor("PaginaReturn")
     IdProdotto      = cdbl("0" & getCurrentValueFor("IdProdotto"))
	 IdDatoTecnico   = getCurrentValueFor("IdDatoTecnico")
     OperTabella     = Session("swap_OperTabella")
  else
	 IdProdotto      = cdbl("0" & getValueOfDic(Pagedic,"IdProdotto"))
	 IdDatoTecnico   = getValueOfDic(Pagedic,"IdDatoTecnico")
	 OperTabella     = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn    = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdProdotto = cdbl(IdProdotto)
  if Cdbl(IdProdotto)=0 then 
     response.redirect RitornaA(PaginaReturn)
     response.end 
  end if 
  MySql = "" 
  MySql = MySql & " Select * From ProdottoDatoTecnico "
  MySql = MySql & " Where IdDatoTecnico = '" & IdDatoTecnico & "'"
  MySql = MySql & " and IdProdotto=" & IdProdotto
 
  if Oper=ucase("Update") then 
     if Request("checkObbligatorio0")="S" then 
	    FlagObbligatorio = 1
     else
	    FlagObbligatorio = 0
     end if 
	 IdDatoTecnico
	 Ordine        = Cdbl("0" & Request("Ordine0"))
	 Rigo          = Cdbl("0" & Request("Rigo0"))
	 qUpd=""
     if cdbl(IdDatoTecnico)=0 then 
	    IdDatoTecnico = Cdbl(Request("IdDatoTecnico0"))
	    qUpd = qUpd & " insert into ProdottoDatoTecnico (IdProdotto,IdDatoTecnico,FlagObbligatorio,Ordine,rigo)"
		qUpd = qUpd & " values("
		qUpd = qUpd & "  " & IdProdotto
		qUpd = qUpd & ", " & IdDatoTecnico
		qUpd = qUpd & ", " & FlagObbligatorio
		qUpd = qUpd & ", " & Ordine 
		qUpd = qUpd & ", " & Rigo
		qUpd = qUpd & " )"
	 else
	    qUpd = qUpd & " update ProdottoDatoTecnico set "
		qUpd = qUpd & " FlagObbligatorio = " & FlagObbligatorio
		qUpd = qUpd & ",Ordine = " & Ordine
		qUpd = qUpd & ",rigo = " & Rigo
        qUpd = qUpd & " Where IdDatoTecnico = " & IdDatoTecnico
        qUpd = qUpd & " and IdProdotto=" & IdProdotto
	 end if 
	 if qUpd<>"" then 
	    err.clear 
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
  xx=setValueOfDic(Pagedic,"IdProdotto" ,IdProdotto)
  xx=setValueOfDic(Pagedic,"IdDatoTecnico"      ,IdDatoTecnico)
  xx=setValueOfDic(Pagedic,"PaginaReturn"       ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescProdotto = LeggiCampo("select * from Prodotto where idProdotto=" & IdProdotto,"DescProdotto")
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
				<div class="col-11"><h3>Elenco Dati Tecnici</b></h3>
				</div>
			</div>
	        <div class="row">
	           <div class="col-1">
	           </div>
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("Prodotto")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescProdotto%>" >
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
nameLoaded = ";IdDatoTecnico,LI;Ordine,INO;Rigo,INO"
l_Id = "0"
%>
<br>
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Dato Tecnico</p>
	   </div>
	   <div class ="col-8"> 

				   <%
				     qIn = ""
				     qIn = qIn & " select IdDatoTecnico from ProdottoDatoTecnico"
					 qIn = qIn & " where IdProdotto=" & IdProdotto
					 qIn = qIn & " and IdDatoTecnico <> " & IdDatoTecnico
		
				     stdClass="class='form-control form-control-sm'"
					 q = ""
		             q = q & " Select * from DatoTecnico "
					 q = q & " where IdDatoTecnico not in (" & qIn & ") "
		             q = q & " order By DescDatoTecnico"
	                 response.write ListaDbChangeCompleta(q,"IdDatoTecnico" & l_Id,IdDatoTecnico ,"IdDatoTecnico","DescDatoTecnico" ,1,"","","","","",stdClass)
					 
					 DatiAssenti=Cdbl("0" & LeggiCampo(q,"IdDatoTecnico"))
	               %>	   
	   </div>

	</div>
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

	    <%if DatiAssenti>0 then %>
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
