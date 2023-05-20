<%
  NomePagina="ProdottoOpzioneModifica.asp"
  titolo="Prodotto : opzioni aggiuntive"
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
     PaginaReturn     = getCurrentValueFor("PaginaReturn")
     IdProdotto       = "0" & getCurrentValueFor("IdProdotto")
	 IdOpzione        = getCurrentValueFor("IdOpzione")
     OperTabella      = Session("swap_OperTabella")
  else
	 IdProdotto       = "0" & getValueOfDic(Pagedic,"IdProdotto")
	 IdOpzione        = getValueOfDic(Pagedic,"IdOpzione")
	 OperTabella      = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdProdotto = cdbl(IdProdotto)
  if Cdbl(IdProdotto)=0 or trim(IdOpzione)="" then 
     response.redirect RitornaA(PaginaReturn)
     response.end 
  end if 
  MySql = "" 
  MySql = MySql & " Select * From ProdottoOpzione "
  MySql = MySql & " Where IdOpzione = '" & IdOpzione & "'"
  MySql = MySql & " and IdProdotto=" & IdProdotto
 
  if Oper=ucase("Update") then 
     if Request("checkObbligatorio0")="S" then 
	    FlagObbligatorio = 1
     else
	    FlagObbligatorio = 0
     end if 
	 Ordine            = Cdbl("0" & Request("Ordine0"))
	 CostoFisso        = Cdbl("0" & Request("CostoFisso0"))
     PercSuAcquisto    = Cdbl("0" & Request("PercSuAcquisto0"))
     CostoMinimoSuPerc = Cdbl("0" & Request("CostoMinimoSuPerc0"))
	 qUpd=""
     if LeggiCampo(MySql,"IdOpzione")="" then 
	    qUpd = qUpd & " insert into ProdottoOpzione (IdProdotto,IdOpzione,FlagObbligatorio,Ordine"
		qUpd = qUpd & ",CostoFisso,PercSuAcquisto,CostoMinimoSuPerc)"
		qUpd = qUpd & " values("
		qUpd = qUpd & "  " & IdProdotto
		qUpd = qUpd & ",'" & IdOpzione & "'" 
		qUpd = qUpd & ", " & FlagObbligatorio
		qUpd = qUpd & ", " & Ordine 
		qUpd = qUpd & ", " & NumForDb(CostoFisso) 
		qUpd = qUpd & ", " & NumForDb(PercSuAcquisto) 
		qUpd = qUpd & ", " & NumForDb(CostoMinimoSuPerc) 
		qUpd = qUpd & " )"
	 else
	    qUpd = qUpd & " update ProdottoOpzione set "
		qUpd = qUpd & " FlagObbligatorio = " & FlagObbligatorio
		qUpd = qUpd & ",Ordine = " & Ordine
		qUpd = qUpd & ",CostoFisso = " & numForDb(CostoFisso)
		qUpd = qUpd & ",PercSuAcquisto = " & numForDb(PercSuAcquisto)
		qUpd = qUpd & ",CostoMinimoSuPerc = " & numForDb(CostoMinimoSuPerc)
        qUpd = qUpd & " Where IdOpzione = '" & IdOpzione & "'"
        qUpd = qUpd & " and IdProdotto=" & IdProdotto
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
  xx=setValueOfDic(Pagedic,"IdProdotto"   ,IdProdotto)
  xx=setValueOfDic(Pagedic,"IdOpzione"    ,IdOpzione)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescProdotto   = LeggiCampo("select * from Prodotto where idProdotto=" & IdProdotto,"DescProdotto")
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
				<div class="col-11"><h3>Elenco Dati Aggiuntivi</b></h3>
				</div>
			</div>
	        <div class="row">
	           <div class="col-1">
	           </div>
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("Prodotto")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescProdotto%>" >
               </div>	
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("Opzione")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescOpzione%>" >
               </div>				   
			</div>			

<%
set Rs = Server.CreateObject("ADODB.Recordset")

MySql = ""
MySql = MySql & " select * from ProdottoOpzione "
MySql = MySql & " Where IdOpzione = '" & IdOpzione & "'"
MySql = MySql & " and IdProdotto=" & IdProdotto
'response.write MySql 

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde
response.write Err.description 
CheckObbligatorio=""
Ordine=1
CostoFisso       =0
PercSuAcquisto   =0
CostoMinimoSuPerc=0
if Rs.eof=false then 
   if Rs("FlagObbligatorio")=1 then 
      CheckObbligatorio = " checked "
   end if
	  Ordine           = Rs("Ordine")
      CostoFisso       = Rs("CostoFisso")

      PercSuAcquisto   = Rs("PercSuAcquisto")
      CostoMinimoSuPerc= Rs("CostoMinimoSuPerc")
    
end if 
Rs.close 
err.clear 

DescLoaded=""
NumCols = numC + 1
NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
nameLoaded = "Ordine,INO;CostoFisso,FLZ;PercSuAcquisto,FLQ;CostoMinimoSuPerc,FLZ"
l_Id = "0"
%>
<br>
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Presenza Opzione</p>
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
		  <p class="font-weight-bold">Ordine Esposizione</p>
	   </div>

	   <div class ="col-1"> 
	   <input id="Ordine<%=l_Id%>" name="Ordine<%=l_Id%>" type="text" value = "<%=ordine%>" class="form-control input-sm" >
	   </div>

	</div>
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Costo Fisso &euro;</p>
	   </div>

	   <div class ="col-1"> 
	   <input id="CostoFisso<%=l_Id%>" name="CostoFisso<%=l_Id%>" type="text" value = "<%=CostoFisso%>" class="form-control input-sm" >
	   </div>
	</div>
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Incremento % Su Acquisto</p>
	   </div>

	   <div class ="col-1"> 
	   <input id="PercSuAcquisto<%=l_Id%>" name="PercSuAcquisto<%=l_Id%>" type="text" value = "<%=PercSuAcquisto%>" class="form-control input-sm" >
	   </div>
	</div>
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Costo Minimo su % &euro;</p>
	   </div>

	   <div class ="col-1">
	   <input id="CostoMinimoSuPerc<%=l_Id%>" name="CostoMinimoSuPerc<%=l_Id%>" type="text" value = "<%=CostoMinimoSuPerc%>" class="form-control input-sm" >
	   
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
