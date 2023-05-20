<%
  NomePagina="FornitoreListinoModifica.asp"
  titolo="Gestione Listino Fornitore"
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

    var vProf = ValoreDi("IdProfiloProdotto0");
	var vProd = ValoreDi("IdProdotto0");
 	if (!(vProf=='-1') && !(vProd=='-1')) {
	   alert('selezionare solo uno fra gruppo e prodotto');
	   return false;
    }
	var pc = GetNumberAsFloat(ValoreDi("PrezzoCompagnia0"));
	var pf = GetNumberAsFloat(ValoreDi("PrezzoFornitore0"));
	var pd = GetNumberAsFloat(ValoreDi("PrezzoDistribuzione0"));
	var pl = GetNumberAsFloat(ValoreDi("PrezzoListino0"));
	var ch = ValoreDi("checkDef0");
	var cc = ValoreDi("checkColl0");
	
	if (pf<pc) {
	   alert("il prezzo fornitore deve essere maggiore o uguale al prezzo di compagnia");
	   return false;
	}
	if (pd<pf && cc=="N") {
	   alert("il prezzo di distribuzione deve essere maggiore o uguale al prezzo del fornitore");
	   return false;
	}
	
	if (ch=="S") {
	   var pdM = GetNumberAsFloat(ValoreDi("PrezzoDistribuzioneDef0")); 
	   if (pd<pdM) {
	      alert("il prezzo di distribuzione deve essere maggiore o uguale al prezzo di distribuzione minimo");
	      return false;
	   }
	}
	
	if (pl<pd) {
	   alert("il prezzo di listino deve essere maggiore o uguale al prezzo di distribuzione");
	   return false;
	}

	ImpostaValoreDi("Oper","update");
	document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->


<%
  NameLoaded= "ValidoDal,DTO;PrezzoCompagnia,FLZ;PrezzoFornitore,FLZ;PrezzoDistribuzione,FLZ;PrezzoListino,FLZ"
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdAccount                = 0
  IdAccountProdottoListino = 0
  IdAccountFornitore       = 0
  
  DescFornitore=""
  if FirstLoad then 

     IdAccountProdottoListino  = getCurrentValueFor("IdAccountProdottoListino")
     IdAccountFornitore        = getCurrentValueFor("IdAccountFornitore")
     PaginaReturn              = getCurrentValueFor("PaginaReturn")   

     IdAccountProdottoListino  = cdbl("0" & IdAccountProdottoListino)
     IdAccountFornitore        = cdbl("0" & IdAccountFornitore)
 
     if cdbl(IdAccountFornitore)>0 then 
        Rs.CursorLocation = 3 
        Rs.Open "Select * from Fornitore where IdAccount=" & IdAccountFornitore, ConnMsde   
	    DescFornitore = Rs("DescFornitore")
        Rs.close 
     end if  	 
  else
     PaginaReturn             = getValueOfDic(Pagedic,"PaginaReturn") 
     IdAccountProdottoListino = "0" & getValueOfDic(Pagedic,"IdAccountProdottoListino")
     IdAccountFornitore       = "0" & getValueOfDic(Pagedic,"IdAccountFornitore")
     DescFornitore            = getValueOfDic(Pagedic,"DescFornitore")	 
  end if 
 
  'response.write IdAccountProdottoListino
  'response.end 
  
  IdAccountProdottoListino = cdbl("0" & IdAccountProdottoListino)
  IdAccountFornitore       = cdbl("0" & IdAccountFornitore)
  if Cdbl(IdAccountFornitore)=0  then 
     response.redirect virtualPath & PaginaReturn
     response.end 
  end if   
 
  if Oper=ucase("Update") then 
     IdProfiloProdotto   = Request("IdProfiloProdotto0")
	 if IdProfiloProdotto="-1" then 
	    IdProfiloProdotto = 0 
	 else 
        IdProfiloProdotto   = Cdbl("0" & IdProfiloProdotto)
	 end if 
	 IdProdotto   = Request("IdProdotto0")
	 if IdProdotto="-1" then 
	    IdProdotto = 0 
	 else 
	    IdProdotto          = Cdbl("0" & IdProdotto)
	 end if 
	 ValidoDalNew        = DataStringa(Request("ValidoDal0"))
	 PrezzoCompagnia     = Cdbl("0" & Request("PrezzoCompagnia0"))
     PrezzoFornitore     = Cdbl("0" & Request("PrezzoFornitore0"))
     PrezzoDistribuzione = Cdbl("0" & Request("PrezzoDistribuzione0"))
	 PrezzoListino       = Cdbl("0" & Request("PrezzoListino0"))
	 qUpd=""
 
     if cdbl(IdAccountProdottoListino)=0 then 
	    qUpd = qUpd & " insert into AccountProdottoListino (IdAccount,IdProfiloProdotto,IdProdotto,ValidoDal,PrezzoCompagnia"
		qUpd = qUpd & ",PrezzoFornitore,PrezzoDistribuzione,PrezzoListino,IdAccountFornitore,TipoRegola)"
		qUpd = qUpd & " values("
		qUpd = qUpd & "  " & IdAccount
		qUpd = qUpd & ", " & IdProfiloProdotto
		qUpd = qUpd & ", " & IdProdotto
		qUpd = qUpd & ", " & NumForDb(ValidoDalNew) 
		qUpd = qUpd & ", " & NumForDb(PrezzoCompagnia) 
		qUpd = qUpd & ", " & NumForDb(PrezzoFornitore) 
		qUpd = qUpd & ", " & NumForDb(PrezzoDistribuzione) 
		qUpd = qUpd & ", " & NumForDb(PrezzoListino)
		qUpd = qUpd & ", " & NumForDb(IdAccountFornitore)
		qUpd = qUpd & ",'" & Session("LoginTipoUtente") & "'"
		qUpd = qUpd & " )"
	 elseif cdbl("0" & IdAccountProdottoListino)>0 then 
	    qUpd = qUpd & " update AccountProdottoListino set "
		qUpd = qUpd & " ValidoDal = " & ValidoDalNew
		qUpd = qUpd & ",IdProdotto = " & numForDb(IdProdotto)
		qUpd = qUpd & ",IdProfiloProdotto = " & numForDb(IdProfiloProdotto)
		qUpd = qUpd & ",PrezzoCompagnia = " & numForDb(PrezzoCompagnia)
		qUpd = qUpd & ",PrezzoFornitore = " & numForDb(PrezzoFornitore)
		qUpd = qUpd & ",PrezzoDistribuzione = " & numForDb(PrezzoDistribuzione)
		qUpd = qUpd & ",PrezzoListino = " & numForDb(PrezzoListino)		
        qUpd = qUpd & " Where IdAccountProdottoListino = " & IdAccountProdottoListino 
	 end if 
	 if qUpd<>"" then 
	    'response.write qUpd 
	    ConnMsde.execute qUpd 
		if Err.number=0 then
		   response.redirect RitornaA(PaginaReturn)
           response.end 
		else
		   MsgErrore=ErroreDb(err.description)
		end if 
	 end if 

  End if 
  
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdAccountProdottoListino" ,IdAccountProdottoListino)
  xx=setValueOfDic(Pagedic,"IdAccountFornitore"       ,IdAccountFornitore)
  xx=setValueOfDic(Pagedic,"DescFornitore"            ,DescFornitore)
  xx=setValueOfDic(Pagedic,"PaginaReturn"             ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 
  
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
				<div class="col-11"><h3>Listino Prodotti Fornitore </b></h3>
				</div>
			</div>
	        <div class="row">
	           <div class="col-1">
	           </div>
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("Fornitore")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescFornitore%>" >
               </div>
			</div>			

<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<%
   set Rs = Server.CreateObject("ADODB.Recordset")
   Rs.CursorLocation = 3
   if MsgErrore="" then 

      MySql = ""
      MySql = MySql & " select * "
      MySql = MySql & " from AccountProdottoListino "
      MySql = MySql & " Where IdAccountProdottoListino = " & NumForDb(IdAccountProdottoListino)
      'response.write MySql 

      Rs.Open MySql, ConnMsde

	  ValidoDal           = DtoS()
	  IdProdotto          = 0
	  IdProfiloProdotto   = 0
      PrezzoCompagnia     = 0 
      PrezzoFornitore     = 0 
      PrezzoDistribuzione = 0 
      PrezzoListino       = 0 

      if Rs.eof=false then 
	     ValidoDal           = rs("ValidoDal")
         IdProdotto          = rs("IdProdotto")
         IdProfiloProdotto   = rs("IdProfiloProdotto")
         PrezzoCompagnia     = rs("PrezzoCompagnia")
         PrezzoFornitore     = rs("PrezzoFornitore") 
         PrezzoDistribuzione = rs("PrezzoDistribuzione")
         PrezzoListino       = rs("PrezzoListino")
      end if 

      Rs.close 
      err.clear 
   end if 
   
   
   'per inserimento recupero prezzo compagnia/fornitore 
   TrovatoPrezzo=false 
   
   TrovatoPrezzo = true 

   if ValidoDal = 0 then 
      ValidoDal = Stod(Dtos())
   else
      ValidoDal = STod(ValidoDal)
   end if    
DescLoaded=""
NumCols = numC + 1
NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
l_Id = "0"
%>
<div class="row"><div class="col-2"><p class="font-weight-bold"></p></div></div>

	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Raggruppamento Prodotti</p>
	   </div>
	   <div class ="col-6"> 
	   <%
	   stdClass="class='form-control form-control-sm'"
	   q = "SELECT * From ProfiloProdotto Where IdTipoProfilo = 'GRUPPO' order By DescProfiloProdotto"
       response.write ListaDbChangeCompleta(q,"IdProfiloProdotto0",IdProfiloProdotto ,"IdProfiloProdotto","DescProfiloProdotto" ,1,"","","","","",stdClass)  
	   
	   %>
	   </div>	   
	</div>
   
		<div class="row" >
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
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Prodotto</p>
	   </div>
	   <div class ="col-6"> 
	   <%
	   stdClass="class='form-control form-control-sm'"
       QueryProdForm = ""
       QueryProdForm = QueryProdForm & "Select a.* From Prodotto a,AccountProdotto B  Where B.IdAccount=" & IdAccountfornitore
       QueryProdForm = QueryProdForm  & " and A.IdProdotto = B.IdProdotto "
   
       response.write ListaDbChangeCompleta(QueryProdForm,"IdProdotto0",IdProdotto ,"IdProdotto","DescProdotto" ,1,"","","","","",stdClass)  
	   
	   %>
	   </div>	   
	</div>
   
<div class="row"><div class="col-2"><p class="font-weight-bold"></p></div></div>


	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Valido Dal</p>
	   </div>

	   <div class ="col-2"> 
	      <input type="text"  name="ValidoDal<%=l_Id%>" id="ValidoDal<%=l_Id%>" value="<%=ValidoDal%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
	   </div>

	</div>
	<input type="hidden" name="checkColl0" id="checkColl0" value = "N" >
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Prezzo Compagnia &euro;</p>
	   </div>
       <%
	   LockCompagnia = ""
       LockFornitore = ""	   
	   if IsAdmin() or IsCollaboratore() then 
		  LockCompagnia = " readonly "
		  LockFornitore = " readonly "
		  if IsCollaboratore() then 
		     
		  end if 
	   end if 
	   
	   %>
	   <div class ="col-1"> 
	   <input id="PrezzoCompagnia<%=l_Id%>" name="PrezzoCompagnia<%=l_Id%>" <%=LockCompagnia%> type="text" value = "<%=PrezzoCompagnia%>" class="form-control input-sm" >
	   </div>
	</div>

	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Prezzo fornitore &euro;</p>
	   </div>

	   <div class ="col-1"> 
	   <input id="PrezzoFornitore<%=l_Id%>" name="PrezzoFornitore<%=l_Id%>" <%=LockFornitore%> type="text" value = "<%=PrezzoFornitore%>" class="form-control input-sm" >
	   </div>
	</div>

	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Prezzo Distribuzione &euro;</p>
	   </div>

	   <div class ="col-1">
	   <input id="PrezzoDistribuzione<%=l_Id%>" name="PrezzoDistribuzione<%=l_Id%>" type="text" value = "<%=PrezzoDistribuzione%>" class="form-control input-sm" >
	   </div>
	   <input id="PrezzoDistribuzioneDef<%=l_Id%>" name="PrezzoDistribuzioneDef<%=l_Id%>" type="hidden" 
	   value = "<%=PrezzoDistribuzioneDef%>">
	   
       <input type="hidden" name="checkDef0" id="checkDef0" value = "N" >

	</div>
	<div class="row">
	   <div class="col-1">
		  <p class="font-weight-bold"></p>
	   </div>
	   
	   <div class="col-2">
		  <p class="font-weight-bold">Prezzo Listino &euro;</p>
	   </div>

	   <div class ="col-1">
	   <input id="PrezzoListino<%=l_Id%>" name="PrezzoListino<%=l_Id%>" type="text" value = "<%=PrezzoListino%>" class="form-control input-sm" >
	   </div>
	   <input id="PrezzoListinoDef<%=l_Id%>" name="PrezzoListinoDef<%=l_Id%>" type="hidden" value = "<%=PrezzoListinoDef%>" >	   

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
