<%
  NomePagina="ProfiloProdottiModifica.asp"
  titolo="Profilo Prodotti"
  default_check_profile="Admin,Coll"  
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

function cambia()
{
   ImpostaValoreDi("Oper","cambia");
   document.Fdati.submit();
}

function localFun(Op,Id)
{
	xx=ImpostaValoreDi("DescLoaded","0");
	xx=ElaboraControlli();
	
 	if (xx==false)
	   return false;

    // almeno un campo deve essere inserito 
	var ok=false;
	var rm=ValoreDi("IdRamo0");
	if (rm!="-1")
	   ok=true;
	rm=ValoreDi("IdAnagServizio0");
	if (rm!="-1")
	   ok=true;
	rm=ValoreDi("IdCompagnia0");
	if (rm!="-1")
	   ok=true;

	if (ok==false) {
	   alert('Selezionare almeno un parametro');  
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
   NameLoaded= "DescProfiloProdotto,TE"
   
   FirstLoad=(Request("CallingPage")<>NomePagina)
   IdProfiloProdotto=0
   if FirstLoad then 
	  IdProfiloProdotto   = "0" & Session("swap_IdProfiloProdotto")
	  if Cdbl(IdProfiloProdotto)=0 then 
		 IdProfiloProdotto = cdbl("0" & getValueOfDic(Pagedic,"IdProfiloProdotto"))
	  end if 
	  PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	  if PaginaReturn="" then 
		 PaginaReturn = Session("swap_PaginaReturn")
      end if 
   else
      IdProfiloProdotto   = "0" & getValueOfDic(Pagedic,"IdProfiloProdotto")
      PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   end if 
   IdProfiloProdotto = cdbl(IdProfiloProdotto)
   if IsAdmin() or IsSupervisor() then 
      IdAccount = 0
   else
      IdAccount = Session("LoginIdAccount")
   end if   
  'inizio elaborazione pagina

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
  
   xx=SetDiz(DizDatabase,"IdProdotto",0)
   xx=SetDiz(DizDatabase,"IdCompagnia",0)
   xx=SetDiz(DizDatabase,"IdRamo",0)
   xx=SetDiz(DizDatabase,"IdAnagServizio","")
   xx=SetDiz(DizDatabase,"IdAnagCaratteristica",0)
   xx=SetDiz(DizDatabase,"IdRischio",0)
  
  'recupero i dati 
  if cdbl(IdProfiloProdotto)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  ProfiloProdotto "
	  MySql = MySql & " Where IdProfiloProdotto=" & IdProfiloProdotto
	  xx=GetInfoRecordset(DizDatabase,MySql)
  end if 
  
  if OperTabella="CALL_DEL" then 
     SoloLettura=true
  end if 
  'inserisco il fornitore 
  descD    = trim(Request("DescProfiloProdotto0"))
  IdRam    = cdbl("0" & Request("IdRamo0"))
  IdRischio= cdbl("0" & Request("IdRischio0"))
  idAna    = Request("IdAnagServizio0")
  if IdAna="-1" then 
     IdAna=""
  end if 
idCom    = cdbl("0" & Request("IdCompagnia0"))  

  idPro    = 0
  idAnaCar = 0
  


  peso     = getPeso(IdRam,IdRischio,idAna,idCom,idPro,"","","","","")
  err.clear
  
  if Oper=ucase("update") and cdbl(IdProfiloProdotto)=0 then 
    Session("TimeStamp")=TimePage
	KK="0"

	MyQ = "" 
	MyQ = MyQ & " INSERT INTO ProfiloProdotto (IdTipoProfilo,IdAccount,DescProfiloProdotto,IdProdotto,IdCompagnia,IdRamo"
	MyQ = MyQ & " ,IdAnagCaratteristica,IdAnagServizio,IdSubRamo,Peso,IdRischio)" 
	MyQ = MyQ & " values (" 
	MyQ = MyQ & " 'PROFILO'"
	MyQ = MyQ & ", " & idAccount
	MyQ = MyQ & ",'" & apici(descD) & "'"
	MyQ = MyQ & ", " & idPro
	MyQ = MyQ & ", " & numforDb(idCom)
	MyQ = MyQ & ", " & numForDb(idRam)
	MyQ = MyQ & ", " & numFordb(idAnaCar)	
	MyQ = MyQ & ",'" & apici(idAna) & "'"
	MyQ = MyQ & ", " & numFordb(idSub)
	MyQ = MyQ & ",'" & apici(Peso) & "'"
	MyQ = MyQ & ", " & numFordb(idRischio)
	MyQ = MyQ & ")"

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If
  end if 
  
  if Oper=ucase("update")  and Cdbl(IdProfiloProdotto)>0 then 

     MyQ = "" 
     MyQ = MyQ & " Update ProfiloProdotto "
     MyQ = MyQ & " Set DescProfiloProdotto = '"  & apici(descD) & "'"
     MyQ = MyQ & ",IdCompagnia="          & NumForDb(IdCom)
     MyQ = MyQ & ",IdRamo="               & NumForDb(IdRam)
	 MyQ = MyQ & ",IdRischio="            & NumForDb(IdRischio)
     MyQ = MyQ & ",IdProdotto="           & NumforDb(IdPro)
     MyQ = MyQ & ",IdAnagServizio='"      & apici(idAna) &"'"
     MyQ = MyQ & ",IdAnagCaratteristica=" & numFordb(idAnaCar)	
     MyQ = MyQ & ",IdSubRamo="            & NumForDb(idSub)
	 MyQ = MyQ & ",Peso='"                & apici(Peso) &"'"
     MyQ = MyQ & " Where IdProfiloProdotto = " & IdProfiloProdotto
	'response.write MyQ
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
  end if

  DescPageOper="Aggiornamento"
   if OperTabella="V" then 
      DescPageOper = "Consultazione"
   elseIf OperTabella="CALL_INS" then 
      DescPageOper = "Inserimento"
   elseIf OperTabella="CALL_DEL" then 
      DescPageOper = "Cancellazione"	  
   end if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdProfiloProdotto" ,IdProfiloProdotto)
  xx=setValueOfDic(Pagedic,"PaginaReturn"      ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescLoaded="0"  
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
				<div class="col-11"><h3>Gestione Profilo Prodotto:</b> <%=DescPageOper%> </h3>
				</div>
			</div>

			
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   l_Id = "0"
   
   ao_lbd = "Descrizione Profilo"       'descrizione label 
   ao_nid = "DescProfiloProdotto" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"DescProfiloProdotto")       'valore di default
   if oper="CAMBIA" then 
      ao_val = "|value=" & Request("DescProfiloProdotto0")
   end if
	   
   ao_Plh = "|placeholder=Descrizione Profilo"              'placeholder
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		
  
   
  <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   	<%
	   ao_lbd = "Ramo"             'descrizione label 
       ao_nid = "IdRamo0"          'nome ed id
       idRamo = GetDiz(DizDatabase,"IdRamo")
	   ao_Att = "1"
	   if oper="CAMBIA" then 
	      idRamo=Request("IdRamo0")
	   end if
       ao_val = idRamo     
	   ao_Tex = "select * from Ramo "
	   'non modificabile se IdProdotto>0 
	   ao_Tex = ao_Tex & " order by DescRamo"
	   'response.write ao_Tex
	   ao_ids = "IdRamo"                  'valore della select 
	   ao_des = "DescRamo"                'valore del testo da mostrare 
	   ao_cla = ""                        'azzero classe
	   ao_Eve = "cambia()" 'azzero evento
	                         'indica se deve mettere vuoto 
	   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
	   ao_Cla = "class='form-control form-control-sm'"   
	%>
	<!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->  

    <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   	<%
	   ao_lbd = "Rischio"             'descrizione label 
       ao_nid = "IdRischio0"          'nome ed id
       idRischio = GetDiz(DizDatabase,"IdRischio")
	   if oper="CAMBIA" then 
	      idRischio=Request("idRischio0")
		  if idRischio="-1" then 
		     idRischio=0
		  end if 
	   end if 	   
	   ao_Att = "1"
       ao_val = idRischio     
	   ao_Tex = "select * from Rischio "
	   if cdbl(IdRamo)>0 then 
	      ao_Tex = ao_Tex & " Where IdRamo=" & idRamo
	   end if 
	   ao_Tex = ao_Tex & " order by DescRischio"
	   'response.write ao_Tex
	   ao_ids = "IdRischio"                  'valore della select 
	   ao_des = "DescRischio"                'valore del testo da mostrare 
	   ao_cla = ""                        'azzero classe
	   ao_Eve = "cambia()"
	   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
	   ao_Cla = "class='form-control form-control-sm'" 
	%>
	<!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->    
	
	
	
	
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   	<%
	   ao_lbd = "Servizio Riferimento"             'descrizione label 
       ao_nid = "IdAnagServizio0"          'nome ed id
	   IdAnagServizio = GetDiz(DizDatabase,"IdAnagServizio")
	   ao_Att = "1" 
	   if oper="CAMBIA" then 
	      IdAnagServizio=Request("IdAnagServizio0")
		  if IdAnagServizio="-1" then 
		     IdAnagServizio=""
		  end if 
	   end if 
	   ao_val = IdAnagServizio
	   
	   ao_Tex = "select * from AnagServizio"
	   if cdbl(idRischio)>0 then 
	      IdAnagCara=LeggiCampo("Select * from Rischio Where IdRischio=" & IdRischio,"IdAnagCaratteristica") 
	      IdAnagServ=LeggiCampo("Select * from AnagCaratteristica Where IdAnagCaratteristica=" & IdAnagCara,"IdAnagServizio")
		  ao_Tex = ao_Tex & " Where IdAnagServizio='" & IdAnagServ & "'"
	   end if 
	   ao_Tex = ao_Tex & " order By DescAnagServizio"
	   'response.write ao_Tex
	   ao_ids = "IdAnagServizio"          'valore della select 
	   ao_des = "DescAnagServizio"        'valore del testo da mostrare 
	   ao_cla = ""                        'azzero classe
	   ao_Eve = "cambia()"                'azzero evento
	                         'indica se deve mettere vuoto 
	   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
	   ao_Cla = "class='form-control form-control-sm'" & disab  
	%>
	<!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"--> 
      
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   	<%
	   'escludo le compagnie già scelte 
	   'response.write qEx 
	   ao_lbd = "Compagnia"             'descrizione label 
       ao_nid = "IdCompagnia0"          'nome ed id
       ao_val = GetDiz(DizDatabase,"IdCompagnia")
	   ao_Tex = "select * from Compagnia "
	   disab="  "
	   ao_Att = "1" 
	   ao_Tex = ao_Tex & "order by DescCompagnia"
	   compEx = "0" & LeggiCampo(ao_Tex,"IdCompagnia")
	   ao_ids = "IdCompagnia"             'valore della select 
	   ao_des = "DescCompagnia"           'valore del testo da mostrare 
	   ao_cla = ""                        'azzero classe
	   ao_Eve = ""                        'azzero evento
	                         'indica se deve mettere vuoto 
	   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
	   ao_Cla = "class='form-control form-control-sm'" & disab  
	%>
	<!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->   


   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->		
	
   <%if SoloLettura=false and CompagniaAssente=false then%>
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div></div>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
		</div>
	<%elseif OperTabella="CALL_DEL" then  %>
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Rimuovi; Rimuovi;localFun('submit','0');S"%>
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

		</div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

</body>

</html>
