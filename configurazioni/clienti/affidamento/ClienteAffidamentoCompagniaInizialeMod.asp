<%
  NomePagina="ClienteAffidamentoCompagniaInizialeMod.asp"
  titolo="Affidamento cliente per compagnia"
  default_check_profile="BackO"
  act_call_upda = CryptAction("CALL_UPDA") 
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/functionAffidamento.asp"-->
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

	ImpostaValoreDi("Oper","<%=act_call_upda%>");
	document.Fdati.submit();
}
</script>
<body>
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
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
  Set Rs = Server.CreateObject("ADODB.Recordset")

  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdCliente=0
  IdRecMod =0
  if FirstLoad then 
     IdAccount      = getCurrentValueFor("IdAccount")
	 IdTipoSvincolo = getCurrentValueFor("IdTipoSvincolo")
	 IdRecMod       = getCurrentValueFor("IdRecMod")
	 IdCompagnia    = getCurrentValueFor("IdCompagnia")
     IdCauzioneProv = getCurrentValueFor("IdCauzioneProv")
     IdCauzioneDefi = getCurrentValueFor("IdCauzioneDefi")
	 PaginaReturn   = getCurrentValueFor("PaginaReturn") 
  else
	 IdAccount      = getValueOfDic(Pagedic,"IdAccount")
	 IdTipoSvincolo = getValueOfDic(Pagedic,"IdTipoSvincolo")
	 IdRecMod       = getValueOfDic(Pagedic,"IdRecMod")
	 IdCompagnia    = getValueOfDic(Pagedic,"IdCompagnia")
     IdCauzioneProv = getValueOfDic(Pagedic,"IdCauzioneProv")
     IdCauzioneDefi = getValueOfDic(Pagedic,"IdCauzioneDefi")
	 PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
   end if 
   IdAccount = cdbl("0" & IdAccount)
   IdRecMod  = cdbl("0" & IdRecMod)
   DenomClie = LeggiCampo("Select * from Cliente Where IdAccount=" & IdAccount ,"Denominazione")
   DataSvincolo = Dtos()
   DescSvincolo = ""
   ImptSvincolo = 0
   
   if Cdbl(IdRecMod)>0 then    
      MySql = "select * from AccountSvincolo Where IdAccountSvincolo=" & IdRecMod  
      Rs.CursorLocation = 3
      Rs.Open MySql, ConnMsde 
      if rs.eof = false then
         DataSvincolo = rs("DataSvincolo")
         DescSvincolo = rs("DescSvincolo")
         ImptSvincolo = rs("ImptSvincolo")
      end if 
	  rs.close 
   end if 
   
  'registro i dati della pagina 
   xx=setValueOfDic(Pagedic,"IdAccount"        ,IdAccount)
   xx=setValueOfDic(Pagedic,"IdTipoSvincolo"   ,IdTipoSvincolo)
   xx=setValueOfDic(Pagedic,"IdRecMod"         ,IdRecMod)
   xx=setValueOfDic(Pagedic,"IdCompagnia"      ,IdCompagnia)
   xx=setValueOfDic(Pagedic,"IdCauzioneProv"   ,IdCauzioneProv)
   xx=setValueOfDic(Pagedic,"IdCauzioneDefi"   ,IdCauzioneDefi)
   xx=setValueOfDic(Pagedic,"PaginaReturn"     ,PaginaReturn)
 
   xx=setCurrent(NomePagina,livelloPagina) 
   DescLoaded="0"  
  
   Oper = DecryptAction(Oper)
   'eseguo aggiornamento ma controllo i dati
   if ucase(oper)="CALL_UPDA" then 
      ImptSvincolo = Request("ImptSvincolo0")
	  DataSvincolo = Request("DataSvincolo0")
	  DescSvincolo = Request("DescSvincolo0")
	  if Cdbl(IdRecMod)=0 then 
	     qIns = ""
		 qIns = qIns & "INSERT INTO AccountSvincolo"
		 qIns = qIns & "(IdAccount,IdTipoSvincolo,IdCompagnia,DataSvincolo"
		 qIns = qIns & ",DescSvincolo,ImptSvincolo,IdCauzioneProv,IdCauzioneDefi)"
		 qIns = qIns & " values "
		 qIns = qIns & "(" & IdAccount & ",'" & IdTipoSvincolo & "'," & IdCompagnia & ",0"
		 qIns = qIns & ",'',0,0,0"
		 qIns = qIns & ")"
		 ConnMsde.execute qIns 
         if err.number<>0 then 
		    MsgErrore = "Errore di sistema contattare assistenza"
		    xx = writeTraceAttivita(qIns & " " & Err.description,"AffidamentoCompagniaInizialeMod",0)
		 else 
		    IdRecMod=cdbl("0"  & GetTableIdentity("AccountSvincolo"))
            if Cdbl(IdRecMod)=0 then 
		       xx = writeTraceAttivita(qIns & " " & Err.description,"AffidamentoCompagniaInizialeMod",0)
		       MsgErrore = "Errore di sistema contattare assistenza"
		    end if 
	     end if 
	  end if 
      If MsgErrore="" then 
	     DataSvincolo = DataStringa(DataSvincolo)
         qUpd = ""
		 qUpd = qUpd & " update AccountSvincolo set "
		 qUpd = qUpd & " DataSvincolo=" & NumForDb(DataSvincolo)
		 qUpd = qUpd & ",DescSvincolo='" & apici(DescSvincolo) & "'"
		 qUpd = qUpd & ",ImptSvincolo=" & NumForDb(ImptSvincolo)
		 qUpd = qUpd & " Where IdAccountSvincolo = " & IdRecMod 
		 connMsde.execute qUpd 
		 
         qUpd = ""
	     qUpd = qUpd & " select sum(ImptSvincolo) as tot from AccountSvincolo"
	     qUpd = qUpd & " Where IdAccount=" & IdAccount
		 qUpd = qUpd & " and IdTipoSvincolo='" & IdTipoSvincolo & "'"
		 'response.write Qupd 
		 Impt = Cdbl("0" & LeggiCampo(qUpd,"tot"))
		 
         qUpd = ""
	     qUpd = qUpd & " update AccountCreditoAffiTotali "
		 qUpd = qUpd & " set ImptInizialeStornato = " & NumForDb(Impt)
	     qUpd = qUpd & " Where IdAccount=" & IdAccount
		 'response.write qUpd 
         ConnMsde.execute qUpd
		 
      end if   
	  
   end if    
  
   DescPageOper="Aggiornamento"
   
   Dim DizAff
   Set DizAff = CreateObject("Scripting.Dictionary")
   esito=GetTotaliAffidamentoComp(DizAff,IdAccount,IdCompagnia)

   ImptIniziale = GetDiz(DizAff,"ImptIniziale")
   ImptStornato = GetDiz(DizAff,"ImptInizialeStornato")
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
				<div class="col-11"><h3>Affidamento iniziale : svincolo importi </b></h3>
				</div>
			</div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

			<div class="row">
			    <div class="col-1"></div>
				<div class="col-4">
					<div class="form-group ">
						<%xx=ShowLabel("Utente")%>
						<input type="text" readonly class="form-control" value="<%=DenomClie%>" >
					</div>		
				</div>
				<div class="col-4">
					<div class="form-group ">
						<%xx=ShowLabel("Compagnia")
						DenomComp = LeggiCampo("Select * from Compagnia Where IdCompagnia=" & IdCompagnia,"DescCompagnia" )
						%>
						<input type="text" readonly class="form-control" value="<%=DenomComp%>" >
					</div>		
				</div>
			</div>
			<div class="row">
			    <div class="col-1"></div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Importo Iniziale Utilizzato &euro;")
						%>
						<input type="text" readonly class="form-control"  value="<%=InsertPoint(ImptIniziale,2)%>" >
					</div>		
				</div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Importo Iniziale Svincolato &euro;")
						%>
						<input type="text" readonly class="form-control"  value="<%=InsertPoint(ImptStornato,2)%>" >
					</div>		
				</div>
			</div>			
<br>
			<div class="row">
			    <div class="col-1"></div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Importo Da Svincolare &euro;")
						NameLoaded  = NameLoaded & ";ImptSvincolo,FLP;"   
						NameRangeN  = "ImptSvincolo;ImptMassimo;0;99999999"
						ImptMassimo = cdbl(ImptIniziale) - cdbl(ImptStornato) + cdbl(ImptSvincolo)
						%>
						<input type="text" name="ImptSvincolo0"  id="ImptSvincolo0" class="form-control" value="<%=ImptSvincolo%>" >
						<input type="hidden" name="ImptMassimo0" id="ImptMassimo0"  value="<%=ImptMassimo%>" >						
					</div>		
				</div>
               <div class="col-2">
                  <%
                     kk="DataSvincolo" 
                     xx=ShowLabel("Data Svincolo")
                     NameLoaded= NameLoaded & kk & ",DTO;"  
                     cls="form-control"
                     if SolaLettura="" then 
                        cls = "form-control mydatepicker "" placeholder=""gg/mm/aaaa"" title=""formato : gg/mm/aaaa""" 
                     end if 
                     tmpDt = StoD(DataSvincolo)
                  %>
                 
                 <input type="text" <%=solaLettura%> class="<%=cls%>" Id="<%=KK%>0" name="<%=KK%>0" value="<%=tmpDt%>" >
                  
               </div>
			</div>
			<div class="row">
			    <div class="col-1"></div>
				<div class="col-10">
					<div class="form-group ">
						<%xx=ShowLabel("Descrizione Svincolo")
						NameLoaded= NameLoaded & ";DescSvincolo,TE"   
						%>
						<input type="text" name="DescSvincolo0" id="DescSvincolo0" class="form-control" value="<%=DescSvincolo%>" >
					</div>		
				</div>
            </Div>
			<br>
			<div class="row"><div class="mx-auto">
		             <%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
		            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		         </div>
			</div>			
        <br>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
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
