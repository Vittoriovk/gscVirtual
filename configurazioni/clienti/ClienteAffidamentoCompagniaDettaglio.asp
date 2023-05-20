<%
  NomePagina="ClienteAffidamentoCompagniaDettaglio.asp"
  titolo="Affidamento cliente per compagnia : dettaglio"
 
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
  Set Rs = Server.CreateObject("ADODB.Recordset")

  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdCliente=0
  IdRecMod =0
  if FirstLoad then 
	 IdCliente   = "0" & Session("swap_IdCliente")
	 if Cdbl(IdCliente)=0 then 
		IdCliente = cdbl("0" & getValueOfDic(Pagedic,"IdCliente"))
	 end if 
	 IdRecMod   = "0" & Session("swap_IdRecMod")
	 if Cdbl(IdRecMod)=0 then 
		IdRecMod = cdbl("0" & getValueOfDic(Pagedic,"IdRecMod"))
	 end if 
      
	 
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	 if PaginaReturn="" then 
		PaginaReturn = Session("swap_PaginaReturn")
	 end if 

  else
	 IdAccount = "0" & getValueOfDic(Pagedic,"IdAccount")
	 DenomClie =       getValueOfDic(Pagedic,"DenomClie")
	 IdCliente = "0" & getValueOfDic(Pagedic,"IdCliente")
	 IdRecMod  = "0" & getValueOfDic(Pagedic,"IdRecMod")
	 PaginaReturn    = getValueOfDic(Pagedic,"PaginaReturn")
   end if 
   IdCliente = cdbl(IdCliente)
   IdAccount = cdbl(IdAccount)
   if Cdbl(IdAccount)=0 then 
      MySql = "Select * from Cliente Where IdCliente=" & IdCliente 
      Rs.CursorLocation = 3
      Rs.Open MySql, ConnMsde  
	  IdAccount=Rs("IdAccount")
	  DenomClie=Rs("Denominazione")
	  Rs.close 
   end if 
   
   IdRecMod = Cdbl(IdRecMod)
  'registro i dati della pagina 
   xx=setValueOfDic(Pagedic,"IdCliente"        ,IdCliente)
   xx=setValueOfDic(Pagedic,"IdAccount"        ,IdAccount)
   xx=setValueOfDic(Pagedic,"IdRecMod"         ,IdRecMod)
   xx=setValueOfDic(Pagedic,"DenomClie"        ,DenomClie)
   xx=setValueOfDic(Pagedic,"PaginaReturn"     ,PaginaReturn)
 
   xx=setCurrent(NomePagina,livelloPagina) 
   DescLoaded="0"  
  
   'eseguo aggiornamento ma controllo i dati
   if ucase(oper)="UPDATE" then 
      ImptMinimo = 0 
	  IdCompagnia = Cdbl("0" & Request("IdCompagnia0"))
	  IdFornitore = Cdbl("0" & Request("IdFornitore0"))
      ImptComplessivo    = "0" & request("ImptComplessivo0")
      ImptSingolaPolizza = "0" & request("ImptSingolaPolizza0")
	  AffidamentoUsato   = 0
	  
      ValidoDal=DataStringa(Request("ValidoDal0"))
	  ValidoAl =DataStringa(Request("ValidoAl0"))
  
      msgErrore=caricaImportoAffidamento("V",IdAccount,IdCompagnia,IdFornitore,IdRecMod,ValidoDal,ValidoAl,ImptComplessivo,ImptSingolaPolizza,AffidamentoUsato)
      if msgErrore="" then 
         msgErrore=caricaImportoAffidamento("I",IdAccount,IdCompagnia,IdFornitore,IdRecMod,ValidoDal,ValidoAl,ImptComplessivo,ImptSingolaPolizza,AffidamentoUsato)
      end if 
      
   end if    
  
   idCompagnia = 0
   ValidoDal = Dtos()
   ValidoAl =  year(date()) & "1231"
   ImptComplessivo    = 0
   ImptSingolaPolizza = 0
   ImptImpegnato      = 0
   ImptDisponibile    = 0
   if cdbl(IdRecMod)>0 then 
      MySql = "select * from AccountCreditoAffi Where IdAccountCreditoAffi=" & IdRecMod
      Rs.CursorLocation = 3
      Rs.Open MySql, ConnMsde 
      if rs.eof = false then 
         idCompagnia     = rs("idCompagnia")
		 idFornitore     = rs("idFornitore")
         ValidoDal       = rs("ValidoDal")
         ValidoAl        = rs("ValidoAl")
         ImptComplessivo          = rs("ImptComplessivo")
		 ImptSingolaPolizza       = rs("ImptSingolaPolizza")
	  end if 
	  rs.close 
   end if 
   if cdbl(IdRecMod)=0 then 
      DescPageOper="Inserimento"
   else
      DescPageOper="Aggiornamento"
   end if
   Dim DizAff
   Set DizAff = CreateObject("Scripting.Dictionary")
   esito=GetTotaliAffidamento(DizAff,IdAccount,IdCompagnia)
   ImptImpegnato = 0
   if Esito then 
      ImptImpegnato = cdbl(Getdiz(DizAff,"TotaleImpegnato"))
   end if 
   ImptDisponibile = Cdbl(ImptComplessivo) - cdbl(ImptImpegnato)

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
				<div class="col-11"><h3>Gestione Cliente : Dettaglio affidamento per compagnia </b></h3>
				</div>
			</div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

			<div class="row">
				<div class="col-6">
					<div class="form-group ">
						<%xx=ShowLabel("Utente")%>
						<input type="text" readonly class="form-control" value="<%=DenomClie%>" >
					</div>		
				</div>

			</div>
			<div class="row">
				<div class="col-6">
					<div class="form-group ">
						<%xx=ShowLabel("Compagnia")

						stdClass="class='form-control form-control-sm'"
						
						q="Select * from Compagnia "
						if Cdbl(IdCompagnia)>0 and cdbl(ImptDisponibile)<> cdbl(ImptComplessivo) then 
						   q=Q & " Where IdCompagnia=" & IdCompagnia 
						end if 
						q=Q & " order by DescCompagnia "
						'Where 
						response.write ListaDbChangeCompleta(q,"IdCompagnia0",IdCompagnia ,"IdCompagnia","DescCompagnia" ,0,"","","","","",stdClass)
                        %>
					</div>		
				</div>			
			</div>
			<div class="row">
				<div class="col-6">
					<div class="form-group ">
						<%xx=ShowLabel("Fornitore")

						stdClass="class='form-control form-control-sm'"
						
						q="Select * from Fornitore "
						if Cdbl(IdFornitore)>0 and cdbl(ImptDisponibile)<> cdbl(ImptComplessivo) then 
						   q=Q & " Where IdFornitore=" & IdFornitore 
						end if 
						q=Q & " order by DescFornitore "
						'Where 
						response.write ListaDbChangeCompleta(q,"IdFornitore0",IdFornitore ,"IdFornitore","DescFornitore" ,0,"","","","","",stdClass)
                        %>
					</div>		
				</div>			
			</div>			
			<div class="row">
			    <div class="col-1"></div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Importo Affidamento &euro;")
						NameLoaded= NameLoaded & ";ImptComplessivo,FLP"   
						NameRangeN= "ImptSingolaPolizza;ImptComplessivo;0;99999999"
						%>
						<input type="text" name="ImptComplessivo0" id="ImptComplessivo0" class="form-control" value="<%=ImptComplessivo%>" >
					</div>		
				</div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Massimo Singola Polizza &euro;")
						NameLoaded= NameLoaded & ";ImptSingolaPolizza,FLP"   
						%>
						<input type="text" name="ImptSingolaPolizza0" id="ImptSingolaPolizza0" class="form-control" value="<%=ImptSingolaPolizza%>" >
					</div>		
				</div>
			</div>

			
			<div class="row">
			    <div class="col-1"></div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Importo Impegnato &euro;")
						tt=InsertPoint(ImptImpegnato,2)
						%>
						<input type="text" readonly class="form-control" value="<%=tt%>" >
					</div>		
				</div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Importo Disponibile &euro;")
						tt=InsertPoint(ImptDisponibile,2)
						%>
						<input type="text" readonly class="form-control" value="<%=tt%>" &euro;>
					</div>		
				</div>				
			</div>
			<div class="row">
			    <div class="col-1"></div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Valido Dal")
                        NameLoaded= NameLoaded & ";ValidoDal,DTO"   		  
		                nome="ValidoDal0"  
		                valo=StoD(ValidoDal)
		                %>	  
	                     <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >	
					</div>		
				</div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Valido Al")
                        NameLoaded= NameLoaded & ";ValidoAl,DTO"   		  
		                nome="ValidoAl0"  
		                valo=StoD(ValidoAl)
		                %>	  
	                     <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >	
					</div>		
				</div>
		
			</div>	
			<div class="row"><div class="mx-auto">
		             <%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
		            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		         </div></div>
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
