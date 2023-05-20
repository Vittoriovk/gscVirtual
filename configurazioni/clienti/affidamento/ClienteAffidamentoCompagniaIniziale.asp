<%
  NomePagina="ClienteAffidamentoCompagniaIniziale.asp"
  titolo="Affidamento cliente per compagnia"
  default_check_profile="BackO"
  act_call_upda = CryptAction("CALL_UPDA") 
  act_call_modi = CryptAction("CALL_MODI")
  act_call_dele = CryptAction("CALL_DELE")
  IdTipoSvincolo = "INIZIO"
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

function modify(id)
{
    ImpostaValoreDi("Oper","<%=act_call_modi%>");
    xx=ImpostaValoreDi("ItemToRemove",id);
    document.Fdati.submit();
}
function remove(id)
{
    ImpostaValoreDi("Oper","<%=act_call_dele%>");
    xx=ImpostaValoreDi("ItemToRemove",id);
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
   idCompagnia = 0
   ValidoDal = Dtos()
   ValidoAl =  year(date()) & "1231"
   ImptComplessivo    = 0
   ImptSingolaPolizza = 0
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
   
  'registro i dati della pagina 
   xx=setValueOfDic(Pagedic,"IdCliente"        ,IdCliente)
   xx=setValueOfDic(Pagedic,"IdAccount"        ,IdAccount)
   xx=setValueOfDic(Pagedic,"IdRecMod"         ,IdRecMod)
   xx=setValueOfDic(Pagedic,"DenomClie"        ,DenomClie)
   xx=setValueOfDic(Pagedic,"PaginaReturn"     ,PaginaReturn)
 
   xx=setCurrent(NomePagina,livelloPagina) 
   DescLoaded="0"  
  
   Oper = DecryptAction(Oper)
   'eseguo aggiornamento ma controllo i dati
   if ucase(oper)="CALL_UPDA" then 
	  ImptIniziale = Request("ImptIniziale0")
      qUpd = ""
	  qUpd = qUpd & " update AccountCreditoAffiTotali"
	  qUpd = qUpd & " set ImptIniziale=" & numForDb(ImptIniziale)
	  qUpd = qUpd & " Where IdAccount=" & IdAccount
	  'response.write qUpd
      ConnMsde.execute qUpd
   end if    
   if ucase(oper)="CALL_DELE" then 
      id = cdbl("0" & Request("ItemToRemove"))
	  if Cdbl(id)>0 then 
         qUpd = ""
	     qUpd = qUpd & " delete from AccountSvincolo"
	     qUpd = qUpd & " Where IdAccountSvincolo=" & Id
         'ConnMsde.execute qUpd
		 
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
   if Oper="CALL_MODI" then 
      id = cdbl("0" & Request("ItemToRemove")) 
      xx=RemoveSwap()
      Session("swap_IdAccount")      = IdAccount
	  Session("swap_IdTipoSvincolo") = IdTipoSvincolo
      Session("swap_IdRecMod")       = id
	  Session("swap_IdCompagnia")    = idCompagnia
	  Session("swap_IdCauzioneProv") = 0
	  Session("swap_IdCauzioneDefi") = 0
      Session("swap_PaginaReturn")   = "configurazioni/Clienti/Affidamento/" & NomePagina
      response.redirect RitornaA("configurazioni/Clienti/Affidamento/ClienteAffidamentoCompagniaInizialeMod.asp")
      response.end 
   end if 
  
   DescPageOper="Aggiornamento"

   Dim DizAff
   Set DizAff = CreateObject("Scripting.Dictionary")
   esito=GetTotaliAffidamentoComp(DizAff,IdAccount,IdCompagnia)
   ImptStornato = 0
   ImptIniziale = 0
   if Esito then 
      ImptStornato = cdbl(Getdiz(DizAff,"ImptInizialeStornato"))
	  ImptIniziale = cdbl(Getdiz(DizAff,"ImptIniziale"))
   end if 

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
				<div class="col-11"><h3>Gestione Cliente : Affidamento iniziale per compagnia </b></h3>
				</div>
			</div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

			<div class="row">
				<div class="col-5">
					<div class="form-group ">
						<%xx=ShowLabel("Utente")%>
						<input type="text" readonly class="form-control" value="<%=DenomClie%>" >
					</div>		
				</div>
				<div class="col-5">
					<div class="form-group ">
						<%xx=ShowLabel("Compagnia")
						DenomComp = LeggiCampo("Select * from Compagnia Where IdCompagnia=" & IdCompagnia,"DescCompagnia" )
						%>
						<input type="text" readonly class="form-control" value="<%=DenomComp%>" >
					</div>		
				</div>
			</div>
			<div class="row">
				<div class="col-5">
					<div class="form-group ">
						<%xx=ShowLabel("Fornitore")
						DenomComp = LeggiCampo("Select * from Fornitore Where IdFornitore=" & IdFornitore,"DescFornitore" )
						%>
						<input type="text" readonly class="form-control" value="<%=DenomComp%>" >
					</div>		
				</div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Importo Affidamento &euro;")
						%>
						<input type="text" readonly class="form-control"  value="<%=InsertPoint(ImptComplessivo,2)%>" >
					</div>		
				</div>
			</div>			

			<div class="row">
			    <div class="col-1"></div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Importo Iniziale Affidamento &euro;")
						NameLoaded= NameLoaded & ";ImptIniziale,FLP"   
						NameRangeN= "ImptStornato;ImptIniziale;0;99999999"
						%>
						<input type="text" name="ImptIniziale0" id="ImptIniziale0" class="form-control" value="<%=ImptIniziale%>" >
					</div>		
				</div>
				<div class="col-2">
					<div class="form-group ">
						<%xx=ShowLabel("Importo Svincolato &euro;")
						tt=InsertPoint(ImptStornato,2)
						%>
						<input type="text" readonly class="form-control" value="<%=tt%>" >
						<input type="hidden" id="ImptStornato0" name="ImptStornato0" value="<%=ImptStornato%>" >
					</div>		
				</div>
			</div>
			<div class="row"><div class="mx-auto">
		             <%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
		            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		         </div>
			</div>			
<br>
           <div class="table-responsive"><table class="table"><tbody>
            <thead>
                <tr>
                <th scope="col" width="11%">Data Svincolo
                <%RiferimentoA="col-2;#;;2;inse;Inserisci;;modify(0);N"%>
                  <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                </th>
                <th scope="col" >Descrizione</th>
                <th scope="col" width="11%">Importo Svincolato</th>
                <th scope="col" width="11%">azioni</th>
                </tr>
            </thead>
			<%
            err.clear
            MySql = ""
            MySql = MySql & " select * "
            MySql = MySql & " From AccountSvincolo"
            MySql = MySql & " Where IdAccount = " & IdAccount
            MySql = MySql & " and IdTipoSvincolo = '" & IdTipoSvincolo & "'"
            MySql = MySql & " order By DataSvincolo desc"


            Rs.CursorLocation = 3 
            Rs.Open MySql, ConnMsde
            DescLoaded=""
            NumCols = numC + 1
            NumRec  = 0
            ShowNew    = true
            ShowUpdate = false
            MsgNoData  = ""			
			%>
<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->  
            <%
            'elenco azioni 
            if MsgNoData="" and MsgErrore="" then 
            if PageSize>0 then 
                Rs.PageSize = PageSize
                pageTotali = rs.PageCount
                NumRec=0
                if Cpag<=0 then 
                    Cpag =1
                end if 
                if Cpag>PageTotali then 
                    CPag=PageTotali
                end if  
                Rs.absolutepage=CPag
            end if
            NumRec=0
            Primo=0
            Do While Not rs.EOF and (NumRec<PageSize or Pagesize<=0)
               Primo=Primo+1
               NumRec=NumRec+1
               Id=Rs("IdAccountSvincolo")
			   %>
               <tr scope="col">

                <td>
                    <input class="form-control" type="text" readonly value="<%=Stod(Rs("DataSvincolo"))%>">
                </td>
                <td>
                    <input class="form-control" type="text" readonly value="<%=Rs("DescSvincolo")%>">
                </td>
                <td>
                    <input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("ImptSvincolo"),2)%>">
                </td>                
                <td>
                  <%RiferimentoA=";#;;2;upda;Aggiorna;;modify(" & Id & ");N"%>
                    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                  <%RiferimentoA=";#;;2;dele;Cancella;;remove(" & Id & ");N"%>
                    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->				  
                </td>

            </tr>
            <%    
               rs.MoveNext
           Loop
           rs.close
           end if 
%>            
 
            </tbody></table></div>			

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
