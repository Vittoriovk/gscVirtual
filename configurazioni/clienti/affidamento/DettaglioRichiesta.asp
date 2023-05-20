<%
  NomePagina="DettaglioRichiesta.asp"
  titolo="Menu - Dettaglio Richiesta Di Affidamento Cliente"
  default_check_profile="Coll,Clie,BackO"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAffidamento.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionEvento.asp"-->
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
<script>


function callerCassettoSel(s)
{
    xx=ImpostaValoreDi("ItemToRemove",s);
    xx=ImpostaValoreDi("Oper","CALL_SEL_CAS");
    document.Fdati.submit();  
}

function localInsDoc(idDoc,FlagObbl,flagScad)
{
    xx=ImpostaValoreDi("ItemToRemove",idDoc + '_' + FlagObbl + '_' + flagScad);
    xx=ImpostaValoreDi("Oper","CALL_INS_DOC");
    document.Fdati.submit();  
}

function localSelCas(idDocumento,IdAccount,tipoRife,idRife)
{
    xx=popolaCassetto(IdAccount,idDocumento,tipoRife,idRife);
}

function localDel(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","CALL_DEL");
    document.Fdati.submit();  
}
function localUpd(idDoc,idAccDoc)
{
    xx=ImpostaValoreDi("ItemToRemove",idDoc + '_' + idAccDoc);
    xx=ImpostaValoreDi("Oper","CALL_UPD");
    document.Fdati.submit();  
}


function localCallComp(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","CALL_DETT_COMP");
    document.Fdati.submit();  
}
function localDelReqComp(id)
{
   if (confirm('Stai per rimuovere la richiesta per una compagnia, sei sicuro ?')) {
       xx=ImpostaValoreDi("ItemToRemove",id);
       xx=ImpostaValoreDi("Oper","CALL_DEL_COMP");
       document.Fdati.submit();  	
   }

}
function localSendReqComp(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","CALL_SEND_COMP");
    document.Fdati.submit();  	
}





</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%
PaginaReturn = ""

if FirstLoad then 
   PaginaReturn           = getValueOfDic(Pagedic,"PaginaReturn")
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
   IdAccountCliente       = Cdbl("0" & getValueOfDic(Pagedic,"IdAccountCliente"))
   if IdAccountCliente=0 then 
      IdAccountCliente       = Session("swap_IdAccountCliente") 
   end if 
   
   IdAffidamentoRichiesta = cdbl("0" & getValueOfDic(Pagedic,"IdAffidamentoRichiesta"))
   if IdAffidamentoRichiesta=0 then 
      IdAffidamentoRichiesta = Session("swap_IdAffidamentoRichiesta")
   end if 
else
   PaginaReturn           = getValueOfDic(Pagedic,"PaginaReturn")
   IdAccountCliente       = getValueOfDic(Pagedic,"IdAccountCliente")
   IdAffidamentoRichiesta = getValueOfDic(Pagedic,"IdAffidamentoRichiesta")
end if 

if PaginaReturn="" and Session("LoginTipoUtente")=ucase("Clie") then 
   PaginaReturn="link/ClienteAffidamento.asp"
end if 
if Session("LoginTipoUtente")=ucase("Clie") then 
   IdAccountCliente = Session("LoginIdAccount")
end if 

xx=setValueOfDic(Pagedic,"PaginaReturn"           ,PaginaReturn)
xx=setValueOfDic(Pagedic,"IdAccountCliente"       ,IdAccountCliente)
xx=setValueOfDic(Pagedic,"IdAffidamentoRichiesta" ,IdAffidamentoRichiesta)
xx=setCurrent(NomePagina,livelloPagina) 

MySql = "select * from Cliente Where IdAccount = " & IdAccountCliente 
Rs.CursorLocation = 3
Rs.Open MySql, ConnMsde
DescCliente = Rs("Denominazione")
cf          = Rs("CodiceFiscale")
PI          = Rs("PartitaIva")
Rs.close 

'selezionato dal cassetto lo associo 
if Oper="CALL_SEL_CAS" then 
   IdAccountDocumento = Cdbl("0" & Request("ItemToRemove"))
   
   if Cdbl(IdAccountDocumento) > 0 then 
      IdDocumento = LeggiCampo("Select * from AccountDocumento where IdAccountDocumento=" & IdAccountDocumento,"IdDocumento")
      qUpd = ""
	  qUpd = qUpd & " Update AffidamentoRichiestaCompDoc "
	  qUpd = qUpd & " Set IdAccountDocumento = " & IdAccountDocumento
	  qUpd = qUpd & " Where IdDocumento = " & IdDocumento
	  qUpd = qUpd & " and   IdAccountDocumento = 0 " 
	  qUpd = qUpd & " and   IdAffidamentoRichiestaComp in "
	  qUpd = qUpd & " (select IdAffidamentoRichiestaComp "
	  qUpd = qUpd & "  from AffidamentoRichiestaComp  "
	  qUpd = qUpd & "  where IdAffidamentoRichiesta=" & IdAffidamentoRichiesta & ")"
      ConnMsde.execute qUpd 
   end if 
end if 

if Oper="CALL_INS_DOC" then 
   tmp = Request("ItemToRemove")
   ptr = instr(tmp,"_")
   IdDocumento        = cdbl("0" & mid(tmp,1,ptr-1))
   tmp = mid(tmp,ptr+1)
   ptr = instr(tmp,"_")
   FlagObbl           = cdbl("0" & mid(tmp,1,ptr-1))
   FlagScad           = cdbl("0" & mid(tmp,ptr+1))
   'response.write tmp & " " & IdDocumento & " " & FlagObbl & " " & FlagScad 
   if IdDocumento > 0 then 
      xx=RemoveSwap()
      Session("swap_IdTabella")          = "CLIENTE_DOC"
      Session("swap_IdTabellaKeyInt")    = IdAccountCliente
      Session("swap_OperTabella")        = "CALL_INS"
      Session("swap_IdAccount")          = IdAccountCliente
      Session("swap_IdAccountDocumento") = 0
      Session("swap_PaginaReturn")       = "configurazioni/Clienti/" & NomePagina
      Session("swap_OperAmmesse")        = "CRUD"
	  Session("swap_ProcedureToCall")    = "AssegnaAffidamentoDoc " & IdAffidamentoRichiesta & ", '$Action$' , $IdAccountDocumento$ , $IdUpload$"
	  Session("swap_IdDocumentoToLoad")  = IdDocumento 
      response.redirect RitornaA("configurazioni/Clienti/DocumentoClienteUpload.asp")
      response.end 
   end if 
end if 

if Oper="CALL_UPD" then 
   tmp = Request("ItemToRemove")
   ptr = instr(tmp,"_")
   IdDocumento        = cdbl("0" & mid(tmp,1,ptr-1))
   IdAccountDocumento = cdbl("0" & mid(tmp,ptr+1))
   response.write tmp & " " & IdDocumento & " " & IdAccountDocumento
   if IdDocumento > 0 then 
      xx=RemoveSwap()
      Session("swap_IdTabella")          = "CLIENTE_DOC"
      Session("swap_IdTabellaKeyInt")    = IdAccountCliente
      Session("swap_OperTabella")        = Oper
      Session("swap_IdAccount")          = IdAccountCliente
      Session("swap_IdAccountDocumento") = IdAccountDocumento
      Session("swap_PaginaReturn")       = "configurazioni/Clienti/" & NomePagina
      Session("swap_OperAmmesse")        = "CRUD"
	  Session("swap_ProcedureToCall")    = "AssegnaAffidamentoDoc " & IdAffidamentoRichiesta & ", '$Action$' , $IdAccountDocumento$ , $IdUpload$"
	  Session("swap_IdDocumentoToLoad")  = IdDocumento 
      response.redirect RitornaA("configurazioni/Clienti/DocumentoClienteUpload.asp")
      response.end 
   end if 
   
end if 


if Oper="CALL_DEL" then 
   'sto staccando un documento 
   xx=RemoveSwap()
   KK = Cdbl("0" & Request("ItemToRemove"))
   if KK>0 then 
      qUpd = ""
      qUpd = qUpd & "update AffidamentoRichiestaCompDoc "
	  qUpd = qUpd & " Set IdAccountDocumento = 0 "
	  qUpd = qUpd & " where IdAccountDocumento = " & kk
	  qUpd = qUpd & " and   IdAffidamentoRichiestaComp in "
	  qUpd = qUpd & " (Select IdAffidamentoRichiestaComp "
	  qUpd = qUpd & "  from AffidamentoRichiestaComp "
	  qUpd = qUpd & "  where IdAffidamentoRichiesta=" & IdAffidamentoRichiesta & ")"
	  'response.write qUpd
	  ConnMsde.execute qUpd
   end if 
end if 

if Oper="CALL_DETT_COMP" then 
   xx=RemoveSwap()
   Session("swap_IdAffidamentoRichiestaComp") = Request("ItemToRemove")
   Session("swap_PaginaReturn")               = "configurazioni/clienti/" & NomePagina
   response.redirect RitornaA("configurazioni/Clienti/AffidamentoClienteDettaglioRichiestaComp.asp")
   Response.end 
end if 

if Oper="CALL_SEND_COMP" then 
   KK = Cdbl("0" & Request("ItemToRemove"))
   if kk>0 then 
      'aggiorno lo stato dei documenti per essere validati 
	  qSel = ""
	  qSel = qSel & " select IdAccountDocumento from AffidamentoRichiestaCompDoc "
	  qSel = qSel & " Where IdAffidamentoRichiestaComp=" & KK
	  qUpd = ""
	  qUpd = qUpd & " Update AccountDocumento "
	  qUpd = qUpd & " set IdTipoValidazione = 'DAVALI'"
	  qUpd = qUpd & " Where IdTipoValidazione='NONRIC'"
	  qUpd = qUpd & " and IdAccountDocumento in (" & qSel & ")"
	  err.clear 
	  ConnMsde.execute qUpd 
	  if err.number=0 then 
	     'cambio lo stato della richiesta 
		 qSel = "select * from AffidamentoRichiestaComp Where IdAffidamentoRichiestaComp=" & KK
		 lSta = LeggiCampo(qSel,"IdStatoAffidamento")
         qUpd = ""
         qUpd = qUpd & " Update AffidamentoRichiestaComp set "
		 qUpd = qUpd & "  DataRichiesta = " & Dtos()
		 qUpd = qUpd & " ,IdStatoAffidamento = 'LAVO' "
		 qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & KK
         ConnMsde.execute qUpd 
		 IdCompagnia = LeggiCampo("select * from AffidamentoRichiestaComp Where IdAffidamentoRichiestaComp=" & KK,"IdCompagnia")
		 DeCompagnia = LeggiCampo("select * from Compagnia Where IdCompagnia=" & IdCompagnia,"DescCompagnia")
		 'registo la richiesta
		 IdProdotto=GetProdottoByTipoComp("CAUZ_PROV",IdCompagnia)
		 infoReq = "Richiesta di affidamento per la Compagnia:" & DeCompagnia
		 infoSta = "LAVO"
		 if lSta="DOCU" then 
		    infoReq = "integrazione documentazione affidamento per la Compagnia:" & DeCompagnia
			infoSta = "INTE"
		 end if 
         XX=createEvento("AFFI",infoSta,Session("LoginIdAccount"),infoReq,"AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & KK,true,IdProdotto)
	  else  
	     MsgErrore=MsgErrore = ErroreDb(Err.description)
	  end if 
   end if 
end if 

if Oper="CALL_DEL_COMP" then 
   KK = Cdbl("0" & Request("ItemToRemove"))
   if kk>0 then 
      ConnMsde.execute "Delete From AffidamentoRichiestaCompDoc Where IdAffidamentoRichiestaComp=" & KK
	  ConnMsde.execute "Delete From AffidamentoRichiestaComp Where IdAffidamentoRichiestaComp=" & KK
   end if 
   qDel = "" 
   qDel = qDel & " Delete From AffidamentoRichiesta"
   qDel = qDel & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
   qDel = qDel & " and IdAffidamentoRichiesta not in "
   qDel = qDel & " ( select IdAffidamentoRichiesta From AffidamentoRichiestaComp "
   qDel = qDel & "   Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta & " ) "
   ConnMsde.execute qDel 
end if 

   Oggi = Dtos() 
   Set Rs = Server.CreateObject("ADODB.Recordset")
 
   MySql = ""
   MySql = MySql & " Select A.*,B.DescStatoAffidamento from AffidamentoRichiesta A,StatoAffidamento B"
   MySql = MySql & " Where A.IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
   MySql = MySql & " and   A.IdStatoAffidamento = B.IdStatoAffidamento"
   err.clear 

   Rs.CursorLocation = 3 
   Rs.Open MySql, ConnMsde
   IdAccountBackOffice = Rs("IdAccountBackOffice")
   DataRichiesta       = Rs("DataRichiesta")
   IdStatoAffidamento  = Rs("IdStatoAffidamento")
   DescStatoAffidamento= Rs("DescStatoAffidamento")
   DataChiusura        = Rs("DataChiusura")
   NoteAffidamento     = Rs("NoteAffidamento")			
   Rs.close 

   DescLoaded=""
   NumRec  = 0
   MsgNoData  = ""
  
%>

<div class="d-flex" id="wrapper">
    <%
      Session("opzioneSidebar")="affi"
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
                <%
                if PaginaReturn<>"" then 
                   RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"
                %>
                   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <% end if %>
                <div class="col-11"><h3>Dettaglio Richiesta Di Affidamento</h3>
                </div>
            </div>
			<% if isCliente()=false then %>
			<div class="row">
			   <div class="col-4">
                  <div class="form-group ">
				     <%xx=ShowLabel("Cliente")%>
					 <input type="text" readonly class="form-control" value="<%=DescCliente%>" >
                  </div>		
			   </div>			
			   <div class="col-2">
                  <div class="form-group ">
				     <%xx=ShowLabel("Cod.fiscale")%>
					 <input type="text" readonly class="form-control" value="<%=cf%>" >
                  </div>		
			   </div>	  
			   <div class="col-2">
                  <div class="form-group ">
				     <%xx=ShowLabel("Partita Iva")%>
					 <input type="text" readonly class="form-control" value="<%=pi%>" >
                  </div>		
			   </div>			   
			</div>
			<% end if %>
            <div class="row">
			   <div class="col-2">
                  <div class="form-group ">
				     <%xx=ShowLabel("Richiesta Del")%>
					 <input type="text" readonly class="form-control" value="<%=Stod(DataRichiesta)%>" >
                  </div>		
			   </div>						
			   <div class="col-2">
                  <div class="form-group ">
				     <%xx=ShowLabel("Stato Richiesta")%>
					 <input type="text" readonly class="form-control" value="<%=DescStatoAffidamento%>" >
                  </div>		
			   </div>	
			   <div class="col-2">
                  <div class="form-group ">
				     <%xx=ShowLabel("Elaborata Il")%>
					 <input type="text" readonly class="form-control" value="<%=Stod(DataChiusura)%>" >
                  </div>		
			   </div>	 
			   <div class="col-6">
                  <div class="form-group ">
				     <%xx=ShowLabel("Annotazioni")%>
					 <input type="text" readonly class="form-control" value="<%=NoteAffidamento%>" >
                  </div>		
			   </div>				
            </div> 



        <div class="table-responsive"><table class="table"><tbody>
        <thead>
        <tr>
            <th scope="col">Compagnia</th>
            <th scope="col">Stato Affidamento</th>
            <th scope="col">Azioni</th>
        </tr>
        </thead>
		
		<%
		'leggo le compagnie in affidamento 
		err.clear
        MySql = ""
        MySql = MySql & " Select A.*,B.DescStatoAffidamento,C.DescCompagnia "
        MySql = MySql & " from AffidamentoRichiestaComp A,StatoAffidamento B,Compagnia C"
        MySql = MySql & " Where A.IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
        MySql = MySql & " and   A.IdStatoAffidamento = B.IdStatoAffidamento"		
		MySql = MySql & " and   A.IdCompagnia = C.IdCompagnia"
        MySql = MySql & " order by C.DescCompagnia"
        Rs.CursorLocation = 3 
        Rs.Open MySql, ConnMsde
   
		%>
<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->
<%
        FlagRichiedi=false
        if MsgNoData="" then 
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
                Id=Rs("IdAffidamentoRichiestaComp")
				IdStato=Rs("IdStatoAffidamento")
				StatoComp=funDoc_StatoDocum(Id)
           
        %>
            <tr scope="col">
                <td>
                    <input class="form-control" type="text" readonly value="<%=Rs("DescCompagnia")%>">
                </td>
                 <%

                 %>
                 <td>
                   <input class="form-control" type="text" readonly value="<%=Rs("DescStatoAffidamento")%>">
                 </td>
                 <td>
					<%RiferimentoA="col-2;#;;2;dett;Dettaglio;;localCallComp(" & Id & ");N"
					%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					<%if false and (StatoComp="COMPL" or StatoComp="OK") and (IdStato="" or IdStato="RICH" or IdStato="DOCU") then 
					     RiferimentoA="col-2;#;;2;ok;Invia;;localSendReqComp(" & Id & ");N"
					%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					<%end if %>
					
					<%if IdStato="" or IdStato="RICH" or IdStato="DOCU" then 
					     RiferimentoA="col-2;#;;2;dele;Annulla;;localDelReqComp(" & Id & ");N"
					%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					<%end if %>
				</td>
            </tr>
        <%    
        rs.MoveNext
    Loop
end if 
rs.close

%>

</tbody></table></div> <!-- table responsive fluid -->
<%
IdCompagnia=0
ShowAction=true
OpDocAmm="IUD"
FunToCall="Local"

if false then 
%>
<!--#include virtual="/gscVirtual/configurazioni/clienti/AffidamentoClienteDocumentiLista.asp"-->

<%
end if 

if FlagRichiedi=true then 
%>
		<div class="row"><div class="mx-auto">
		<%
		RiferimentoA="center;#;;2;save;Registra; Registra;localIns();S"
		%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div></div>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
		</div>
		<br>
<%
end if 
%>
            <input type="hidden" name="localVirtualPath" id="localVirtualPath" value = "<%=VirtualPath%>">
            <!--#include virtual="/gscVirtual/utility/SelezioneDaCassetto.asp"-->

            <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
            <!--#include virtual="/gscVirtual/include/paginazione.asp"-->

			
            </form>
        </div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->

<!--#include virtual="/gscVirtual/include/scripts.asp"-->

  <!-- Menu Toggle Script -->
  <script>
    $("#menu-toggle").click(function(e) {
      e.preventDefault();
      $("#wrapper").toggleClass("toggled");
    });
  </script>
  <script>
    $(document).ready(function(){
      $('[data-toggle="tooltip" = Rs("")').tooltip();   
    });
  </script>
  <script>
$('.btn').onClick(function(e){
  e.preventDefault();
});  
</script>
</body>

</html>
