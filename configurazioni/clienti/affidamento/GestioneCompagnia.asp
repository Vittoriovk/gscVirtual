<%
  NomePagina="GestioneCompagnia.asp"
  titolo="Richiesta Di Affidamento Cliente"
  default_check_profile="BackO"
  act_call_affi = CryptAction("CALL_AFFI")
  act_call_docu = CryptAction("CALL_DOCU")
  act_call_dele = CryptAction("CALL_DELE")
  act_call_rest = CryptAction("CALL_REST")
  act_call_forn = CryptAction("CALL_FORN") 
  act_call_forg = CryptAction("CALL_FORG") 
  act_call_fcam = CryptAction("CALL_FCAM") 
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

function localDocu(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","<%=act_call_docu%>");
    document.Fdati.submit();  
}

function localAffi(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","<%=act_call_affi%>");
    document.Fdati.submit();  
}
function localForn(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","<%=act_call_forn%>");
    document.Fdati.submit();  
}
function localForg(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","<%=act_call_forg%>");
    document.Fdati.submit();  
}
function localForc(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","<%=act_call_fcam%>");
    document.Fdati.submit();  
}

function localDele(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=myConfirmInfo("Affidamento Compagnia","Richiesta di Cancellazione","<%=act_call_dele%>","Motivo dell'annullamento");

}
function localRest(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=myConfirmInfo("Affidamento Compagnia","Richiesta di Recupero","<%=act_call_rest%>","Motivo dell'annullamento");

}
</script>

<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%
PaginaReturn = ""

if FirstLoad then 
   PaginaReturn               = getCurrentValueFor("PaginaReturn")
   IdAffidamentoRichiesta     = cdbl("0" & getCurrentValueFor("IdAffidamentoRichiesta"))
   IdAffidamentoRichiestaComp = cdbl("0" & getCurrentValueFor("IdAffidamentoRichiestaComp"))
else
   PaginaReturn               = getValueOfDic(Pagedic,"PaginaReturn")
   IdAffidamentoRichiesta     = getValueOfDic(Pagedic,"IdAffidamentoRichiesta")
   IdAffidamentoRichiestaComp = getValueOfDic(Pagedic,"IdAffidamentoRichiestaComp")
   IdAccountCliente           = getValueOfDic(Pagedic,"IdAccountCliente")
end if 

if Cdbl(IdAffidamentoRichiesta)=0 then 
   qSel = "Select * from AffidamentoRichiestaComp Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp 
   IdAffidamentoRichiesta=LeggiCampo(qSel,"IdAffidamentoRichiesta")
end if 
if Cdbl(IdAffidamentoRichiestaComp)=0 then 
   qSel = "Select * from AffidamentoRichiestaComp Where IdCompagnia=0 and IdAffidamentoRichiesta=" & IdAffidamentoRichiesta 
   IdAffidamentoRichiestaComp=LeggiCampo(qSel,"IdAffidamentoRichiesta")
end if 

if Cdbl(IdAffidamentoRichiesta)=0 and cdbl(IdAffidamentoRichiestaComp)=0 then 
   response.redirect PaginaReturn 
   response.end 
end if 

 if IsCliente() then 
    IdAccountCliente = Session("LoginIdAccount")
 else 
    IdAccountCliente = cdbl("0" & IdAccountCliente)
 end if 
 if cdbl(IdAccountCliente)=0 then 
    qSel = "Select * from AffidamentoRichiesta Where IdAffidamentoRichiesta=" & IdAffidamentoRichiesta 
    IdAccountCliente = LeggiCampo(qSel,"IdAccountCliente")
 end if 
 DescCliente              = LeggiCampo("Select * from Account Where idAccount=" & IdAccountCliente,"Nominativo" )

 'registrazione dei dati :
 
 xx=setValueOfDic(Pagedic,"PaginaReturn"               ,PaginaReturn)
 xx=setValueOfDic(Pagedic,"IdAccountCliente"           ,IdAccountCliente)
 xx=setValueOfDic(Pagedic,"IdAffidamentoRichiesta"     ,IdAffidamentoRichiesta) 
 xx=setValueOfDic(Pagedic,"IdAffidamentoRichiestaComp" ,IdAffidamentoRichiestaComp)
 xx=setCurrent(NomePagina,livelloPagina) 
 
 Set Rs = Server.CreateObject("ADODB.Recordset")
 Oggi = Dtos() 
 'eventuale messaggio da proporre a video
 infoProcesso = "" 
 Oper = DecryptAction(Oper)
 
 'carico una nuova richiesta
 if Oper="CALL_AFFI" and CheckTimePageLoad() then 
    IdCompagnia = cdbl("0" & Request("ItemToRemove"))
    if Cdbl(IdCompagnia)>0 then 
       MyQ = "" 
       MyQ = MyQ & "insert into AffidamentoRichiestaComp (IdAffidamentoRichiesta,IdAccountCliente,IdCompagnia"
       MyQ = MyQ & ",DataRichiesta,IdStatoAffidamento,IdFlussoProcesso)"
       MyQ = MyQ & " values (" & IdAffidamentoRichiesta & "," & IdAccountCliente & "," & IdCompagnia
       MyQ = MyQ & "," &  Dtos() & ",'LAVO','SEL_FORNITORE')"    
       ConnMsde.execute MyQ 
   
       'metto in lavorazione anche la richiesta padre 
       MyQ = "" 
       MyQ = MyQ & " update AffidamentoRichiesta "
       MyQ = MyQ & " Set IdStatoAffidamento='LAVO',DataChiusura=0"
       MyQ = MyQ & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
       ConnMsde.execute MyQ 
	   
       MyQ = "" 
       MyQ = MyQ & " update AffidamentoRichiesta "
       MyQ = MyQ & " Set IdFlussoProcessoBackO='SEL_FORNITORE'"
       MyQ = MyQ & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
	   MyQ = MyQ & " and IdFlussoProcessoBackO <> 'SEL_COMPAGNIA'"
       ConnMsde.execute MyQ 
   
    end if 
 end if 
 if Oper="CALL_DELE" and CheckTimePageLoad() then 
    IdCompagnia = cdbl("0" & Request("ItemToRemove"))
    if Cdbl(IdCompagnia)>0 then 
       MyQ = "" 
       MyQ = MyQ & " update AffidamentoRichiestaComp "
       MyQ = MyQ & " Set IdStatoAffidamento='ANNU'"
       MyQ = MyQ & " ,NoteAffidamento='" & apici(Request("myConfirmAddinfo")) & "'"
       MyQ = MyQ & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
       MyQ = MyQ & " And IdCompagnia = " & IdCompagnia
       ConnMsde.execute MyQ 
    end if 
 end if  
 if Oper="CALL_REST" and CheckTimePageLoad() then 
    IdCompagnia = cdbl("0" & Request("ItemToRemove"))
    if Cdbl(IdCompagnia)>0 then 
       MyQ = "" 
       MyQ = MyQ & " update AffidamentoRichiestaComp "
       MyQ = MyQ & " Set IdStatoAffidamento='LAVO'"
       MyQ = MyQ & " ,NoteAffidamento='" & apici(Request("myConfirmAddinfo")) & "'"
       MyQ = MyQ & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
       MyQ = MyQ & " And IdCompagnia = " & IdCompagnia
       ConnMsde.execute MyQ 
    end if 
 end if  
 'gestisco il fornitore per una compagnia 
 if Oper="CALL_FORN" then 
    IdAffidamentoRichiestaComp  = "0" & request("ItemToRemove")
    if Cdbl(IdAffidamentoRichiestaComp)>0 then 
       xx=RemoveSwap()
       Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
       Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
       response.redirect     RitornaA("configurazioni/Clienti/affidamento/SelezioneFornitore.asp")
       response.end 
    end if 
 end if 
 if Oper="CALL_DOCU" then 
    IdAffidamentoRichiestaComp  = "0" & request("ItemToRemove")
    if Cdbl(IdAffidamentoRichiestaComp)>0 then 
       xx=RemoveSwap()
       Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
       Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
       response.redirect     RitornaA("configurazioni/Clienti/affidamento/DocumentazioneRichiestaComp.asp")
       response.end 
    end if 
 end if  
 if Oper="CALL_FORG" then 
    IdAffidamentoRichiestaComp  = "0" & request("ItemToRemove")
    if Cdbl(IdAffidamentoRichiestaComp)>0 then 
       xx=RemoveSwap()
       Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
       Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
       response.redirect     RitornaA("configurazioni/Clienti/affidamento/GestioneFornitore.asp")
       response.end 
    end if 
 end if 
 if Oper="CALL_FCAM" then 
    IdAffidamentoRichiestaComp  = "0" & request("ItemToRemove")
    if Cdbl(IdAffidamentoRichiestaComp)>0 then 
       xx=RemoveSwap()
       Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
       Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
       response.redirect     RitornaA("configurazioni/Clienti/affidamento/TrasferisciFornitore.asp")
       response.end 
    end if 
 end if 
 
 
  'leggo lo stato generale dell'affidamento    
 MySql = ""
 MySql = MySql & " Select A.*,StFl.DescrizioneUtente as DescStatoAffidamento "
 MySql = MySql & " from AffidamentoRichiesta A," & getStatoFlusso() 
 MySql = MySql & " Where A.IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
 MySql = MySql & " and   A.IdStatoAffidamento = StFl.IdStatoSorgente"
 if IsBackOffice() then 
    MySql = MySql & " and StFl.IdFlussoProcesso  in ('*',A.IdFlussoProcessoBackO )"    
 else 
    MySql = MySql & " and StFl.IdFlussoProcesso  in ('*',A.IdFlussoProcessoCliente )"
 end if 
 'response.write MySql
 err.clear 

 Rs.CursorLocation = 3 
 Rs.Open MySql, ConnMsde
 IdAccountBackOffice = Rs("IdAccountBackOffice")
 DataRichiesta       = Rs("DataRichiesta")
 IdStatoAffidamento  = Rs("IdStatoAffidamento")
 qSel = "SELECT FlagStatoFinale from StatoAffidamento Where IdStatoAffidamento='" & apici(IdStatoAffidamento) & "'"
 FlagStatoFinale       = cdbl("0" & LeggiCampo(qSel,"FlagStatoFinale"))
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
                <%end if %>
                <div class="col-11">
                <h3>Gestione Compagnie per richiesta di affidamento</h3>
                </div>
            </div>
            <div class="row">
               <div class="col-3">
                  <div class="form-group ">
                     <%xx=ShowLabel("Utente")%>
                     <input type="text" readonly class="form-control" value="<%=DescCliente%>" >
                  </div>        
               </div>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Richiesta Del")%>
                     <input type="text" readonly class="form-control" value="<%=Stod(DataRichiesta)%>" >
                  </div>        
               </div>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Stato Richiesta")

                     %>
                     <input type="text" readonly class="form-control" value="<%=DescStatoAffidamento%>" >
                  </div>        
               </div>
               
            </div>
            
            <div class="row">

               <div class="col-10">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Annotazioni")
                     DescNote = NoteAffidamento
                     if isCliente() or IsCollaboratore() then 
                        DescNote = NoteAffidamento
                     end if 

                     
                     %>
                  <input type="textArea" readonly class="form-control" value="<%=DescNote%>" >      
                  </div>               
               </div> 
            </div> 

        <div class="table-responsive"><table class="table"><tbody>
        <thead>
        <tr>
            <th scope="col" width="12%" >Sel.</th>
            <th scope="col">Compagnia</th>
            <th scope="col">fornitore</th>
            <th scope="col">Stato</th>
            <th scope="col">Azioni</th>
        </tr>
        <%
        MySql = ""
        MySql = MySql & " Select distinct A.IdCompagnia,A.DescCompagnia"
        MySql = MySql & " from Compagnia A, v_prodottiAttivi B"
        MySql = MySql & " Where A.IdCompagnia = B.IdCompagnia "
        MySql = MySql & " order by DescCompagnia "

        Rs.CursorLocation = 3 
        Rs.Open MySql, ConnMsde
        %>
        <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
        <!--#include virtual="/gscVirtual/include/CheckRs.asp"-->
        <%
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
                Id=Rs("IdCompagnia")
                
                'controllo se esiste un affidamento in corso 
                'function in FunAffidamento
                Dim DizAffComp
                Set DizAffComp = CreateObject("Scripting.Dictionary")
                Esito = GetDettaglioAffComp(DizAffComp,0,IdAffidamentoRichiesta,Id)
                IdAffiComp = cdbl("0" & GetDiz(DizAffComp,"IdAffidamentoRichiestaComp"))
                IdStatoAffidamento = ""
                IdFlussoProcesso = ""
                DescStato        = "Da Affidare"
                IdFornitore      = 0 
                DescFornitore    = ""
                TrovatoAffi      = false 
                canDelete        = false  
                canRestore       = false 
                canFornitore     = false 
                canFornitoreGes  = false
                canFornitoreCam  = false 
                canFornitoreDoc  = false 
				canRinnovare     = false 
				FlagStatoFinale  = 0
                'ho trovato una richiesta in corso : trascivo 
                if cdbl(IdAffiComp)>0 then 
                   IdStatoAffidamento = GetDiz(DizAffComp,"IdStatoAffidamento")
                   IdFlussoProcesso   = GetDiz(DizAffComp,"IdFlussoProcesso")
                   IdFornitore        = GetDiz(DizAffComp,"IdFornitore")
				   'response.write IdStatoAffidamento
                   if instr("ANNU_AFFI_RIFI",IdStatoAffidamento)>0 then 
				      FlagStatoFinale = 1
                      if IdStatoAffidamento="ANNU" then 
                         DescStato = "Annullata:" & GetDiz(DizAffComp,"NoteAffidamento")
                         canRestore = true 
                      end if 
                      if IdStatoAffidamento="AFFI" then 
                         DescStato = "Affidata"
                         canFornitoreGes = true
						 canRinnovare    = true 
                      end if 
                      if IdStatoAffidamento="RIFI" then 
                         DescStato = "Rifiutata"
                         canRestore = true 
                      end if 
                   else 
                      if IdFlussoProcesso="SEL_FORNITORE" then 
                         DescStato = "Assegnazione fornitore"
                         canFornitore = true 
                         canDelete    = true 
                      end if 
                      if IdFlussoProcesso="GES_FORNITORE" then 
                         DescStato = "Gestione fornitore"
                         
                         canFornitoreGes = true 
                         canDelete       = true 
                      end if 
                      if instr("CAM_FORNITORE CAC_FORNITORE",IdFlussoProcesso)>0 then 
                         DescStato = "Cambio fornitore"
                         if instr("CAC_FORNITORE",IdFlussoProcesso)>0 then 
                            DescStato = "Cambio fornitore - attesa firma cliente"
                         end if 
                         canFornitoreCam = true
                         canFornitoreGes = true 
                         canDelete       = true 
                      end if                       
                      if instr("CAV_FORNITORE",IdFlussoProcesso)>0 then 
                         DescStato = "Cambio fornitore - validazione documento firmato"
                         canFornitoreDoc = true
                         canFornitoreGes = true 
                         canDelete       = true 
                      end if                       
                   end if 
                   if instr("INT_FORNITORE INC_FORNITORE",IdFlussoProcesso)>0 then 
                      DescStato = "Integrazione Documenti"
                      if instr("INC_FORNITORE",IdFlussoProcesso)>0 then 
                         DescStato = "Integrazione Documenti - attesa cliente"
                      end if 
					  canFornitoreDoc = true
                      canFornitoreGes = true 
                      canDelete       = true 
                   end if                    
                   
                   TrovatoAffi = true
                end if 
                
                
                'controllo se esiste un affidamento gia' effettuato 
                if TrovatoAffi=false then 
                end if 
           
                Trovato=""
                if TrovatoAffi=true then 
                   Trovato="SI"
                end if 
                if Cdbl(IdFornitore)>0 then 
                   DescFornitore=LeggiCampo("select * from fornitore Where IdFornitore=" & IdFornitore,"DescFornitore")
                end if 
                
                bgColor = ""
                if IdStatoAffidamento="AFFI" then  
                   bgcolor="bgcolor='#CAFFE0'"
                elseif instr("ANNU_RIFI",IdStatoAffidamento)>0 then 
                   bgcolor="bgcolor='#FF9A9A'"
                else 
                   bgcolor="bgcolor='#FFFFE0'"
                end if 
                                
        %>
            <tr scope="col" >
                <td><input class="form-control" type="text" readonly 
                    value="<%=Trovato%>">
                </td>            
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=Rs("DescCompagnia")%>">
                </td>
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=DescFornitore%>">
                </td>
                <td <%=bgColor%>>
                    <input class="form-control" type="text" readonly 
                    value="<%=DescStato%>">
                </td> 
                <td >
                <%if TrovatoAffi=false and cdbl(FlagStatoFinale)=0 then
                     RiferimentoA="col-2;#;;2;plus;Affida;;localAffi(" & Id & ");N"
                    %>
                    <!--#include virtual="/gscVirtual/include/Anchor.asp"--> 
                <%end if%>
                 <%if canFornitore then
                     RiferimentoA="col-2;#;;2;effe;Gestione Fornitore;;localForn(" & IdAffiComp & ");N"%>
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                <%end if%>                
                 <%if canFornitoreGes then
                     RiferimentoA="col-2;#;;2;effe;Gestione Fornitore;;localForg(" & IdAffiComp & ");N"%>
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                <%end if%>                                
                 <%if canFornitoreCam then
                     RiferimentoA="col-2;#;;2;upda;Cambio Fornitore;;localForc(" & IdAffiComp & ");N"%>
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                <%end if%>
                 <%if canFornitoreDoc then
                     RiferimentoA="col-2;#;;2;penn;Gestione Documenti;;localDocu(" & IdAffiComp & ");N"%>
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                <%end if%>
                 <%if canDelete then
                     RiferimentoA="col-2;#;;2;dele;Annulla;;localDele(" & Id & ");N"%>
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                <%end if%>
                 <%if canRestore and cdbl(FlagStatoFinale)=0 then
                     RiferimentoA="col-2;#;;2;logi;Recupera;;localRest(" & Id & ");N"%>
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                <%end if%>   
                <%if canRinnovare then
                     RiferimentoA="col-2;#;;2;erre;Rinnovare;;noaction();N"%>
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                <%end if%>   
                </td>
                
            </tr>
        <%    
        rs.MoveNext
    Loop
end if 
rs.close

%>

        
        
        
        </thead>            
        </tbody></table></div> <!-- table responsive fluid -->


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

</body>

</html>
