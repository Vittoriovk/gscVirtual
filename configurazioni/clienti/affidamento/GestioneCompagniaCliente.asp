<%
  NomePagina="GestioneCompagniaCliente.asp"
  titolo="Richiesta Di Affidamento Cliente"
  default_check_profile="Coll,Clie"
  act_call_firm = CryptAction("CALL_FIRM") 
  act_call_inte = CryptAction("CALL_INTE") 
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
function firma(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","<%=act_call_firm%>");
    document.Fdati.submit();  
}
function integra(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","<%=act_call_inte%>");
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
 
 'passo alla firma del documento
 if instr("CALL_FIRM CALL_INTE",Oper)>0 then 
    IdAffidamentoRichiestaComp = cdbl("0" & Request("ItemToRemove"))
    if Cdbl(IdAffidamentoRichiestaComp)>0 then 
       xx=RemoveSwap()
       Session("swap_IdAffidamentoRichiesta")     = IdAffidamentoRichiesta
       Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
       Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
       response.redirect     RitornaA("configurazioni/Clienti/affidamento/DocumentazioneRichiestaComp.asp")
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
                     DescNote = NoteAffidamentoComp
                     if isCliente() or IsCollaboratore() then 
                        DescNote = noteAffidamentoClie
                     end if 

                     
                     %>
                  <input type="textArea" readonly class="form-control" value="<%=DescNote%>" >      
                  </div>               
               </div> 
            </div> 

        <div class="table-responsive"><table class="table"><tbody>
        <thead>
        <tr>
            <th scope="col">Compagnia</th>
            <th scope="col">Stato</th>
            <th scope="col" width="12%" >Importo Affidato &euro;</th>
            <th scope="col" width="12%" >Max Impt.richiesta &euro;</th>
			<th scope="col" width="12%" >Valido Dal</th>
			<th scope="col" width="12%" >Valido Al</th>
            <%if cdbl(FlagStatoFinale)=0 then%>
            <th scope="col">Azioni</th>
            <%end if %>
        </tr>
        <%
        MySql = ""
        MySql = MySql & " Select A.*,B.DescCompagnia"
        MySql = MySql & " from AffidamentoRichiestaComp A,Compagnia B"
        MySql = MySql & " Where A.IdCompagnia = B.IdCompagnia "
        MySql = MySql & " And   A.IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
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
                ValoreAffidamento="n.d."
                MaxValorePolizza ="n.d."
				ValidoDal = ""
				if Cdbl(Rs("ValidoDal"))>0 then 
				   ValidoDal = StoD(Rs("ValidoDal"))
				end if 
				ValidoAl  = ""
				if Cdbl(Rs("ValidoAl"))>0 then 
				   ValidoAl  = StoD(Rs("ValidoAl"))
				end if 
                
                canFirma = false 
				canInteg = false 
                'controllo se esiste un affidamento in corso 
                'function in FunAffidamento
                IdAffiComp = cdbl("0" & RS("IdAffidamentoRichiestaComp"))
                IdStatoAffidamento = RS("IdStatoAffidamento")
                IdFlussoProcesso   = RS("IdFlussoProcesso")
                if instr("ANNU_AFFI_RIFI",IdStatoAffidamento)>0 then 
                   if IdStatoAffidamento="ANNU" then 
                       DescStato = "Annullata:" & GetDiz(DizAffComp,"NoteAffidamento")
                   end if 
                   if IdStatoAffidamento="AFFI" then 
                      DescStato = "Affidata"
                      ValoreAffidamento = InsertPoint(Rs("ValoreAffidamento"),2)
                      MaxValorePolizza  = InsertPoint(Rs("ImptSingolaPolizza"),2)
                   end if 
                   if IdStatoAffidamento="RIFI" then 
                      DescStato = "Rifiutata"
                   end if 
                else 
                   if IdFlussoProcesso="SEL_FORNITORE" then 
                      DescStato = "Gestione fornitore"
                   end if 
                   if IdFlussoProcesso="GES_FORNITORE" then 
                      DescStato = "Gestione fornitore"
                   end if 
                   if IdFlussoProcesso="CAM_FORNITORE" then 
                      DescStato = "Gestione fornitore"
                   end if                    
                   if IdFlussoProcesso="CAC_FORNITORE" then 
                      DescStato = "Gestione fornitore - firma trasferimento"
                      canFirma  = true 
                   end if                    
                   if IdFlussoProcesso="CAV_FORNITORE" then 
                      DescStato = "Gestione fornitore - Validazione firma trasferimento"
                      canFirma  = true 
                   end if                    
                end if 
                if IdFlussoProcesso="INC_FORNITORE" then 
                   DescStato = "Integrazione documentazione"
				   canInteg  = true 
                end if                 
        %>
            <tr scope="col">
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=Rs("DescCompagnia")%>">
                </td>
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=DescStato%>">
                </td> 
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=ValoreAffidamento%>">
                </td>         
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=MaxValorePolizza%>">
                </td>         
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=ValidoDal%>">
                </td>
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=ValidoAl%>">
                </td>				
                <%if cdbl(FlagStatoFinale)=0 then%>
                <td>
                 <%if canFirma then
                     RiferimentoA="col-2;#;;2;penn;Gestione Trasferimento Fornitore;;firma(" & IdAffiComp & ");N"%>
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                <%end if%>                                
                 <%if canInteg then
                     RiferimentoA="col-2;#;;2;penn;Integrazione Documenti;;integra(" & IdAffiComp & ");N"%>
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                <%end if%>  
                 <%if canFornitore then
                     RiferimentoA="col-2;#;;2;effe;Gestione Fornitore;;localForn(" & IdAffiComp & ");N"%>
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                <%end if%>                
                </td>
                <%end if %>
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
