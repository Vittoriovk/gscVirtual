<%
  'elenco delle richieste 
  NomePagina="ListaRichiesta.asp"
  titolo="Lista Richieste di Affidamento"
  default_check_profile="Coll,Clie,BackO"
  act_call_del  = CryptAction("CALL_DEL") 
  act_call_docu = CryptAction("CALL_DOCU") 
  act_call_dett = CryptAction("CALL_DETT") 
  act_call_inca = CryptAction("CALL_CARI") 
  act_call_comp = CryptAction("CALL_COMP") 
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

function localDel(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=myConfirmInfo("Affidamento","Richiesta di Annullamento","<%=act_call_del%>","Motivo dell'annullamento");

}
function localDett(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper",'<%=act_call_dett%>');
    document.Fdati.submit();
}
function localDocu(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper",'<%=act_call_docu%>');
    document.Fdati.submit();
}
function localCarico(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper",'<%=act_call_inca%>');
    document.Fdati.submit();
}
function localCompagnia(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper",'<%=act_call_comp%>');
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
   'dall'esterno arrivano filtri di ricerca
   TipoRicercaExt   = Session("swap_TipoRicercaExt")
   testo_ricercaExt = Session("swap_testo_ricercaExt")
   Session("swap_TipoRicercaExt") = ""
   Session("swap_testo_ricercaExt") = ""
   if TipoRicercaExt<>"" then 
      v_TipoRicerca = TipoRicercaExt
   end if 
   if testo_ricercaExt<>"" then 
      v_cercatesto = testo_ricercaExt
   end if 
   'response.write testo_ricercaExt & " - " & testo_ricerca
   'response.end 
   
else
   PaginaReturn           = getValueOfDic(Pagedic,"PaginaReturn")
end if 

xx=setValueOfDic(Pagedic,"PaginaReturn"           ,PaginaReturn)
xx=setCurrent(NomePagina,livelloPagina) 
Oper = DecryptAction(Oper)
if Oper="CALL_DOCU" then 
   xx=RemoveSwap()
   KK = Cdbl("0" & Request("ItemToRemove"))
   if KK>0 then 
      IdAff=LeggiCampo("select * from AffidamentoRichiestaComp where IdCompagnia=0 and IdAffidamentoRichiesta=" & KK,"IdAffidamentoRichiestaComp")
      Session("swap_IdAffidamentoRichiestaComp") = IdAff
      Session("swap_PaginaReturn")       = "configurazioni/Clienti/affidamento/" & NomePagina
      response.redirect RitornaA("configurazioni/Clienti/Affidamento/DocumentazioneRichiesta.asp")
      response.end 
   end if 
end if 

if Oper="CALL_DEL" then 
   'sto staccando un documento 
   xx=RemoveSwap()
   KK = Cdbl("0" & Request("ItemToRemove"))
   addInfo = Request("ItemToModify")
   if KK>0 then 
      qUpd = ""
      qUpd = qUpd & "update AffidamentoRichiestaComp "
      qUpd = qUpd & " Set IdStatoAffidamento = 'ANNU',DataChiusura= " & Dtos()
      qUpd = qUpd & " ,NoteAffidamento = '" & apici(addInfo) & "' "
      qUpd = qUpd & " where IdAffidamentoRichiestaComp in "
      qUpd = qUpd & " (Select IdAffidamentoRichiestaComp "
      qUpd = qUpd & "  from AffidamentoRichiestaComp "
      qUpd = qUpd & "  where IdAffidamentoRichiesta=" & KK & ")"
      'response.write qUpd
      ConnMsde.execute qUpd
  
      qUpd = ""
      qUpd = qUpd & "update AffidamentoRichiesta "
      qUpd = qUpd & " Set IdStatoAffidamento = 'ANNU' "
      qUpd = qUpd & " ,NoteAffidamento = '" & apici(addInfo) & "' "
      qUpd = qUpd & " where IdAffidamentoRichiesta = " & KK
      ConnMsde.execute qUpd
 
   end if 
end if 

if Oper="CALL_CARI" then 
   KK = Cdbl("0" & Request("ItemToRemove"))
   if Cdbl(KK)>0 then 
      qUpd = ""
      qUpd = qUpd & "update AffidamentoRichiesta "
      qUpd = qUpd & " Set IdStatoAffidamento = 'LAVO' "
      qUpd = qUpd & " ,IdAccountBackOffice = " & Session("LoginIdAccount")
      qUpd = qUpd & " ,IdFlussoProcessoBackO='CHECK_DOCU'"
      qUpd = qUpd & " ,IdFlussoProcessoCliente='CHECK_DOCU'"
      qUpd = qUpd & " where IdAffidamentoRichiesta = " & KK
      'response.write qUpd
      ConnMsde.execute qUpd
      qUpd = ""
      qUpd = qUpd & "update AffidamentoRichiestaComp "
      qUpd = qUpd & " Set IdStatoAffidamento = 'LAVO' "
      qUpd = qUpd & " ,IdFlussoProcesso='CHECK_DOCU'"
      qUpd = qUpd & " where IdAffidamentoRichiesta = " & KK
      qUpd = qUpd & " and IdCompagnia = 0 "
      'response.write qUpd 
      ConnMsde.execute qUpd   
   end if 
end if 
   
if Oper="CALL_COMP" then 
   xx=RemoveSwap()
   KK = Cdbl("0" & Request("ItemToRemove"))
   if KK>0 then 
      IdAff=LeggiCampo("select * from AffidamentoRichiestaComp where IdCompagnia=0 and IdAffidamentoRichiesta=" & KK,"IdAffidamentoRichiestaComp")
      Session("swap_IdAffidamentoRichiesta")     = KK
      Session("swap_IdAffidamentoRichiestaComp") = IdAff
      Session("swap_PaginaReturn")       = "configurazioni/Clienti/affidamento/" & NomePagina
      if IsBackOffice() then 
         response.redirect RitornaA("configurazioni/Clienti/Affidamento/GestioneCompagnia.asp")
      else
         response.redirect RitornaA("configurazioni/Clienti/Affidamento/GestioneCompagniaCliente.asp")
      end if 
      response.end 
   end if 
end if 
   Oggi = Dtos() 
   Set Rs = Server.CreateObject("ADODB.Recordset")

   DescLoaded=""
   NumRec  = 0
   MsgNoData  = ""
   
   IdCollaboratore = NumForDb("0" & Request("IdCollaboratore0"))
  
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
                <div class="col-11"><h3>Elenco Richieste Di Affidamento</h3>
                </div>
            </div>
            
    <%if IsCliente() = false then %>
    
       <div class="row">
           <div class="col-1 s1 no-margin font-weight-bold">Collaboratore</div>
           <div class="col-9 no-margin">
           <%
           stdClass="class='form-control form-control-sm'"
        
           if IsBackOffice() then 
              CondRef = "livello = 1"
           else 
              CondRef = getCondForLevel(session("LivelloAccount"),Session("LoginIdAccount"))
           end if    
           q = "Select * from Collaboratore Where " & CondRef 
           q=Q & " order by Denominazione "
           'Where 
           response.write ListaDbChangeCompleta(q,"IdCollaboratore0",IdCollaboratore ,"IdAccount","Denominazione" ,1,"","","","","",stdClass)
      
          %>
          </div>    
          <div class="col-2 no-margin"></div>    
    </div>
    <%end if %>
    
            <%
            AddRow=true
            dim CampoDb(10)
            ElencoOption = ";0;Denominazione Cliente;1;Codice Fiscale;2;Partita Iva;3"
            CampoDB(1)   = "c.Denominazione"
            CampoDB(2)   = "c.CodiceFiscale"
            CampoDB(3)   = "c.PartitaIva"
            
            %>
        <!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
        <!--#include virtual="/gscVirtual/include/ShowStatoFlusso.asp"-->

        <div class="table-responsive"><table class="table"><tbody>
        <thead>
        <tr>
            <th scope="col" width="12%" >Richiesta Del</th>
            <%if IsBackOffice() then %>
                 <th scope="col">Collaboratore</th>
                 <th scope="col">Cliente</th>
                 <th scope="col">Gestita Da</th>                 
            <%end if %>
            <%if IsCollaboratore() then %>
                 <th scope="col">Collaboratore</th>
                 <th scope="col">Cliente</th>
            <%end if %>
            <th scope="col">Stato Affidamento</th>
            <th scope="col">Azioni</th>
        </tr>
        </thead>
        
        <%
        'leggo lo stato di affidamento delle richieste 
        err.clear
        if Condizione<>"" then 
           Condizione = " and " & Condizione
        end if 

        MySql = ""
        MySql = MySql & " Select A.*,E.DescrizioneUtente,E.DescTipoStato,C.Denominazione,"
        MySql = MySql & " C.IdAccountLivello1,C.IdAccountLivello2,D.Denominazione as DescCollLev1 "
        MySql = MySql & " from AffidamentoRichiesta A,cliente C,Collaboratore D"
        MySql = MySql & " ,StatoFlusso E"
        MySql = MySql & " Where A.IdAccountCliente = C.IdAccount "
        MySql = MySql & " And   A.IdStatoAffidamento = E.IdStatoSorgente "
        MySql = MySql & " And  (E.IdTipoUtente = '*' or E.IdTipoUtente like'%" & Session("LoginTipoUtente") & "%' ) "
        if IsBackOffice() then 
           MySql = MySql & " And  (A.IdFlussoProcessoBackO = E.IdFlussoProcesso or E.IdFlussoProcesso='*'  ) "
		   MySql = MySql & " And A.IdStatoAffidamento<>'Compila' "
        else
           MySql = MySql & " And  (A.IdFlussoProcessoCliente = E.IdFlussoProcesso or E.IdFlussoProcesso='*') "
        end if 
        MySql = MySql & " And   C.IdAccountLivello1 = D.IdAccount "
        if IsCliente() then 
           MySql = MySql & " and   A.IdAccountCliente = " & Session("LoginIdAccount")
        else 
		   if IsCollaboratore() then 
              MySql = MySql & " And (C.IdAccountLivello1=" & Session("LoginIdAccount") & "  or  C.IdAccountLivello2=" & Session("LoginIdAccount") & ")"
           end if 
           if Cdbl(IdCollaboratore)>0 then 
              MySql = MySql & " And (C.IdAccountLivello1=" & IdCollaboratore & "  or  C.IdAccountLivello2=" & IdCollaboratore & ")"
           end if 
        end if 
        MySql = MySql & Condizione
        if IdTipoStato <> "TUTTI" then 
           MySql = MySql & " And E.IdTipoStato='" & IdTipoStato & "' "
        end if 
        
        MySql = MySql & " order by A.DataRichiesta Desc"
        Rs.CursorLocation = 3 
        Rs.Open MySql, ConnMsde
        'response.write MySql 
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
                Id=Rs("IdAffidamentoRichiesta")
                IdStato=ucase(Rs("IdStatoAffidamento"))
                if IsCliente() or IsCollaboratore() then 
                   idFlussoProcesso=Rs("IdFlussoProcessoCliente")
                else 
                   idFlussoProcesso=Rs("IdFlussoProcessoBackO")
                end if 
                StatoComp=funDoc_StatoDocum(Id)
           
        %>
            <tr scope="col">
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=StoD(Rs("DataRichiesta"))%>">
                </td>
                <%if isCollaboratore() then 
                     if cdbl(RS("IdAccountLivello2"))>0 and cdbl(RS("IdAccountLivello2")) <> cdbl(Session("LoginIdAccount")) then 
                        DescColl2 = "secondo"
                     else
                        DescColl2 = "n.d."
                     end if 
                %>
                
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=descColl2%>">
                </td>
                
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=Rs("Denominazione")%>">
                </td>
                <%end if %>

                <%if isBackOffice() then %>
                
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=Rs("DescCollLev1")%>">
                </td>
                
                <td>
                    <input class="form-control" type="text" readonly 
                    value="<%=Rs("Denominazione")%>">
                </td>
                <td>
                    <%
                    if Cdbl(Rs("IdAccountBackOffice"))=0 then 
                       GestitaDa="n.d."
                    else 
                       GestitaDa=LeggiCampo("Select * from Account Where idAccount=" & Rs("IdAccountBackOffice"),"Nominativo")
                    end if
                    
                    
                    %>
                    <input class="form-control" type="text" readonly 
                    value="<%=GestitaDa%>">
                </td>                
                <%end if %>
                 <td>
                   <input class="form-control" type="text" readonly value="<%=Rs("DescrizioneUtente") %>">
                 </td>
                 <td>
                    <%canDelete = false%>
                    
                    <%if IsBackOffice() then
                         showDocu = false 
                         showComp = false 
						 ShowDele = true 
                         if IdStato="AFFI" then 
                            showDocu = true
                            showComp = true  
							showDele = false  
                         end if 
                         if IdStato="ANNU" then 
                            showDocu = true
                            showComp = false  
							showDele = false  
                         end if 						 
                    %>
                         <%if Cdbl(Rs("IdAccountBackOffice"))=0 then
                             RiferimentoA=";#;;2;sele;Prendi in carico;;localCarico('" & Id & "');N"
                             %>
                             <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                         
                         <%end if %>
                         <%if IdStato="LAVO" and instr("CHECK_DOCU INT_DOCU",idFlussoProcesso)>0 then
                             showDocu = true 
                           end if%> 
                         <%if IdStato="LAVO" and instr("SEL_COMPAGNIA",idFlussoProcesso)>0 then
                             ShowDocu = true 
                             ShowComp = true
                           end if%> 
                         <%if showDocu then
                             canDelete = true
                             RiferimentoA="col-2;#;;2;docu;Gestisci Documenti;;localDocu(" & Id & ");N"
                             %>
                             <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                         <%end if%> 
                         <%if ShowComp then
                             canDelete = true
                             RiferimentoA="col-2;#;;2;copy;Gestione Compagnia;;localCompagnia(" & Id & ");N"
                             %>
                             <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                         <%end if%> 

                    <%end if %>

                    
                    <%if IsCliente() or IsCollaboratore() then %>
                         <%if len(IdStato)>0 and Instr("COMPILA",IdStato)>0 then
                             canDelete = true
                             RiferimentoA="col-2;#;;2;docu;Gestisci Documenti;;localDocu(" & Id & ");N"
                             %>
                             <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                         <%end if%> 
                         <%if len(IdStato)>0 and Instr("ANNU_RIFI",IdStato)>0 then
                             canDelete = true
                             RiferimentoA="col-2;#;;2;docu;Verifica Dettaglio;;localDocu(" & Id & ");N"
                             %>
                             <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                         <%end if%> 
						 
                         <%if IdStato="RICH" then
                             canDelete = true
                           end if      
                         %>
                         <%if IdStato="LAVO" and instr("SEL_COMPAGNIA CHECK_DOCU INT_DOCU",idFlussoProcesso)>0 then
                             canDelete = true
                             RiferimentoA="col-2;#;;2;docu;Vedi Documenti;;localDocu(" & Id & ");N"
                             %>
                             <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                         <%end if%> 
                         <%if IdStato="LAVO" and instr("SEL_COMPAGNIA",idFlussoProcesso)>0 then
                             canDelete = true
                             RiferimentoA="col-2;#;;2;copy;Gestione Compagnia;;localCompagnia(" & Id & ");N"
                             %>
                             <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                         <%end if%> 
                         <%if IdStato="AFFI" then
                             RiferimentoA="col-2;#;;2;copy;Gestione Compagnia;;localCompagnia(" & Id & ");N"
                             %>
                             <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                         <%end if%>                          
                    <%end if %>
                     <%if canDelete and ShowDele then
                         RiferimentoA="col-2;#;;2;dele;Annulla;;localDel(" & Id & ");N"%>
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

</tbody></table></div> <!-- table responsive fluid -->

            <input type="hidden" name="localVirtualPath" id="localVirtualPath" value = "<%=VirtualPath%>">

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
<%
if FirstLoad then 
	response.write "<script language=javascript>document.Fdati.submit();</script>" 
	response.end 
end if
%>
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
