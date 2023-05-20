<%
  NomePagina="SelezioneFornitore.asp"
  titolo="Selezione Fornitore Per Compagnia"
  default_check_profile="BackO"
  act_call_forn = CryptAction("CALL_FORN")  
  act_call_send = CryptAction("CALL_SEND")  
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
function registraForn()
{
    var of = ValoreDi("idFornitoreOld0");
	if (!(of=="0" || of=="-1" || of=="")) {
	      xx=myConfirm("Cambio Fornitore","Conferma operazione: i dati registrati saranno persi.","<%=act_call_forn%>");
		  return false;
    }
    xx=ImpostaValoreDi("Oper","<%=act_call_forn%>");
    document.Fdati.submit();  
}
function myConfirmYes()
  {
    var act = $("#myConfirmAction").val();
    ImpostaValoreDi("Oper",act);
    document.Fdati.submit();
  
}
function inviaForn()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_send%>");
    document.Fdati.submit();  
}
function creaZip(id)
{
   var vp=$("#hiddenVirtualPath").val();
   var dataIn="IdAffidamentoRichiestaComp=" + id;
   $.ajax({
      type: "POST",
      async: false,
      url: vp + "configurazioni/clienti/AffidamentoScaricaZip.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
        retVal = msg;
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        var descE = xhr.status + ":Chiamata esegui Upload fallita, si prega di riprovare..." + thrownError;
        alert(descE);
        retVal = "ERR:Chiamata Fallita"
      }
    });
    document.Fdati.submit();
}

function gesForn(oper)
{
   if (oper=='CONFERMA') {
      xx=ControllaAffidamento();
    
      if (xx==false)
         return false;
   } 
   if (oper=='ANNULLA') {
      xx=ImpostaValoreDi("NameLoaded","NewDescStato,TE;NewDescStatoClie,TE");
      xx=ImpostaValoreDi("DescLoaded","0");
      xx=ElaboraControlli();
    
      if (xx==false)
         return false;
   }   
  
   xx=ImpostaValoreDi("Oper",oper);
   document.Fdati.submit();
}
function ControllaAffidamento()
{
      xx=ImpostaValoreDi("NameLoaded","ValoreAffidamento,FLP;ImptSingolaPolizza,FLP;ValidoDalComp,DTO;ValidoAlComp,DTO");
      xx=ImpostaValoreDi("DescLoaded","0");
      xx=ElaboraControlli();
      if (xx==false)
         return false;

      return true; 
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
end if 

if cdbl("0" & IdAffidamentoRichiesta)=0 then 
   qSel = "Select * from AffidamentoRichiestaComp Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp 
   IdAffidamentoRichiesta = LeggiCampo(qSel,"IdAffidamentoRichiesta")
end if 

 'registrazione dei dati :
 deSt = trim(Request("NewDescStato0"))
 clSt = trim(Request("NewDescStatoClie0"))
 
 Oper = DecryptAction(Oper)
 
 'salva il fornitore 
 if Oper = "CALL_FORN" then 
    qSel = "Select * from AffidamentoRichiestaComp Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp 
    olSt = LeggiCampo(qSel,"IdFornitore")
    IdSt = Request("IdFornitore0")
	qUpd = ""
    qUpd = qUpd & " Update AffidamentoRichiesta set  "
    qUpd = qUpd & " IdFlussoProcessoCliente = 'SEL_COMPAGNIA'"
    qUpd = qUpd & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
	qUpd = qUpd & " and IdFlussoProcessoCliente = 'CHECK_DOCU'"
	'response.write qUpd 
	
	ConnMsde.execute qUpd
	qUpd = ""
    qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
    qUpd = qUpd & " IdFornitore = " & IdSt 
    qUpd = qUpd & ",NoteAffidamento = '" & apici(deSt) & "'"
    qUpd = qUpd & ",NoteAffidamentoCliente = '" & apici(clSt) & "'"
	if Cdbl(olSt)<>Cdbl(IdSt) then
	   qUpd = qUpd & ",PathDocumentoZip = ''"
	end if 
    qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    ConnMsde.execute qUpd
    if Cdbl(olSt)<>Cdbl(IdSt) then 
      'cancello i documenti caricati 
       qDel = "Delete From AffidamentoRichiestaCompDoc Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
       'response.write qdel 
       ConnMsde.execute qDel

       'devo inserire i documenti necessari 
       IdCompagnia = LeggiCampo(qSel,"IdCompagnia")
       IdAccForn   = LeggiCampo("select * from fornitore Where IdFornitore=" & IdSt ,"IdAccount")
       qDoc = ""
       qDoc = qdoc & " select distinct IdDocumento"  
       qDoc = qdoc & " From AccountProdottoDocAff a, Prodotto B"
       qDoc = qdoc & " Where A.IdProdotto = b.IdProdotto "
       qDoc = qdoc & " and   A.TipoDoc = 'AFFI' "
       qDoc = qdoc & " and   B.IdAnagServizio = 'CAUZ_PROV' "
       qDoc = qdoc & " and   B.IdCompagnia = " & IdCompagnia
       qDoc = qdoc & " and   A.IdAccount = " & IdAccForn 

       IdAffidamentoRichiestaComp0 = 0
       qSel = ""
       qSel = qSel & " select * from AffidamentoRichiestaComp "
       qSel = qSel & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
       qSel = qSel & " and IdCompagnia = 0 "
       IdAffidamentoRichiestaComp0 = LeggiCampo(qSel,"IdAffidamentoRichiestaComp")
       qIns = ""
       qIns = qIns & " INSERT INTO AffidamentoRichiestaCompDoc"
       qIns = qIns & " (IdAffidamentoRichiestaComp,IdDocumento,TipoRife"
       qIns = qIns & " ,IdRife,FlagObbligatorio,IdAccountDocumento,FlagDataScadenza)"
       qIns = qIns & " select " & IdAffidamentoRichiestaComp & " as IdAffidamentoRichiestaComp,IdDocumento,TipoRife"
       qIns = qIns & " ,IdRife,FlagObbligatorio,IdAccountDocumento,FlagDataScadenza"
       qIns = qIns & " from AffidamentoRichiestaCompDoc "
       qIns = qIns & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp0
       qIns = qIns & " and IdDocumento in ("  & qdoc & ")"
          
       ConnMsde.execute qIns 
        
       'response.write qIns 
    end if 
    
 end if 
 
 'segna come inviato
 if Oper = "CALL_SEND" then 
    qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
    qUpd = qUpd & " IdFlussoProcesso = 'GES_FORNITORE' "
    qUpd = qUpd & ",NoteAffidamento = '"    & apici(deSt) & "'"
    qUpd = qUpd & ",NoteAffidamentoCliente = '" & apici(clSt) & "'"
    qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    qUpd = qUpd & " and   IdFlussoProcesso = 'SEL_FORNITORE' "
    'response.write qUpd
    ConnMsde.execute qUpd
    xx=AggiornaRichiestaAffidamento(IdAffidamentoRichiesta)
    if IsBackOffice()=false then 
       response.redirect RitornaA(PaginaReturn)
    else 
       xx=RemoveSwap()
       Session("swap_IdAccountCliente")           = IdAccountCliente
       Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
       Session("swap_PaginaReturn")               = "configurazioni/Clienti/" & NomePagina
       response.redirect RitornaA("configurazioni/Clienti/Affidamento/GestioneCompagnia.asp")
       response.end 
    end if 
 end if  
 

   xx=setValueOfDic(Pagedic,"PaginaReturn"               ,PaginaReturn)
   xx=setValueOfDic(Pagedic,"IdAffidamentoRichiesta"     ,IdAffidamentoRichiesta)
   xx=setValueOfDic(Pagedic,"IdAffidamentoRichiestaComp" ,IdAffidamentoRichiestaComp)
   xx=setCurrent(NomePagina,livelloPagina) 

   'carico i dati della richiesta 
   Dim DizAffComp
   Set DizAffComp = CreateObject("Scripting.Dictionary")
   Esito = GetDettaglioAffComp(DizAffComp,IdAffidamentoRichiestaComp,0,0)

   qReq = "select * from AffidamentoRichiesta Where IdAffidamentoRichiesta=" & IdAffidamentoRichiesta
   IdAccountCliente  = LeggiCampo(qReq,"IdAccountCliente")
   DescCliente       = LeggiCampo("Select * from Account Where idAccount=" & IdAccountCliente,"Nominativo" )
   idCompagniaComp   = GetDiz(DizAffComp,"IdCompagnia")
   DescCompagniaComp = LeggiCampo("Select * from Compagnia Where idCompagnia=" & idCompagniaComp,"DescCompagnia")
   DataRichiestaComp = GetDiz(DizAffComp,"DataRichiesta")
   
   IdStatoAffidamentoComp = GetDiz(DizAffComp,"IdStatoAffidamento")
   IdFlussoProcesso       = GetDiz(DizAffComp,"IdFlussoProcesso")
   if IdFlussoProcesso = "SEL_FORNITORE" then 
      if cdbl(IdFornitore)=0 then 
         DescStatoAffidamentoComp = "Assegnazione Fornitore"
      else 
         DescStatoAffidamentoComp = "Attesa esito Fornitore"
      end if 
   else 
      DescStatoAffidamentoComp = ""
   end if 
   IdFornitore            = GetDiz(DizAffComp,"IdFornitore")
   if Cdbl(IdFornitore)>0 then 
      DescFornitore=LeggiCampo("Select * from Fornitore Where IdFornitore=" & Idfornitore,"DescFornitore")
   else
      DescFornitore=""
   end if    
   PathDocumentoZip    = GetDiz(DizAffComp,"PathDocumentoZip")
   NoteAffidamentoComp = GetDiz(DizAffComp,"NoteAffidamento")
   NoteAffidamentoClie = GetDiz(DizAffComp,"NoteAffidamentoCliente")
   DataChiusuraComp    = GetDiz(DizAffComp,"DataChiusura")

   ValoreAffidamento   = GetDiz(DizAffComp,"ValoreAffidamento")
   ValidoDalComp       = GetDiz(DizAffComp,"ValidoDal")
   ValidoAlComp        = GetDiz(DizAffComp,"ValidoAl")
   
   Oggi = Dtos() 
   Set Rs = Server.CreateObject("ADODB.Recordset")

   richiedimotivo=true 
   FlagStatoFinale = LeggiCampo("Select * from StatoAffidamento where IdStatoAffidamento='" & IdStatoAffidamentoComp & "'","FlagStatoFinale")
   if Cdbl("0"& flagStatoFinale)=1 then 
      richiedimotivo=false
   end if 
   DescLoaded = ""
   NumRec     = 0
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
                <% end if %>
                <div class="col-11"><h3>Dettaglio Richiesta Di Affidamento per compagnia/fornitore</h3>
                </div>
            </div>
            <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
            <div class="row">
               <div class="col-3">
                  <div class="form-group ">
                     <%xx=ShowLabel("Utente")%>
                     <input type="text" readonly class="form-control" value="<%=DescCliente%>" >
                  </div>        
               </div>
               <div class="col-3">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Compagnia")%>
                     <input type="text" readonly class="form-control" value="<%=DescCompagniaComp%>" >
                  </div>        
               </div>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Richiesta Del")%>
                     <input type="text" readonly class="form-control" value="<%=Stod(DataRichiestaComp)%>" >
                  </div>        
               </div>
               <div class="col-4">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Stato Richiesta")%>
                     <input type="text" readonly class="form-control" value="<%=DescStatoAffidamentoComp%>" >
                  </div>        
               </div>
               
            </div>
            
            <div class="row">
               <div class="col-3">
                  <div class="form-group">
                     <%xx=ShowLabel("Fornitore")%>
                     <input type="text" readonly class="form-control" value="<%=DescFornitore%>" >
                  </div>                       
               </div>             
               <div class="col-3">
                  <div class="form-group">
                     <%xx=ShowLabel("Elaborata il")%>
                     <input type="text" readonly class="form-control" value="<%=Stod(DataChiusuraComp)%>" >
                  </div>                       
               </div> 

               <div class="col-6">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Annotazioni")%>
                  <input type="textArea" readonly class="form-control" value="<%=NoteAffidamentoComp%>" >      
                  </div>               
               </div> 
            </div> 
            <%if richiedimotivo then %>
            <div class="row">
               <div class="col-12">
                  <div class="form-group ">
                     <%xx=ShowLabel("Annotazioni relative alla richiesta")%>
                     <input type="text" id="NewDescStato0" name="NewDescStato0" class="form-control" value="<%=NoteAffidamentoComp%>" >
                  </div>        
               </div>
            </div>
            <div class="row">
               <div class="col-12">
                  <div class="form-group ">
                     <%xx=ShowLabel("Annotazioni per il cliente")%>
                     <input type="text" id="NewDescStatoClie0" name="NewDescStatoClie0" class="form-control" value="<%=NoteAffidamentoClie%>" >
                  </div>        
               </div>
            </div>            
            
            <%end if %>
            <div class="row">
               <div class="col-4">
                  <div class="form-group ">
                  <%xx=ShowLabel("Fornitore")
                  qIn = ""
                  qin = qIn & " select B.IdAccountFornitore "
                  qin = qIn & " from Prodotto A, ProdottoAttivo B "
                  qin = qIn & " Where A.IdCompagnia = " & idCompagniaComp
                  qin = qIn & " and   A.IdProdotto = B.IdProdotto "
                     
                  stdClass="class='form-control form-control-sm'"
                  q = ""
                  q = q & " Select * from Fornitore "
                  q = q & " Where IdAccount in (" & qIn & ")"
                  tt=1
                  'blocco se è stata già gestita la richiesta 
                  if IdFlussoProcesso <> "SEL_FORNITORE" then 
                     q = q & " and IdFornitore =  " & IdFornitore
                     tt=0
                  end if 
                  q = q & " order by DescFornitore  "
                  'response.write q
                  response.write ListaDbChangeCompleta(q,"IdFornitore0",IdFornitore ,"IdFornitore","DescFornitore" ,tt,"","","","","",stdClass)
                  %>
                  </div>
               </div>
               <%if IdFlussoProcesso = "SEL_FORNITORE" then
			   %>
			    <input type="hidden" name="idFornitoreOld0" id="idFornitoreOld0" value="<%=cdbl("0" & IdFornitore)%>">
                <div class="col-2">
                   <div class="form-group ">
                   <br>
                       <button type="button" onclick="registraForn()" class="btn btn-success">Assegna Fornitore</button>
                   </div>
                </div> 
                <%end if %>
                <%if Cdbl(Idfornitore)>0 then %> 
                     <%if IdFlussoProcesso = "SEL_FORNITORE" then%>
                     <div class="col-2">
                          <div class="form-group ">
                          <br>
                          <button type="button" onclick="inviaForn()" class="btn btn-info">Invia Richiesta</button>
                          </div>
                     </div>
                     <%end if%>
                     <div class="col-2">
                         <div class="form-group ">
                         <br>
                         <%if PathDocumentoZip="" then 
                              DescInfo="Crea Zip"
                           else
                              DescInfo="Rigenera Zip"
                           end if 
                         %>
                            
                         <button type="button" onclick="creaZip(<%=IdAffidamentoRichiestaComp%>)" class="btn btn-warning"><%=DescInfo%></button>
                         <%Linkdocumento=PathDocumentoZip%>
                         <!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->                             
                      </div>
                   </div>
                <%end if %>


                <% if IdStatoAffidamentoComp="FORN" then %>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Importo Affidato")%>
                     <input type="text" class="form-control" Id="ValoreAffidamento0" name="ValoreAffidamento0" value="<%=ValoreAffidamento%>" >
                  </div>        
               </div>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Importo Max Cauzione")%>
                     <input type="text" class="form-control" Id="ImptSingolaPolizza0" name="ImptSingolaPolizza0" value="<%=ImptSingolaPolizza%>" >
                  </div>        
               </div>
               
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Valido Dal")
                     nome="ValidoDalComp0"
                     if ValidoDalComp>0 then 
                        valo=StoD(ValidoDalComp)
                     else
                        valo = StoD(DtoS())
                     end if
                     %>
                     <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                            class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >                         
                  </div>        
               </div>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Valido Al")
                     nome="ValidoAlComp0"
                     if ValidoAlComp>0 then 
                        valo=StoD(ValidoAlComp)
                     else
                        valo = StoD(Year(date()) & "1231")
                     end if
                     %>
                     <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                            class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >                         
                  </div>        
               </div>               
                <div class="col-2">
                   <div class="form-group ">
                     <%xx=ShowLabel("azioni")%>
                    <br>  
                   <%RiferimentoA="col-2;#;;2;hand;Conferma Affidamento;;gesForn('CONFERMA');N"%>  
                     <!--#include virtual="/gscVirtual/include/Anchor.asp"-->                          
                   </div>
                </div>  
                <% end if %>
            </div>        

<%
if Cdbl(IdFornitore)>0 then 
   'compagnia per cui affido
   IdAccount   = LeggiCampo("Select * from Fornitore Where IdFornitore=" & IdFornitore,"IdAccount")
   IdCompagnia = idCompagniaComp
   'cerco prodotto
   qProdCauzComp = " select IdProdotto from Prodotto where idAnagServizio = 'CAUZ_PROV' and IdCompagnia=" & IdCompagnia
   qProdCauzForn = ""
   qProdCauzForn = qProdCauzForn & " Select IdProdotto from AccountProdotto "
   qProdCauzForn = qProdCauzForn & " Where IdProdotto in (" & qProdCauzComp & ")"
   qProdCauzForn = qProdCauzForn & " and IdAccount = " & IdAccount
   
   'response.write qProdCauzForn 
   IdProdotto = LeggiCampo(qProdCauzForn,"IdProdotto")
 
   qListaDoc = ""
   qListaDoc = qListaDoc & " select A.*,B.*,C.DescDocumento "
   qListaDoc = qListaDoc & " from AccountProdottoDocAff A, AffidamentoRichiestaCompDoc B,Documento C  "
   qListaDoc = qListaDoc & " Where A.IdAccount = " & IdAccount
   qListaDoc = qListaDoc & " and   A.IdProdotto = " & IdProdotto
   qListaDoc = qListaDoc & " and   A.IdDocumento = B.IdDocumento "
   qListaDoc = qListaDoc & " and   A.IdDocumento = C.IdDocumento "
   qListaDoc = qListaDoc & " and   B.IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp

   qListaDoc = ""
   qListaDoc = qListaDoc & " select B.*,C.DescDocumento "
   qListaDoc = qListaDoc & " from  AffidamentoRichiestaCompDoc B,Documento C  "
   qListaDoc = qListaDoc & " Where B.IdDocumento = C.IdDocumento "
   qListaDoc = qListaDoc & " and   B.IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
   
   'response.write qListaDoc    
end if 
%>
            <!--#include virtual="/gscVirtual/configurazioni/clienti/Affidamento/FornitoreListaDocumenti.asp"-->
            
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
