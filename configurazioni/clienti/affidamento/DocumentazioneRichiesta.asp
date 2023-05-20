<%
  NomePagina="DocumentazioneRichiesta.asp"
  titolo="Richiesta Di Affidamento Cliente"
  default_check_profile="Coll,Clie,BackO"
  act_call_send  = CryptAction("CALL_SEND") 
  act_call_addd  = CryptAction("CALL_ADDD") 
  act_call_gest  = CryptAction("CALL_GEST") 
  act_call_vali  = CryptAction("CALL_VALI") 
  act_call_regi  = CryptAction("CALL_REGI")
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
function localSend()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_send%>");
    document.Fdati.submit();

}
function LocalAddRow()
{
    xx=$('#confirmModal').modal('toggle');
}
function LocalAddRowNew()
{
    var sel = $('input[name="gruppo1"]:checked').val();
    if(typeof sel != 'undefined') 
    {
        xx=ImpostaValoreDi("Oper","<%=act_call_addd%>");
        document.Fdati.submit();
    }
    else
        alert("selezionare un documento");
}
function localGes(idDoc,tipoRife,idRife)
{
    xx=ImpostaValoreDi("ItemToRemove",idDoc);
    xx=ImpostaValoreDi("ItemToModify",tipoRife);
    xx=ImpostaValoreDi("IdParm",idRife);
    xx=ImpostaValoreDi("Oper","<%=act_call_gest%>");
    document.Fdati.submit();  
}
function localVal(idDoc)
{
    xx=ImpostaValoreDi("ItemToRemove",idDoc);
    xx=ImpostaValoreDi("Oper","<%=act_call_vali%>");
    document.Fdati.submit();  
}

function registraStato(stato)
{
    if (stato=="ANNU" || stato=="DOCU" || stato=="RIFI") {
       oldN=ValoreDi("NameLoaded");
       oldD=ValoreDi("DescLoaded");
       yy=ImpostaValoreDi("NameLoaded","NewDescStato,TE;NewDescStatoClie,TE");
       yy=ImpostaValoreDi("DescLoaded",'0');
       xx=ElaboraControlli();
       yy=ImpostaValoreDi("NameLoaded",oldN);
       yy=ImpostaValoreDi("DescLoaded",oldD);
       if (xx==false)
          return false;
    }
    xx=ImpostaValoreDi("Oper","<%=act_call_regi%>");
    xx=ImpostaValoreDi("IdParm",stato);
    document.Fdati.submit();  
}

</script>

<script>
function localSelCas(idDocumento,idAccount,tipoRife,idRife)
{
    xx=popolaCassetto(idAccount,idDocumento,tipoRife,idRife);
}




function callerCassettoSel(s)
{
    xx=ImpostaValoreDi("ItemToRemove",s);
    xx=ImpostaValoreDi("Oper","CALL_SEL_CAS");
    document.Fdati.submit();  
}


function localDel(idDoc)
{
    xx=ImpostaValoreDi("ItemToRemove",idDoc);
    xx=ImpostaValoreDi("Oper","CALL_DEL");
    document.Fdati.submit();  
}

function LocalAddRowCoob()
{
    xx=$('#confirmModalCoob').modal('toggle');
}
function LocalAddRowCert()
{
    xx=$('#confirmModalCert').modal('toggle');
}

function LocalRemRowCoob(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","CALL_REM_COOB");
    document.Fdati.submit();
}

function LocalRemRowCert(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=ImpostaValoreDi("Oper","CALL_REM_CERT");
    document.Fdati.submit();
}

function LocalAddRowNewCoob()
{
    var sel = $('input[name="gruppo1Coob"]:checked').val();
    if(typeof sel != 'undefined') 
    {
        xx=ImpostaValoreDi("Oper","CALL_ADD_COOB");
        document.Fdati.submit();
    }
    else
        alert("selezionare un coobbligato");
}
function LocalAddRowNewCert()
{
    var sel = $('input[name="gruppo1Cert"]:checked').val();
    if(typeof sel != 'undefined') 
    {
        xx=ImpostaValoreDi("Oper","CALL_ADD_CERT");
        document.Fdati.submit();
    }
    else
        alert("selezionare un certificato");
}
function reinviaGes(idDoc)
{
    xx=ImpostaValoreDi("ItemToRemove",idDoc);
    xx=ImpostaValoreDi("Oper","CALL_REI");
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
end if 

if PaginaReturn="" and IsCliente() then 
   PaginaReturn="link/ClienteAffidamento.asp"
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
   IdAccountCliente = 0
end if 
if cdbl(IdAccountCliente)=0 then 
   qSel = "Select * from AffidamentoRichiesta Where IdAffidamentoRichiesta=" & IdAffidamentoRichiesta 
   IdAccountCliente = LeggiCampo(qSel,"IdAccountCliente")
end if 

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
 
 'invio della richiesta 
 if Oper="CALL_SEND" then 
    MySql = ""
    MySql = MySql & " Select * "
    MySql = MySql & " from AffidamentoRichiesta "
    MySql = MySql & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
    IdStatoAffidamento = LeggiCampo(MySql,"IdStatoAffidamento")

    MySql = ""
    MySql = MySql & " Select * "
    MySql = MySql & " from AffidamentoRichiestaComp "
    MySql = MySql & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    IdFlussoProcesso   = LeggiCampo(MySql,"IdFlussoProcesso")

    'prima richiesta - si inviano i dati al back office 
    if ucase(IdStatoAffidamento)="COMPILA" then 
       'cambio stato e la porto in richiesta 
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiestaComp "
       qUpd = qUpd & " Set IdStatoAffidamento = 'RICH' "
       qUpd = qUpd & " Where IdAffidamentoRichiesta=" & IdAffidamentoRichiesta
       ConnMsde.execute qUpd
      
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiesta "
       qUpd = qUpd & " Set IdStatoAffidamento = 'RICH' "
       qUpd = qUpd & " Where IdAffidamentoRichiesta=" & IdAffidamentoRichiesta
       ConnMsde.execute qUpd
         
       infoProcesso = "La sua richiesta di affidamento e' stata inviata;"
       XX=createEvento("AFFI","RICH",Session("LoginIdAccount"),"Nuova richiesta di affidamento","AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,0)
    end if 
    'integrazione documenti : riporto in CHECK_DOCU
    if ucase(IdFlussoProcesso)="INT_DOCU" then 
       'cambio stato e la porto in richiesta 
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiestaComp "
       qUpd = qUpd & " Set IdFlussoProcesso = 'CHECK_DOCU' "
       qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
       ConnMsde.execute qUpd
      
	   '?????? 
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiesta "
       qUpd = qUpd & " Set IdFlussoProcessoBackO = 'CHECK_DOCU' "
	   qUpd = qUpd & "    ,IdFlussoProcessoCliente = 'CHECK_DOCU' "
       qUpd = qUpd & " Where IdAffidamentoRichiesta=" & IdAffidamentoRichiesta
       'ConnMsde.execute qUpd
         
       infoProcesso = "La documentazione di integrazione e' stata inviata;"
       XX=createEvento("AFFI","INTE",Session("LoginIdAccount"),"Documentazione integrata","AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,0)
    end if 
	
 end if 
 
 'carico un nuovo documento 
 if Oper="CALL_ADDD"  and CheckTimePageLoad() then 
    IdDocumento = Request("gruppo1")
    IdAccountDocumento = 0

    qIns = ""
    qIns = qIns & "Insert into AffidamentoRichiestaCompDoc ("
    qIns = qIns & " IdAffidamentoRichiestaComp"
    qIns = qIns & ",IdDocumento"
    qIns = qIns & ",flagObbligatorio"
    qIns = qIns & ",FlagDataScadenza"
    qIns = qIns & ",IdAccountDocumento) "
    qIns = qIns & " values ("
    qIns = qIns & " " & IdAffidamentoRichiestaComp
    qIns = qIns & "," & IdDocumento
    qIns = qIns & ",1"
    qIns = qIns & ",1"
    qIns = qIns & "," & IdAccountDocumento & " ) "
    ConnMsde.execute qIns 
 end if 
 'gestito il documento
 if Oper="CALL_GEST" then 
    IdDocumento  = "0" & request("ItemToRemove")
    if Cdbl(IdDocumento)>0 then 
       xx=RemoveSwap()
       Session("swap_IdTabella")          = "CLIENTE_DOC"
       Session("swap_IdTabellaKeyInt")    = IdAccountCliente
       Session("swap_IdAccount")          = IdAccountCliente
       Session("swap_IdRichiesta")        = IdAffidamentoRichiestaComp
       Session("swap_IdDocumento")        = IdDocumento
       Session("swap_TipoRife")           = request("ItemToModify")
       Session("swap_IdRife")             = request("IdParm")
       Session("swap_PaginaReturn")       = "configurazioni/Clienti/affidamento/" & NomePagina
       response.redirect RitornaA("configurazioni/Clienti/DocumentoClienteUploadGestione.asp")
       response.end 
    end if 
 end if 
 'validato un documento 
 if Oper="CALL_VALI" then 
    idRichiesta  = "0" & request("ItemToRemove")
    if Cdbl(idRichiesta)>0 then 
       qUpd = ""
       qUpd = qUpd & " Update AccountDocumento set"
       qUpd = qUpd & " IdTipoValidazione='VALIDO'"
       qUpd = qUpd & ",DataValidazione=getdate()"
       qUpd = qUpd & ",NoteValidazione=''"
       qUpd = qUpd & " Where IdAccountDocumento=" & IdRichiesta 
       'response.write qUpd
       ConnMsde.execute qUpd
    end if 
 end if 
 'cliccato su tasto salva 
 if Oper="CALL_REGI" then 
    idSt = trim(request("IdParm"))
	'aggiorno le descrizioni 
	if idSt<>"" then 
       DeSt = Request("NewDescStato0")
       ClSt = Request("NewDescStatoClie0")
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
       qUpd = qUpd & " NoteAffidamento = '"    & apici(deSt) & "'"
       qUpd = qUpd & ",NoteAffidamentoCliente = '"    & apici(clSt) & "'"
       qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
       'response.write qUpd 
       ConnMsde.execute qUpd 
       'richiesta di integrazione 
       if IdSt="DOCU" then 
          qUpd = ""
          qUpd = qUpd & " Update AffidamentoRichiesta set  "
          qUpd = qUpd & " IdFlussoProcessoBackO='INT_DOCU'"
		  qUpd = qUpd & ",IdFlussoProcessoCliente='INT_DOCU'"
          qUpd = qUpd & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
          ConnMsde.execute qUpd 
   
          qUpd = ""
          qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
          qUpd = qUpd & " IdFlussoProcesso='INT_DOCU'"
          qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
          ConnMsde.execute qUpd 

          descInfo="Integrazione documentazione per affidamamento"
          IdProdotto = 0
          XX=createEvento("AFFI",IdSt,Session("LoginIdAccount"),descInfo,"AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,IdProdotto)
  
       end if 
       if IdSt="COMP" then 
          MySql = ""
          MySql = MySql & " Select * "
          MySql = MySql & " from AffidamentoRichiestaComp "
          MySql = MySql & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
          IdFlussoProcesso   = LeggiCampo(MySql,"IdFlussoProcesso")
          if IdFlussoProcesso="CHECK_DOCU" then
             qUpd = ""
             qUpd = qUpd & " Update AffidamentoRichiesta set  "
             qUpd = qUpd & " IdFlussoProcessoBackO='SEL_COMPAGNIA'"
             qUpd = qUpd & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
             ConnMsde.execute qUpd 
   
             qUpd = ""
             qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
             qUpd = qUpd & " IdFlussoProcesso='SEL_COMPAGNIA'"
             qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
             ConnMsde.execute qUpd 
         end if  
         xx=RemoveSwap()
         Session("swap_IdAffidamentoRichiesta")     = IdAffidamentoRichiesta
         Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
         Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
         response.redirect RitornaA("configurazioni/Clienti/affidamento/GestioneCompagnia.asp")
         response.end 

       end if 
	   if IdSt="ANNU" then 
          qUpd = ""
          qUpd = qUpd & "update AffidamentoRichiestaComp "
          qUpd = qUpd & " Set IdStatoAffidamento = 'ANNU',DataChiusura= " & DtoS() 
		  qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
          ConnMsde.execute qUpd 		  
       end if    
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

   'leggo il dettaglio della richiesta 
   MySql = ""
   MySql = MySql & " Select A.*,StFl.DescrizioneUtente as DescStatoAffidamento "
   MySql = MySql & " from AffidamentoRichiestaComp A," & getStatoFlusso() 
   MySql = MySql & " Where A.IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
   MySql = MySql & " and   A.IdStatoAffidamento = StFl.IdStatoSorgente"
   MySql = MySql & " and   StFl.IdFlussoProcesso  in ('*',A.IdFlussoProcesso)"
   'response.write MySql
   err.clear 

   
   Rs.CursorLocation = 3 
   Rs.Open MySql, ConnMsde
   StatoComp                = funDoc_StatoDocum(IdAffidamentoRichiestaComp)     
   IdAccountCliente         = Rs("IdAccountCliente")
   DescCliente              = LeggiCampo("Select * from Account Where idAccount=" & IdAccountCliente,"Nominativo" )
   idCompagniaComp          = Rs("IdCompagnia")
   if cdbl(idCompagniaComp)>0 then 
      DescCompagniaComp        = LeggiCampo("select * from Compagnia Where IdCompagnia=" & idCompagniaComp,"DescCompagnia")
   else 
      DescCompagniaComp        = ""
   end if 
   DataRichiestaComp        = Rs("DataRichiesta")
   IdStatoAffidamentoComp   = Rs("IdStatoAffidamento")
   DescStatoAffidamentoComp = Rs("DescStatoAffidamento")
   IdFlussoProcesso         = Rs("IdFlussoProcesso")

   DataChiusuraComp         = Rs("DataChiusura")
   NoteAffidamentoComp      = Rs("NoteAffidamento")
   NoteAffidamentoClie      = Rs("NoteAffidamentoCliente")
   ValoreAffidamento        = Rs("ValoreAffidamento")
   ImptSingolaPolizza       = Rs("ImptSingolaPolizza")
   ValidoDalComp            = Rs("ValidoDal")
   ValidoAlComp             = Rs("ValidoAl")
   IdFornitore              = Rs("IdFornitore")
   if cdbl(IdFornitore)>0 then 
      DescFornitore = LeggiCampo("Select * from Fornitore Where IdFornitore=" & IdFornitore,"DescFornitore" )
   else
      DescFornitore = ""
   end if   
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
                <% end if %>
                <div class="col-11">
                <h3>
				<%if IdStatoAffidamentoComp="ANNU" then %>
				    Richiesta Di Affidamento 
                <%elseif Cdbl(idCompagniaComp)>0 then%>
                    Documentazione Richiesta Di Affidamento per compagnia
                <%else%>
                    Documentazione Richiesta Di Affidamento 
                <%end if%>
                </h3>
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
                     <input type="text" readonly class="form-control" value="<%=Stod(DataRichiestaComp)%>" >
                  </div>        
               </div>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Stato Richiesta")
                     %>
                     <input type="text" readonly class="form-control" value="<%=DescStatoAffidamentoComp%>" >
                  </div>        
               </div>
               
            </div>
            
            <div class="row">
               <div class="col-2">
                  <div class="form-group">
                     <%xx=ShowLabel("Elaborata il")%>
                     <input type="text" readonly class="form-control" value="<%=Stod(DataChiusuraComp)%>" >
                  </div>                       
               </div> 

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
            <%if IdStatoAffidamentoComp = "AFFI" then %>
            <div class="row">
               <div class="col-2">
                  <div class="form-group">
                     <%xx=ShowLabel("Importo Affidato &euro;")%>
                     <input type="text" readonly class="form-control" value="<%=InsertPoint(ValoreAffidamento,2)%>" >
                  </div>                       
               </div> 
               <div class="col-2">
                  <div class="form-group">
                     <%xx=ShowLabel("x Singola Polizza &euro;")%>
                     <input type="text" readonly class="form-control" value="<%=InsertPoint(ImptSingolaPolizza,2)%>" >
                  </div>                       
               </div> 


               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Valido Dal")
                     %>
                  <input type="textArea" readonly class="form-control" value="<%=Stod(ValidoDalComp)%>" >      
                  </div>               
               </div> 
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Valido Al")
                     %>
                  <input type="textArea" readonly class="form-control" value="<%=Stod(ValidoAlComp)%>" >      
                  </div>               
               </div> 
               
            </div>             
            <%end if %>

<%

FlagInviaRichiesta=false 
TextInviaRichiesta=""
'prima richiesta : dati completi - si puo' procedere con la richiesta 
if (isCliente() or IsCollaboratore()) then 
   if cdbl(FlagStatoFinale)=0 then 
      if ucase(IdStatoAffidamento)="COMPILA" and len(StatoComp)>0 and INSTR("COMPL_OK",StatoComp)>0 then 
         FlagInviaRichiesta=true  
         TextInviaRichiesta="La Richiesta &egrave; completa: &egrave; possibile procedere per l'affidamento"
      end if 
      if ucase(IdFlussoProcesso)="INT_DOCU" and len(StatoComp)>0 and INSTR("COMPL_OK",StatoComp)>0 then 
         FlagInviaRichiesta=true  
         TextInviaRichiesta="La Richiesta &egrave; completa: &egrave; possibile reinviare la documentazione"
      end if 	  
   end if 
end if 
%>

<%if infoProcesso <>"" then 
     MsgInfo = infoProcesso
%>

    <!--#include virtual="/gscVirtual/include/showInfoDivRow.asp"-->
<%end if %>


<%
IdCompagnia=idCompagniaComp
ShowColors = true 
ShowAction = true
OpDocAmm=""
ContaDocAssenti=0
ContaDocKo=0
FiltraDocumenti=""
ElencoIdDocumenti="0"
if IsCliente() or IsCollaboratore() then 
   OpDocAmm=""
   'metto in gestione solo i casi ammessi 
   if ucase(IdStatoAffidamentoComp)="COMPILA" then 
      OpDocAmm="G"
   end if 
   if ucase(IdFlussoProcesso)="INT_DOCU" then 
      OpDocAmm="G"
   end if 
   
end if 

if IsBackOffice() and FlagStatoFinale=0 then 
  'validazione
  OpDocAmm="QOG"
  if IdFlussoProcesso<>"CHECK_DOCU" then 
     OpDocAmm="O"
  end if    
end if 

%>
<!--#include virtual="/gscVirtual/configurazioni/clienti/Affidamento/DocumentiLista.asp"-->

<%
if (IsCliente() or IsCollaboratore()) and NoteAffidamentoClie<>"" then 
%>
   <div class="row">
        <div class="col-1 text-right"><%xx=ShowLabel("Note")%></div>
        <div class="col-9">
             <div class="form-group ">
                  <input type="text" readonly id="NewDescStatoClie0" name="NewDescStatoClie0" class="form-control" value="<%=NoteAffidamentoClie%>" >
             </div>        
        </div>
        <div class="col-2">
        </div>                   
   </div>            

<%
end if 
%>

<%
  if FlagInviaRichiesta=true then
  %>
   <div class="row">
        <div class="col-1"></div>
        <div class="col-2">
             <button type="button" onclick="localSend()"   class="btn btn-success">Invia Richiesta</button>
        </div>
   </div>

<%end if%>


<%
'gestisco i dati se è possibile modificare 
if IsBackOffice() then 
   if IdStatoAffidamentoComp="LAVO" and IdFlussoProcesso="CHECK_DOCU" then 
      cantModify = "" 
   else
      cantModify = " readonly " 
   end if 
%>
   <div class="row">
        <div class="col-1"><%xx=ShowLabel("Note interne")%></div>            
        <div class="col-9">
             <div class="form-group ">
                  <input type="text" <%=cantModify%> id="NewDescStato0" name="NewDescStato0" class="form-control" value="<%=NoteAffidamentoComp%>" >
             </div>        
        </div>
        <div class="col-2">
             <%idDaPulire ="NewDescStato0"
             idDaSalvare=""
		     if cantModify="" then%>
                <!--#include virtual="/gscVirtual/include/pulisciSalva.asp"-->
			 <%end if %>
       </div>               
   </div>
   <div class="row">
        <div class="col-1"><%xx=ShowLabel("Note per il cliente")%></div>
        <div class="col-9">
             <div class="form-group ">
                  <input type="text" <%=cantModify%> id="NewDescStatoClie0" name="NewDescStatoClie0" class="form-control" value="<%=NoteAffidamentoClie%>" >
             </div>        
        </div>
        <div class="col-2">
             <%idDaPulire ="NewDescStatoClie0"
               idDaSalvare=""
	           if cantModify="" then%>
                  <!--#include virtual="/gscVirtual/include/pulisciSalva.asp"-->
			   <%end if %>
        </div>                   
   </div>            
   <div class="row">
        <div class="col-1"></div>            
		
		<%if IdStatoAffidamentoComp="LAVO" and IdFlussoProcesso="CHECK_DOCU" then%>
             <div class="col-2">
                  <button type="button" onclick="registraStato('SAVE')"   class="btn btn-success">Salva</button>
             </div>
             <div class="col-2">
                  <button type="button" onclick="registraStato('ANNU')"   class="btn btn-danger">Annulla</button>
             </div>
			 <%if cdbl(ContaDocKo)>0 then%>
             <div class="col-2">
                  <button type="button" onclick="registraStato('DOCU')"   class="btn btn-info">Integr.Documenti</button>
             </div>
			 <%end if %>
        <%end if %>	
		<%if ContaDocAssenti=0 and ContaDocKo=0 and instr("ANNU",IdStatoAffidamento)=0 then%>
             <div class="col-2">
                  <button type="button" onclick="registraStato('COMP')"   class="btn btn-success">Assegna Compagnie</button>
             </div>
        <%end if%>     
     </div>            
<% end if 'gestione back office
  %>
            <input type="hidden" name="localVirtualPath" id="localVirtualPath" value = "<%=VirtualPath%>">
            <!--#include virtual="/gscVirtual/utility/SelezioneDaCassetto.asp"-->

            
            <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
            <!--#include virtual="/gscVirtual/include/paginazione.asp"-->

            
            
<div class="modal fade" id="confirmModal"  aria-hidden="true" role="dialog">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">

        <h2>Documento Da Caricare </h2> 
        <button type="button" class="close" data-dismiss="modal">
          <span aria-hidden="true">×</span><span class="sr-only">Chiudi</span>
        </button>
      </div>

      <div class="modal-body"> 
        <div>
          <%
             Conta=0
             Rs.CursorLocation = 3 
             Rs.Open "select * from Documento where IdDocumentoInterno='' and idDocumento not in (" & ElencoIdDocumenti & ") Order By Descdocumento", ConnMsde
             Do while not rs.eof
                Conta=Conta+1
             %>
              <div class="form-check">
                 <input name="gruppo1" type="radio" id="radio1"  value="<%=Rs("IdDocumento")%>">
                 <label for="radio1"><%=Rs("DescDocumento")%></label>
             </div> 
             <%
                Rs.movenext
             loop 
             Rs.close 
          
          %>
        
        </div>          
      </div> 

      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
        <%if conta>0 then %>
        <button type="button" class="btn btn-primary" onclick="LocalAddRowNew();";>Seleziona</button>
        <%end if %>
      </div>
    </div>
  </div>
</div>
            
<div class="modal fade" id="confirmModalCoob"  aria-hidden="true" role="dialog">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">

        <h2>Coobbligato Da Aggiungere </h2> 
        <button type="button" class="close" data-dismiss="modal">
          <span aria-hidden="true">×</span><span class="sr-only">Chiudi</span>
        </button>
      </div>

      <div class="modal-body"> 
        <div>
          <%
             Conta=0
             Rs.CursorLocation = 3 
             qNot = ""
             qNot = qNot & " select IdAccountCoobbligato "
             qNot = qNot & " from AffidamentoRichiestaCompCoob "
             qNot = qNot & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
             q = ""
             q = q & " select * from AccountCoobbligato "
             q = q & " where IdAccount=" & IdAccountCliente 
             q = q & " and ValidoDal<=" & Dtos()
             q = q & " and ValidoAl>=" & Dtos()
             q = q & " and IdAccountCoobbligato not in (" & qNot & ") Order By RagSoc"
             'response.write q
             Rs.Open q, ConnMsde
             if err.number = 0 then 
                Do while not rs.eof
                   Conta=Conta+1
             %>
              <div class="form-check">
                 <input name="gruppo1Coob" type="radio" id="radio1"  value="<%=Rs("IdAccountCoobbligato")%>">
                 <label for="radio1"><%=Rs("RagSoc")%></label>
             </div> 
             <%
                   Rs.movenext
                loop 
                Rs.close 
             else 
                response.write err.description 
             end if 
          
          %>
        
        </div>          
      </div> 

      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
        <%if conta>0 then %>
        <button type="button" class="btn btn-primary" onclick="LocalAddRowNewCoob();";>Aggiungi</button>
        <%end if %>
      </div>
    </div>
  </div>
</div>
            
<div class="modal fade" id="confirmModalCert"  aria-hidden="true" role="dialog">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">

        <h2>Certificazione Da Aggiungere </h2> 
        <button type="button" class="close" data-dismiss="modal">
          <span aria-hidden="true">×</span><span class="sr-only">Chiudi</span>
        </button>
      </div>

      <div class="modal-body"> 
        <div>
          <%
             NotIn = "select IdCertificazione from AffidamentoRichiestaCompCert Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
             qSele = ""
             qSele = qSele & " Select A.* "
             qSele = qSele & " From Certificazione A, AccountCertificazione B"
             qSele = qSele & " Where A.IdCertificazione = B.IdCertificazione"
             qSele = qSele & " and B.IdAccount=" & IdAccountCliente
             qSele = qSele & " and A.IdCertificazione not in (" & NotIn & ")"
             qSele = qSele & " order by A.DescBreveCertificazione"
             'response.write qSele 
             Conta=0
             Rs.CursorLocation = 3 
             Rs.Open qSele
             if err.number=0 then 
                Do while not rs.eof
                   Conta=Conta+1
             %>
              <div class="form-check">
                 <input name="gruppo1Cert" type="radio" id="radio1"  value="<%=Rs("IdCertificazione")%>">
                 <label for="radio1"><%=Rs("DescBreveCertificazione")%></label>
             </div> 
             <%
                   Rs.movenext
                loop 
                Rs.close 
             end if 
          
          %>
        
        </div>          
      </div> 

      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
        <%if conta>0 then %>
        <button type="button" class="btn btn-primary" onclick="LocalAddRowNewCert();";>Aggiungi</button>
        <%end if %>
      </div>
    </div>
  </div>
</div>
            
            
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
