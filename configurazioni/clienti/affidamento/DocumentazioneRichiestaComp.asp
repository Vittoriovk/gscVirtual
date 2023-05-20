<%
  NomePagina="DocumentazioneRichiestaComp.asp"
  titolo="Richiesta Di Affidamento Cliente"
  default_check_profile="Coll,Clie,BackO"
  act_call_tfir  = CryptAction("CALL_TFIR") 
  act_call_inte  = CryptAction("CALL_INTE") 
  act_call_intd  = CryptAction("CALL_DOCU") 
  act_call_vali  = CryptAction("CALL_VALI") 
  act_call_icli  = CryptAction("CALL_INTC") 
  act_call_vint  = CryptAction("CALL_INTV") 
  act_call_save  = CryptAction("CALL_SAVE")
  act_call_newc  = CryptAction("CALL_NEWC")
  act_call_incc  = CryptAction("CALL_INCC")
  act_call_decc  = CryptAction("CALL_DECC")
  act_call_send  = CryptAction("CALL_SEND")
  
  act_call_addd  = CryptAction("CALL_ADDD") 
  act_call_gest  = CryptAction("CALL_GEST") 
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
function salva()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_save%>");
    document.Fdati.submit();
}
function inviaTrasfFirmato()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_tfir%>");
    document.Fdati.submit();
}
function inviaIntegrazione()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_inte%>");
    document.Fdati.submit();
}
function inviaIntegraDoc()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_intd%>");
    document.Fdati.submit();
}

function validaDocumenti()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_vali%>");
    document.Fdati.submit();
}
function inviaClieIntegra()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_icli%>");
    document.Fdati.submit();
}
function validaIntDoc()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_vint%>");
    document.Fdati.submit();
}
function LocalAddNewCoob()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_newc%>");
    document.Fdati.submit();
}

function LocalIncCoob()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_incc%>");
    document.Fdati.submit();
}
function LocalDecCoob()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_decc%>");
    document.Fdati.submit();
}

function inviaForn()
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
 
 'richiama la gestione dei coobbligati
 if Oper="CALL_NEWC" then 
   xx=RemoveSwap()
   Session("swap_IdCliente")    = 0
   Session("swap_IdAccount")    = IdAccountCliente
   Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
   response.redirect RitornaA("configurazioni/Clienti/ClienteCoobbligati.asp")
   response.end  
 end if  
 'invio del documento firmato 
 if Oper="CALL_TFIR" then 
    MySql = ""
    MySql = MySql & " Select * "
    MySql = MySql & " from AffidamentoRichiestaComp "
    MySql = MySql & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    IdFlussoProcesso   = LeggiCampo(MySql,"IdFlussoProcesso")
    if IdFlussoProcesso="CAC_FORNITORE" then 
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiestaComp "
       qUpd = qUpd & " Set IdFlussoProcesso = 'CAV_FORNITORE' "
       qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
       ConnMsde.execute qUpd

       infoProcesso = "La sua richiesta di trasferimento fornitore e' stata inviata;"
       XX=createEvento("AFFI","RICH",Session("LoginIdAccount"),"Richiesta di trasferimento","AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,0)
	end if 
 end if 
 'il cliente manda i documenti integrati 
 if Oper="CALL_INTC" then 
    MySql = ""
    MySql = MySql & " Select * "
    MySql = MySql & " from AffidamentoRichiestaComp "
    MySql = MySql & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    IdFlussoProcesso   = LeggiCampo(MySql,"IdFlussoProcesso")
    if IdFlussoProcesso="INC_FORNITORE" then 
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiestaComp "
       qUpd = qUpd & " Set IdFlussoProcesso = 'INT_FORNITORE' "
       qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
       ConnMsde.execute qUpd

       infoProcesso = "La sua richiesta di integrazione e' stata inviata;"
       XX=createEvento("AFFI","RICH",Session("LoginIdAccount"),"Integrazione Documenti","AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,0)
	end if 
 end if 
 
 
 'integrazione documentazione 
 if Oper="CALL_INTE" then 
    NoteClie = Request("NewDescStatoClie0")
    NoteInte = Request("NewDescStato0")
    MySql = ""
    MySql = MySql & " Select * "
    MySql = MySql & " from AffidamentoRichiestaComp "
    MySql = MySql & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    IdFlussoProcesso   = LeggiCampo(MySql,"IdFlussoProcesso")
	'sono nel processo di firma : rimetto in modifica per permettere al cliente di aggiornare 
	'response.write IdFlussoProcesso
    if IdFlussoProcesso="CAV_FORNITORE" then 
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiestaComp set "
       qUpd = qUpd & " IdFlussoProcesso = 'CAC_FORNITORE' "
	   qUpd = qUpd & ",NoteAffidamento='" & apici(NoteInte) & "'" 
	   qUpd = qUpd & ",NoteAffidamentoCliente='" & apici(NoteClie) & "'" 
       qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
       ConnMsde.execute qUpd

       infoProcesso = "Richiesta di integrazione inviata ;"
       XX=createEvento("AFFI","RICH",Session("LoginIdAccount"),"Richiesta di trasferimento : documentazione da integrare","AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,0)
	end if 
 end if
 
 if Oper="CALL_DOCU" then 
    NoteClie = Request("NewDescStatoClie0")
    NoteInte = Request("NewDescStato0")
    MySql = ""
    MySql = MySql & " Select * "
    MySql = MySql & " from AffidamentoRichiestaComp "
    MySql = MySql & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    IdFlussoProcesso   = LeggiCampo(MySql,"IdFlussoProcesso")
	'sono nel processo di firma : rimetto in modifica per permettere al cliente di aggiornare 
	'response.write IdFlussoProcesso
    if IdFlussoProcesso="INT_FORNITORE" then 
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiestaComp set "
       qUpd = qUpd & " IdFlussoProcesso = 'INC_FORNITORE' "
	   qUpd = qUpd & ",NoteAffidamento='" & apici(NoteInte) & "'" 
	   qUpd = qUpd & ",NoteAffidamentoCliente='" & apici(NoteClie) & "'" 
       qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
       ConnMsde.execute qUpd

       infoProcesso = "Richiesta di integrazione inviata ;"
       XX=createEvento("AFFI","RICH",Session("LoginIdAccount"),"Richiesta di affidamento : documentazione da integrare","AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,0)
	end if 
 end if

 if Oper="CALL_SAVE" then 
    NoteClie = Request("NewDescStatoClie0")
    NoteInte = Request("NewDescStato0")
    NumeCoob = cdbl("0" & Request("NumCoobbligatiRichiesti0"))
    qUpd = ""
    qUpd = qUpd & " Update AffidamentoRichiestaComp set "
    qUpd = qUpd & " NoteAffidamento='" & apici(NoteInte) & "'" 
    qUpd = qUpd & ",NoteAffidamentoCliente='" & apici(NoteClie) & "'" 
    qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
    ConnMsde.execute qUpd
 end if   
  if Oper="CALL_INCC" then 
    qUpd = ""
    qUpd = qUpd & " Update AffidamentoRichiestaComp set "
	qUpd = qUpd & " NumCoobbligatiRichiesti=NumCoobbligatiRichiesti+1 "
    qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
    ConnMsde.execute qUpd
 end if  
  if Oper="CALL_DECC" then 
    qUpd = ""
    qUpd = qUpd & " Update AffidamentoRichiestaComp set "
	qUpd = qUpd & " NumCoobbligatiRichiesti=NumCoobbligatiRichiesti-1 "
    qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
	qUpd = qUpd & " and NumCoobbligatiRichiesti>0"
    ConnMsde.execute qUpd
 end if  

 
 if Oper="CALL_VALI" then 
    NoteClie = Request("NewDescStatoClie0")
    NoteInte = Request("NewDescStato0")
    MySql = ""
    MySql = MySql & " Select * "
    MySql = MySql & " from AffidamentoRichiestaComp "
    MySql = MySql & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    IdFlussoProcesso   = LeggiCampo(MySql,"IdFlussoProcesso")
	'sono nel processo di firma : rimetto in modifica per permettere al cliente di aggiornare 
	'response.write IdFlussoProcesso
    if IdFlussoProcesso="CAV_FORNITORE" then 
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiestaComp set "
       qUpd = qUpd & " IdFlussoProcesso = 'CAM_FORNITORE' "
	   qUpd = qUpd & ",FlagTrasferimento = 99 "
	   qUpd = qUpd & ",NoteAffidamento='" & apici(NoteInte) & "'" 
	   qUpd = qUpd & ",NoteAffidamentoCliente='" & apici(NoteClie) & "'" 
       qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
       ConnMsde.execute qUpd
       xx=RemoveSwap()
       Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
       Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento//ListaRichiesta.asp"
       response.redirect     RitornaA("configurazioni/Clienti/affidamento/GestioneFornitore.asp")
       response.end 
	end if 
 end if 
 'integrazione validata 
 if Oper="CALL_INTV" then 
    NoteClie = Request("NewDescStatoClie0")
    NoteInte = Request("NewDescStato0")
    MySql = ""
    MySql = MySql & " Select * "
    MySql = MySql & " from AffidamentoRichiestaComp "
    MySql = MySql & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    IdFlussoProcesso   = LeggiCampo(MySql,"IdFlussoProcesso")
	'sono nel processo di firma : rimetto in modifica per permettere al cliente di aggiornare 
	'response.write IdFlussoProcesso
    if IdFlussoProcesso="INT_FORNITORE" then 
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiestaComp set "
       qUpd = qUpd & " IdFlussoProcesso = 'GES_FORNITORE' "
	   qUpd = qUpd & ",NoteAffidamento='" & apici(NoteInte) & "'" 
	   qUpd = qUpd & ",NoteAffidamentoCliente='" & apici(NoteClie) & "'" 
       qUpd = qUpd & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
       ConnMsde.execute qUpd
       infoProcesso = "Documentazione validata ;"
       XX=createEvento("AFFI","RICH",Session("LoginIdAccount"),"Integrazione Documenti : documentazione validata","AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,0)	   
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
   if Oper="CALL_ADD_COOB" and CheckTimePageLoad() then 
      
      IdCoob = "0" & Request("gruppo1Coob")
      if Cdbl(IdCoob)>0 then 
         qSelCoob = ""
         qSelCoob = qSelCoob & " Select " & IdAffidamentoRichiestaComp & " as IdAffidamentoRichiestaComp"
         qSelCoob = qSelCoob & " ,PI,CF,RagSoc,Indirizzo,Cap,Comune,Provincia, 0 as FlagValidato,'' as Note,IdAccountCoobbligato "
         qSelCoob = qSelCoob & " FROM AccountCoobbligato where IdAccountCoobbligato=" & IdCoob

         qIns = ""
         qIns = qins & " INSERT INTO AffidamentoRichiestaCompCoob ( "
         qIns = qins & " IdAffidamentoRichiestaComp,PI,CF,RagSoc,Indirizzo,Cap,Comune,Provincia,FlagValidato,Note,IdAccountCoobbligato"
         qIns = qins & " ) " & qSelCoob 
         ConnMsde.execute qIns 

      end if 
      
   end if 
   if Oper="CALL_REM_COOB" and CheckTimePageLoad() then 
      IdCoob = "0" & Request("ItemToRemove")
       if Cdbl(IdCoob)>0 then 
          IdCoobbligato = LeggiCampo("Select * from AffidamentoRichiestaCompCoob Where IdAffidamentoRichiestaCompCoob=" & IdCoob,"IdAccountCoobbligato")
          ConnMsde.execute "Delete from AffidamentoRichiestaCompCoob Where IdAffidamentoRichiestaCompCoob=" & IdCoob
          qDel = ""
          qDel = qDel & "Delete from AffidamentoRichiestaCompDoc "
          qDel = qDel & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
          qDel = qDel & " and TipoRife = 'COOB'"
          qDel = qDel & " and IdRife = " & IdCoobbligato
          'response.write qDel 
          ConnMsde.execute qDel  		  
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
   IdFlussoProcesso         = rs("IdFlussoProcesso")
   NumCoobbligatiRichiesti  = rs("NumCoobbligatiRichiesti")
   PathDocumentoZip         = rs("PathDocumentoZip")     
   if cdbl(IdFornitore)>0 then 
      DescFornitore = LeggiCampo("Select * from Fornitore Where IdFornitore=" & IdFornitore,"DescFornitore" )
   else
      DescFornitore = ""
   end if   
   Rs.close
   
   Set DizProcesso = CreateObject("Scripting.Dictionary")
   esito=getInfoProcessoAffi(DizProcesso,IdStatoAffidamentoComp,IdFlussoProcesso)
   if esito = true then 
      descProcesso = " - " & GetDiz(DizProcesso,"DescrizioneUtente")
   else
      descProcesso = "" 
   end if 
   
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
               <div class="col-4">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Stato Richiesta")
                     %>
                     <input type="text" readonly class="form-control" value="<%=DescStatoAffidamentoComp %>" >
                  </div>        
               </div>
            </div>
            <div class="row">			   
			   <%if Cdbl(idCompagniaComp)>0 then %>
               <div class="col-4">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Compagnia")
                     %>
                     <input type="text" readonly class="form-control" value="<%=DescCompagniaComp%>" >
                  </div>        
               </div>
			   <%end if %>
			   <%if IsBackOffice() and Cdbl(IdFornitore)>0 then %>
               <div class="col-4">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Fornitore")
                     %>
                     <input type="text" readonly class="form-control" value="<%=DescFornitore%>" >
                  </div>        
               </div>
			   <%end if %>               
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
			
			<%'mostro il documento da scaricare e da firmare 
			  if IdFlussoProcesso="CAC_FORNITORE" and IsBackOffice()=false then
			     
                 IdDocTraD = LeggiCampo("Select * from Documento Where IdDocumentoInterno='MODULO_TRASF_D'","IdDocumento")
                 IdDocTraD = TestNumeroPos(IdDocTraD)
                 qSel = ""
                 qSel = qSel & " select * from Upload "
                 qSel = qSel & " Where IdTabella='AFFIDAMENTO'" 
                 qSel = qSel & " and IdTabellaKeyInt = " & IdAffidamentoRichiestaComp 
                 qSel = qSel & " and IdTipoDocumento = " & IdDocTraD
                 Linkdocumento = LeggiCampo(qSel,"PathDocumento")
				 if Linkdocumento<>"" then %>
		            <div class="row">
					   <div class="col-1"></div>
					   <div class="col-1">
						   <!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->
					   </div>					   					   
		               <div class="col-9">
			               <h4>Modulo da scaricare , firmare ed aggiungere al rigo relativo al modulo di trasferimento</h4>
					   </div>

					</div>
			<%   end if 
			  end if %>
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

  if FlagInviaRichiesta=true then
     RiferimentoA="col-1;#;;2;ok;Procedi;;localSend();N"
  %>
   <div class="row">
       <div class="col-2"></div>
       <div class="col-5 form-group font-weight-bold bg-info text-white"><%=TextInviaRichiesta%></div>
       <div class="col-1"></div>
       <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
       <div class="col-2"></div>
   </div>

<%end if%>

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
FiltraDocumenti  =""
ElencoIdDocumenti="0"
if IsCliente() or IsCollaboratore() then 
   OpDocAmm=""
   'firma trasferimento 
   if ucase(IdFlussoProcesso)="CAC_FORNITORE" and FlagStatoFinale=0 then 
      OpDocAmm="G"
      IdDocumento = LeggiCampo("Select * from Documento Where IdDocumentoInterno='MODULO_TRASF_U'","IdDocumento")
      IdDocumento = TestNumeroPos(IdDocumento)		  
	  'mostro solo il documento da caricare per il trasferimento 
	  FiltraDocumenti = IdDocumento
   end if 
   if ucase(IdFlussoProcesso)="CAV_FORNITORE" and FlagStatoFinale=0 then 
      IdDocumento = LeggiCampo("Select * from Documento Where IdDocumentoInterno='MODULO_TRASF_U'","IdDocumento")
      IdDocumento = TestNumeroPos(IdDocumento)		  
	  'mostro solo il documento da caricare per il trasferimento 
	  FiltraDocumenti = IdDocumento
   end if    
   'integrazione documenti
   if ucase(IdFlussoProcesso)="INC_FORNITORE" and FlagStatoFinale=0 then 
      OpDocAmm="G"
   end if    
end if 

if IsBackOffice() and FlagStatoFinale=0 then 
  OpDocAmm="O"
  'cambio fornitore da back office   
  if instr("CAV_FORNITORE INT_FORNITORE",IdFlussoProcesso)>0 then 
     OpDocAmm="QOG"   
  end if    
end if 

%>
<!--#include virtual="/gscVirtual/configurazioni/clienti/Affidamento/DocumentiLista.asp"-->

<%
'devi permettere al back office di richiedere coobbligati 

SalvaRichiesta = false

'gestione dei coobbligati
OpDocAmm=""
ElencoRagSoc="''"
ShowElencoCoob=false 
'indica se tutti i sono presenti 
CoobPresentiTutti = true 
'indica se la documentazione è stata caricata 
CoobPresenteDocum = true 
'indica se tutti sono validi
CoobPresenteValid = true 
if (isBackOffice()) then 
   OpDocAmm="N"
end if 
'possibile aggiungere o rimuovere coobbligati
CoobCanAddRem = false 
if IsBackOffice() and instr("INT_FORNITORE",IdFlussoProcesso)>0 then
   CoobCanAddRem=true
end if 


if (isCliente() or IsCollaboratore()) and FlagStatoFinale=0 then 
   ShowElencoCoob=true 
   OpDocAmm="IG"
   if instr("INC_FORNITORE",IdFlussoProcesso)>0 then 
      OpDocAmm=OpDocAmm & "CI"
   end if 
end if 
if cdbl(NumCoobbligatiRichiesti)>0 or IsBackOffice() then 
   CoobPresentiTutti = false 
   CoobPresenteDocum = false 
   CoobPresenteValid = false
   ShowElencoCoob    = false  
   if CoobCanAddRem or cdbl(NumCoobbligatiRichiesti)>0 then 
      ShowElencoCoob = true
   end if 
%>
<!--#include virtual="/gscVirtual/configurazioni/clienti/Affidamento/CoobbligatiLista.asp"-->

<%
end if 

if (IsCliente() or IsCollaboratore()) then 
   if NoteAffidamentoClie<>"" then 
%>
   <div class="row">
        <div class="col-1 text-right"><%xx=ShowLabel("Note Per Cliente")%></div>
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
end if 
%>


<%
'gestisco i dati se è possibile modificare 
ShowAction    =false      
ClieInviafirma=false 
ClieInviaInteg=false 
BackInviafirma=false
BackInviaInteg=false 
BackIntegraDoc=false
BackIntegraVal=false 
if IsBackOffice()=false then 
   if ContaDocAssenti=0 and IdFlussoProcesso="CAC_FORNITORE" then
      ShowAction    =true
      ClieInviafirma=true 
   end if  
   'integrazione documenti controllo documenti ed eventuali coobbligati
   if IdFlussoProcesso="INC_FORNITORE" then
      'response.write ContaDocAssenti & CoobPresentiTutti & CoobPresenteDocum
      if ContaDocAssenti=0 and CoobPresentiTutti and CoobPresenteDocum then
         ShowAction    =true
         ClieInviaInteg=true 
	  end if 
   end if  

end if 

if IsBackOffice() then 
   cantModify = " readonly " 
   'puo' confermare o rimandare al cliente la richiesta di validazione del documento
   if instr("CAV_FORNITORE CAC_FORNITORE CAM_FORNITORE",IdFlussoProcesso)>0 then 
      cantModify = ""
	  ShowAction = true
      if ContaDocAssenti=0 and ContaDocKo=0 then
         BackInviafirma = true 
      end if 
	  if instr("CAV_FORNITORE",IdFlussoProcesso)>0 then 
	     BackInviaInteg=true  
	  end if 
   
   end if 
   
   if instr("INT_FORNITORE",IdFlussoProcesso)>0 then 
      cantModify = ""
	  ShowAction = true
      if ContaDocAssenti=0 and ContaDocKo=0 then
         BackIntegraVal = true 
      end if 
      if instr("INT_FORNITORE",IdFlussoProcesso)>0 then 
         BackIntegraDoc = true 
      end if  
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
         
  
        
	<% end if 'gestione back office
  %>
<%if ShowAction then %>
<div class="row">
   <div class="col-1"></div>            
   <% if SalvaRichiesta then %>
         <div class="col-2">
              <button type="button" onclick="salva()"   class="btn btn-info">Registra</button>
         </div>   
   <% end if %>   
   <!--  Gestione Trasferimento -->
   <% if ClieInviafirma then %>
         <div class="col-2">
              <button type="button" onclick="inviaTrasfFirmato()"   class="btn btn-success">Invia Richiesta</button>
         </div>   
   <% end if %>
   <% if BackInviafirma then %>
         <div class="col-2">
              <button type="button" onclick="validaDocumenti()"     class="btn btn-success">Valida Documenti</button>
         </div>   
   <% end if %>
   <% if BackInviaInteg then %>
         <div class="col-2">
              <button type="button" onclick="inviaIntegrazione()"   class="btn btn-info">Integr.Documenti</button>
         </div>   
   <% end if %>

   <!--  Gestione Integrazione Documenti -->

   <!--  il cliente invia i dati integrati -->
   <% if ClieInviaInteg then %>
         <div class="col-2">
              <button type="button" onclick="inviaClieIntegra()"   class="btn btn-success">Invia Integrazione</button>
         </div>   
   <% end if %>

   <!--  il back office valida l'integrazione -->
   <% if BackIntegraVal then %>
         <div class="col-2">
              <button type="button" onclick="validaIntDoc()"   class="btn btn-success">Valida Documenti</button>
         </div>   
         <div class="col-2">
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
 
   <% end if %>
   
   <!--  il back office richiede l'integrazione -->
   <% if BackIntegraDoc then %>
         <div class="col-2">
              <button type="button" onclick="inviaIntegraDoc()"   class="btn btn-info">Integr.Documenti</button>
         </div>   
   <% end if %>
   
   
</div>

<%end if %>

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
             'q = q & " and ValidoDal<=" & Dtos()
             'q = q & " and ValidoAl>=" & Dtos()
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
