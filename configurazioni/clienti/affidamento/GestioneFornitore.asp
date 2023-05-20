<%
  NomePagina="GestioneFornitore.asp"
  titolo="Menu - Dettaglio Richiesta Di Affidamento Cliente per Fornitore"
  default_check_profile="BackO"
  act_call_camb = CryptAction("CALL_CAMB")
  act_call_tras = CryptAction("CALL_TRAS")
  act_call_inte = CryptAction("CALL_INTE")
  act_call_coob = CryptAction("CALL_COOB")
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
function cambiaForn()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_camb%>");
    document.Fdati.submit();  
}
function trasferisciConf()
{
    xx=myConfirm("Affidamento","Conferma Trasferimento","<%=act_call_tras%>");
}
function trasferisci()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_tras%>");
    document.Fdati.submit();  

}
function integra()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_inte%>");
    document.Fdati.submit();  
}


function registraForn()
{
   xx=ImpostaValoreDi("Oper","GES_FORN");
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
	  xx=ImpostaValoreDi("NameRangeN","");
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
      xx=ImpostaValoreDi("NameLoaded","ValoreAffidamento,FLP;AffidamentoUsato,FLZ;ImptSingolaPolizza,FLP;ValidoDalComp,DTO;ValidoAlComp,DTO");
	  xx=ImpostaValoreDi("NameRangeN","AffidamentoUsato;ValoreAffidamento;0;9999999;ImptSingolaPolizza;ValoreAffidamento;0;9999999")
      xx=ImpostaValoreDi("DescLoaded","0");
      
      xx=ElaboraControlli();
      if (xx==false)
         return false;

      return true; 
}
function LocalGesRowCoob(id)
{
    xx=ImpostaValoreDi("IdParm",id);
    xx=ImpostaValoreDi("Oper","<%=act_call_coob%>");
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

if cdbl("0" & IdAffidamentoRichiesta)=0 then 
   qSel = "Select * from AffidamentoRichiestaComp Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp 
   IdAffidamentoRichiesta = LeggiCampo(qSel,"IdAffidamentoRichiesta")
end if  

 'registrazione dei dati :
 deSt = trim(Request("NewDescStato0"))
 clSt = trim(Request("NewDescStatoClie0"))
 
 Oper = DecryptAction(Oper)
 'response.write Oper
 if Oper = "CALL_COOB" then 
    Id=Request("IdParm")
    xx=RemoveSwap()
	Session("swap_TipoRife") = "COOB"
    Session("swap_IdRife")   = Id
    Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
    response.redirect     RitornaA("configurazioni/Clienti/ValidazioneBackODettaglio.asp")
    response.end  
 end if 
 
 if Oper = "CALL_CAMB" then 
    qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
    qUpd = qUpd & " IdFlussoProcesso = 'SEL_FORNITORE' "
    qUpd = qUpd & ",NoteAffidamento = '"    & apici(deSt) & "'"
    qUpd = qUpd & ",NoteAffidamentoCliente = '" & apici(clSt) & "'"
    qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    qUpd = qUpd & " and   IdFlussoProcesso = 'GES_FORNITORE' "
    ConnMsde.execute qUpd
    xx=RemoveSwap()
    Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
    Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
    response.redirect     RitornaA("configurazioni/Clienti/affidamento/SelezioneFornitore.asp")
    response.end 
 end if  
 if Oper = "CALL_TRAS" then 
    qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
    qUpd = qUpd & " IdFlussoProcesso = 'CAM_FORNITORE' "
	qUpd = qUpd & ",FlagTrasferimento = 1"
    qUpd = qUpd & ",NoteAffidamento = '"    & apici(deSt) & "'"
    qUpd = qUpd & ",NoteAffidamentoCliente = '" & apici(clSt) & "'"
    qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    qUpd = qUpd & " and   IdFlussoProcesso = 'GES_FORNITORE' "
	qUpd = qUpd & " and   FlagTrasferimento = 0"
    ConnMsde.execute qUpd
    xx=RemoveSwap()
	Session("swap_IdAffidamentoRichiesta")     = IdAffidamentoRichiesta
    Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
    Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
    response.redirect     RitornaA("configurazioni/Clienti/affidamento/TrasferisciFornitore.asp")
    response.end 
 end if 
 if Oper = "CALL_INTE" then 
    qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
    qUpd = qUpd & " IdFlussoProcesso = 'INT_FORNITORE' "
    qUpd = qUpd & ",NoteAffidamento = '"    & apici(deSt) & "'"
    qUpd = qUpd & ",NoteAffidamentoCliente = '" & apici(clSt) & "'"
    qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    ConnMsde.execute qUpd
    xx=RemoveSwap()
	Session("swap_IdAffidamentoRichiesta")     = IdAffidamentoRichiesta
    Session("swap_IdAffidamentoRichiestaComp") = IdAffidamentoRichiestaComp
    Session("swap_PaginaReturn") = "configurazioni/Clienti/affidamento/" & NomePagina
    response.redirect     RitornaA("configurazioni/Clienti/affidamento/DocumentazioneRichiestaComp.asp")
    response.end 
 end if 
 
 if Oper = "ANNULLA" then 
    qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
    qUpd = qUpd & " IdStatoAffidamento = 'ANNU' "
    qUpd = qUpd & ",DataChiusura=" & DtoS()
    qUpd = qUpd & ",NoteAffidamento = '"    & apici(deSt) & "'"
    qUpd = qUpd & ",NoteAffidamentoCliente = '" & apici(clSt) & "'"
    qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    qUpd = qUpd & " and   IdStatoAffidamento = 'LAVO'"
    ConnMsde.execute qUpd
    xx=AggiornaRichiestaAffidamento(IdAffidamentoRichiesta)
    response.redirect RitornaA(PaginaReturn)
 end if 
 ReloadImpt = true 
 if Oper = "CONFERMA" then 
    MySql = "Select * from AffidamentoRichiestaComp Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    IdFornitore=Cdbl("0" & LeggiCampo(MySql,"IdFornitore"))
    ValoreAffidamento  = Cdbl("0" & Request("ValoreAffidamento0"))
    AffidamentoUsato   = Cdbl("0" & Request("AffidamentoUsato0"))
    ImptSingolaPolizza = Cdbl("0" & Request("ImptSingolaPolizza0"))
    ValidoDalComp      = DataStringa(Request("ValidoDalComp0"))
    ValidoAlComp       = DataStringa(Request("ValidoAlComp0"))
    'salvo solo i valori 
    qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
    qUpd = qUpd & " NoteAffidamento = '"    & apici(deSt) & "'"
    qUpd = qUpd & ",NoteAffidamentoCliente = '" & apici(clSt) & "'"
    qUpd = qUpd & ",ValoreAffidamento = " & NumForDb(ValoreAffidamento)
    qUpd = qUpd & ",ImptSingolaPolizza = " & NumForDb(ImptSingolaPolizza)
    qUpd = qUpd & ",AffidamentoUsato = " & NumForDb(AffidamentoUsato)
    qUpd = qUpd & ",ValidoDal = " & NumForDb(ValidoDalComp)
    qUpd = qUpd & ",ValidoAl = " & NumForDb(ValidoAlComp)
    qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
	qUpd = qUpd & " and   IdStatoAffidamento = 'LAVO' "
    qUpd = qUpd & " and   IdFlussoProcesso in ('GES_FORNITORE','CAM_FORNITORE') "

    ConnMsde.execute qUpd

    MySql = ""
    MySql = MySql & " Select C.IdCompagnia,C.DescCompagnia from AffidamentoRichiestaComp A,Compagnia C"
    MySql = MySql & " Where A.IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    MySql = MySql & " and   A.IdCompagnia = C.IdCompagnia"
    DescCompagnia = LeggiCampo(MySql,"DescCompagnia")
    IdCompagnia   = LeggiCampo(MySql,"IdCompagnia")

    MySql = ""
    MySql = MySql & " select a.IdAccountCliente "
    MySql = MySql & " from AffidamentoRichiestaComp a, AffidamentoRichiesta B "
    MySql = MySql & " where A.IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
    MySql = MySql & " and A.IdAffidamentoRichiesta = B.IdAffidamentoRichiesta"
    IdAccountCliente = LeggiCampo(MySql,"IdAccountCliente")
    
    msgErrore=caricaImportoAffidamento("V",IdAccountCliente,IdCompagnia,IdFornitore,0,ValidoDalComp,ValidoAlComp,ValoreAffidamento,ImptSingolaPolizza,AffidamentoUsato)
    if msgErrore="" then 
       qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
       qUpd = qUpd & " IdStatoAffidamento = 'AFFI' "
       qUpd = qUpd & ",DataChiusura = " & NumForDb(Dtos())
       qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
       qUpd = qUpd & " and   IdStatoAffidamento = 'LAVO' "
       qUpd = qUpd & " and   IdFlussoProcesso in ('GES_FORNITORE','CAM_FORNITORE') "
       ConnMsde.execute qUpd
       
	   IdProdotto=GetProdottoByTipoComp("CAUZ_PROV",IdCompagnia)
       XX=createEvento("AFFI","ACCE",Session("LoginIdAccount"),"Affidamento per la Compagnia:" & DescCompagnia,"AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,IdProdotto)         

       xx=AggiornaRichiestaAffidamento(IdAffidamentoRichiesta)
       msgErrore=caricaImportoAffidamento("I",IdAccountCliente,IdCompagnia,IdFornitore,0,ValidoDalComp,ValidoAlComp,ValoreAffidamento,ImptSingolaPolizza,AffidamentoUsato)
       
       if MsgErrore="" then 
          response.redirect RitornaA(paginaReturn)
       end if 
    else 
	   ReloadImpt = false 
    end if 
 end if 
'selezionato dal cassetto lo associo 

   xx=setValueOfDic(Pagedic,"PaginaReturn"               ,PaginaReturn)
   xx=setValueOfDic(Pagedic,"IdAccountCliente"           ,IdAccountCliente)
   xx=setValueOfDic(Pagedic,"IdAffidamentoRichiestaComp" ,IdAffidamentoRichiestaComp)
   xx=setCurrent(NomePagina,livelloPagina) 

   IdAffidamentoRichiesta=LeggiCampo("Select * from AffidamentoRichiestaComp Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,"IdAffidamentoRichiesta")

   Oggi = Dtos() 
   Set Rs = Server.CreateObject("ADODB.Recordset")
 
   MySql = ""
   MySql = MySql & " Select A.*,B.DescStatoServizio from AffidamentoRichiesta A,StatoServizio B"
   MySql = MySql & " Where A.IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
   MySql = MySql & " and   A.IdStatoAffidamento = B.IdStatoServizio"
   'response.write MySql
   err.clear 

   Rs.CursorLocation = 3 
   Rs.Open MySql, ConnMsde
   IdAccountBackOffice = Rs("IdAccountBackOffice")
   DataRichiesta       = Rs("DataRichiesta")
   IdStatoAffidamento  = Rs("IdStatoAffidamento")
   DescStatoServizio   = Rs("DescStatoServizio")
   DataChiusura        = Rs("DataChiusura")
   NoteAffidamento     = Rs("NoteAffidamento")            
   Rs.close 

   MySql = ""
   MySql = MySql & " Select A.*,B.DescStatoServizio,B.FlagStatoFinale,C.DescCompagnia "
   MySql = MySql & " from AffidamentoRichiestaComp A,StatoServizio B,Compagnia C"
   MySql = MySql & " Where A.IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
   MySql = MySql & " and   A.IdCompagnia = C.IdCompagnia"
   MySql = MySql & " and   A.IdStatoAffidamento = B.IdStatoServizio"
   'response.write MySql
   err.clear 

   Rs.CursorLocation = 3 
   Rs.Open MySql, ConnMsde
   IdAccountCliente         = Rs("IdAccountCliente")
   DescCliente              = LeggiCampo("Select * from Account Where idAccount=" & IdAccountCliente,"Nominativo" )
   idCompagniaComp          = Rs("IdCompagnia")
   DescCompagniaComp        = Rs("DescCompagnia")
   DataRichiestaComp        = Rs("DataRichiesta")
   IdStatoAffidamentoComp   = Rs("IdStatoAffidamento")
   DescStatoServizioComp     = Rs("DescStatoServizio")
   FlagStatoFinale          = Rs("FlagStatoFinale")
   DataChiusuraComp         = Rs("DataChiusura")
   NoteAffidamentoComp      = Rs("NoteAffidamento")
   NoteAffidamentoClie      = Rs("NoteAffidamentoCliente")
   if ReloadImpt = true then 
      ValoreAffidamento     = Rs("ValoreAffidamento")
      AffidamentoUsato      = Rs("AffidamentoUsato")
      ImptSingolaPolizza    = Rs("ImptSingolaPolizza")
   end if 
   FlagTrasferimento        = Rs("FlagTrasferimento")
   ValidoDalComp            = Rs("ValidoDal")
   ValidoAlComp             = Rs("ValidoAl")
   IdFornitore              = Rs("IdFornitore")
   if Cdbl(IdFornitore)>0 then 
      DescFornitore=LeggiCampo("Select * from Fornitore Where IdFornitore=" & Idfornitore,"DescFornitore")
   else
      DescFornitore=""
   end if 
   IdFlussoProcesso        = Rs("IdFlussoProcesso")
   NumCoobbligatiRichiesti = Rs("NumCoobbligatiRichiesti")
   Rs.close 
   
   Set DizProcesso = CreateObject("Scripting.Dictionary")
   esito=getInfoProcessoAffi(DizProcesso,IdStatoAffidamentoComp,IdFlussoProcesso)
   if esito = true then 
      descProcesso = " - " & GetDiz(DizProcesso,"DescrizioneUtente")
   else
      descProcesso = "" 
   end if 

   'posso mettere un solo importo iniziale per la compagnia 
   qSel = "" 
   qSel = qSel & " select ImptIniziale "
   qSel = qSel & " From AccountCreditoAffiTotali "
   qSel = qSel & " Where IdAccount=" & IdAccount  
   qSel = qSel & " and IdCompagnia=" & IdCompagnia
   impt = "0" & LeggiCampo(qSel,"ImptIniziale")
   if Cdbl(impt)=0 then 
      primaVoltaCompagnia = true 
   else 
      primaVoltaCompagnia = false 
   end if 
   richiedimotivo=true 
   flagModifica=true 
   readOnly=""
   FlagStatoFinale = LeggiCampo("Select * from StatoServizio where IdStatoServizio='" & IdStatoAffidamentoComp & "'","FlagStatoFinale")
   if Cdbl("0"& flagStatoFinale)=1 then 
      richiedimotivo=false
	  flagModifica=false 
	  readOnly = " readonly "
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
                <div class="col-11"><h3>Richiesta Di Affidamento per Fornitore</h3>
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
                     <input type="text" readonly class="form-control" value="<%=DescStatoServizioComp & descProcesso%>" >
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
            <%if richiedimotivo or flagModifica=false then%>
            <div class="row">
               <div class="col-12">
                  <div class="form-group ">
                     <%xx=ShowLabel("Annotazioni relative alla richiesta")%>
                     <input type="text" <%=readOnly%> id="NewDescStato0" name="NewDescStato0" class="form-control" value="<%=NoteAffidamentoComp%>" >
                  </div>        
               </div>
            </div>
            <div class="row">
               <div class="col-12">
                  <div class="form-group ">
                     <%xx=ShowLabel("Annotazioni per il cliente")%>
                     <input type="text" <%=readOnly%> id="NewDescStatoClie0" name="NewDescStatoClie0" class="form-control" value="<%=NoteAffidamentoClie%>" >
                  </div>        
               </div>
            </div>            
            
            <%end if %>
            <div class="row">
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Importo Affidato")%>
                     <input type="text" <%=readOnly%> class="form-control" Id="ValoreAffidamento0" name="ValoreAffidamento0" value="<%=ValoreAffidamento%>" >
                  </div>        
               </div>
			   <%if primaVoltaCompagnia and cdbl(FlagTrasferimento)=99 and IdFlussoProcesso="CAM_FORNITORE" then %>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Di cui gi&agrave; usato")%>
                     <input type="text" <%=readOnly%> class="form-control" Id="AffidamentoUsato0" name="AffidamentoUsato0" value="<%=AffidamentoUsato%>" >
                  </div>        
               </div> 
               <%elseif cdbl(FlagTrasferimento)=99 and IdFlussoProcesso="GES_FORNITORE" then %>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Di cui gi&agrave; usato")%>
                     <input type="text" readonly class="form-control" Id="AffidamentoUsato0" name="AffidamentoUsato0" value="<%=AffidamentoUsato%>" >
                  </div>        
               </div> 
			   <%else %>
			      <input type="hidden" Id="AffidamentoUsato0" name="AffidamentoUsato0" value="<%=AffidamentoUsato%>" >
			   <%end if %>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Importo Max Cauzione")%>
                     <input type="text" <%=readOnly%> class="form-control" Id="ImptSingolaPolizza0" name="ImptSingolaPolizza0" value="<%=ImptSingolaPolizza%>" >
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
					 if readonly="" then 
					    clDt = "mydatepicker"
					 else 
					    clDt = ""
					 end if 
                     %>
                     <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                            class="form-control <%=clDT%> " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >                         
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
                            class="form-control <%=clDT%> " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >                         
                  </div>        
               </div>               
            </div>
			<br>

			<%if readonly="" then%>
            <div class="row">
			   <%if cdbl(FlagTrasferimento)=0 or (cdbl(FlagTrasferimento)=99 and instr("CAM_FORNITORE",IdFlussoProcesso)>0) then %>
               <div class="col-1">
                 <div class="form-group ">
                      <button type="button" onclick="gesForn('CONFERMA')" class="btn btn-success">Affida</button>
                 </div>               
               </div>  
			   <%end if %>
			   <%if cdbl(FlagTrasferimento)=0 then 
			        if instr("GES_FORNITORE",IdFlussoProcesso)>0 then %>
               <div class="col-2">
                 <div class="form-group ">
                      <button type="button" onclick="trasferisciConf()" class="btn btn-success">Trasferisci</button>
                 </div>               
               </div>
			      <%end if %>
               <%elseif instr("CAM_FORNITORE CAC_FORNITORE",IdFlussoProcesso)>0 and cdbl(FlagTrasferimento)<>99 then %>   
               <div class="col-2">
                 <div class="form-group ">
                      <button type="button" onclick="trasferisci()" class="btn btn-success">Trasferisci</button>
                 </div>               
               </div>
               <%end if %>   
			   <%if cdbl(FlagTrasferimento)=0 then %>
               <div class="col-2">
                 <div class="form-group ">
                      <button type="button" onclick="cambiaForn()" class="btn btn-info">Cambia Fornitore</button>
                 </div>               
               </div>  
               
               <div class="col-2">
                 <div class="form-group ">
                      <button type="button" onclick="integra()" class="btn btn-warning">Integra Documenti</button>
                 </div>               
               </div>  
			   <%end if %>              
               <div class="col-2">
                 <div class="form-group ">
                      <button type="button" onclick="gesForn('ANNULLA')" class="btn btn-danger">Annulla Richiesta</button>
                 </div>               
               </div>                
            </div>
			<%end if %>

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
   qListaDoc = qListaDoc & " from AffidamentoRichiestaCompDoc B,Documento C  "
   qListaDoc = qListaDoc & " Where B.IdDocumento = C.IdDocumento "
   qListaDoc = qListaDoc & " and   B.IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
   
   'response.write qListaDoc    
end if 
%>
            <!--#include virtual="/gscVirtual/configurazioni/clienti/Affidamento/ClienteFornitoreLista.asp"-->
			<%
if IsBackOffice() then 
   CoobPresentiTutti = false 
   CoobPresenteDocum = false 
   CoobPresenteValid = false
   ShowElencoCoob    = false  
   OpDocAmm          = "VAL_COOB"
%>
      <!--#include virtual="/gscVirtual/configurazioni/clienti/Affidamento/CoobbligatiLista.asp"-->

<%
end if 
%>

            
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
