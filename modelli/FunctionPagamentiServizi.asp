<%
Function UpdatePagaAccount(idAccount,idTipoCredito)
Dim v_id , MySql , Esito , v_ret
    on error resume next 
    err.clear
    v_ret = true
    MySql = "updateCreditoAccount " & idAccount & ",'" & idTipoCredito & "'"
    ConnMsde.execute MySql
    if err.number<>0 then 
       xx=writeTrace("UpdatePagaAccount:" & MySql & " " & Err.description)
       v_ret = false 
    end if 
    UpdatePagaAccount = v_ret
End function 

Function nuovoMovEco(IdAccount,IdTipoCredito,IdStatoCredito,DescStatoCredito,DescMovEco,ImptMovEco,SistemaSorgente,Segno,IdServizioRichiesto,DataFineValidita,OraFineValidita)
Dim MyQ,retV
   retV=0
   MyQ = ""
   MyQ = MyQ & " insert into AccountMovEco"
   MyQ = MyQ & "(IdAccount,IdAccountGestore,IdTipoCredito,IdStatoCredito,DescStatoCredito"
   MyQ = MyQ & ",DescMovEco,DataMovEco,TimeMovEco,ImptMovEco"
   MyQ = MyQ & ",FlagStorico,NoteStorico,IdUpload,SistemaSorgente,SegnoSistema "
   MyQ = MyQ & ",IdServizioRichiesto,DataFineValidita,OraFineValidita) values"
   MyQ = MyQ & "(" & IdAccount & ",0,'" & IdTipoCredito & "','" & IdStatoCredito & "','" & apici(DescStatoCredito) & "'"
   MyQ = MyQ & ",'" & apici(DescMovEco) & "',"  & Dtos() & "," & TimeToS() & "," & NumForDb(ImptMovEco)
   MyQ = MyQ & ",0,'',0,'" & apici(sistemaSorgente) & "'," & Segno & "," & IdServizioRichiesto 
   MyQ = MyQ & "," & NumForDb(DataFineValidita) & "," & NumForDb(OraFineValidita)
   MyQ = MyQ & ")" 
   err.clear 
   on error resume next 
   ConnMsde.execute MyQ 
   if Err.number=0 then 
      retV=getTableIdentity("AccountMovEco")
   else
      xx=writeTrace("nuovoMovEco:" & MyQ & " " & Err.description)
      retV=0
   end if 
   nuovoMovEco = retV
end function 
   
'effettura il pagamento di un servizio per la catena
'il movimento deve essere non contabilizzato
function pagaServizioRichiesto(IdAttivita,IdNumAttivita) 
Dim paga_debug,moduloC 
Dim paga_RetVal,MyRs 
Dim DizPag,IdAccCli,IdAccRic,IdAccL1,IdAccL2
Dim idProdotto,DescProdotto,ImptTotaleServizio,IdTipoCreditoSel
Dim qSel,tmpStr,tmpInt
Dim FlagBorsellino,FlagFido,FlagEstratto,FlagPagaOk,ImptDisp,FlagPagaAll
Dim IdServizioRichiesto,IdStatoServizio 
Dim IdProcessoElaborativo

Dim dizServRich 
   moduloC = "FunctionPagamentiServizi:"
   paga_RetVal=""
   paga_debug = false 
   IdStatoServizio = "PAGA"   
 
   xx=writeTraceAttivita(moduloC & "Inizio Pagamento",IdAttivita,IdNumAttivita)
   set dizServRich = GetDizServizioRichiesto(0,IdAttivita,IdNumAttivita)
   
   tmpStr = "0" & Getdiz(dizServRich,"N_IdServizioRichiesto")
   if Cdbl(tmpStr)=0 then 
      paga_RetVal="KO:Servizio assente"
   end if 
   if paga_RetVal="" then 
      tmpStr = "0" & Getdiz(dizServRich,"N_FlagContabilizzato")
      if Cdbl(tmpStr)=1 then 
         paga_RetVal="KO:Movimento contabilizzato "
      end if 
   end if 
   'cancello eventuali dati precedenti e aggiorno
   IdServizioRichiesto = GetDiz(dizServRich,"N_IdServizioRichiesto")
   if paga_RetVal="" then 
      xx=cambiaStatoMovEco(IdServizioRichiesto,"ANNU")
   end if       
   
   if paga_RetVal="" then 
      
      if Cdbl("0" & GetDiz(dizServRich,"N_IdAccountLivello1"))=0 then 
         xx=writeTraceAttivita("Controllo livelli ",IdAttivita,IdNumAttivita)
         'se non ci sono i livelli li carico e rileggo 
         xx = verifyAccountServizioRichiesto(IdAttivita,IdNumAttivita)
         set dizServRich = GetDizServizioRichiesto(0,IdAttivita,IdNumAttivita)
      end if 
      
      IdAccCli           = Cdbl("0" & GetDiz(dizServRich,"N_IdAccountCliente"))
      IdAccRic           = Cdbl("0" & GetDiz(dizServRich,"N_IdAccountRichiedente"))
      IdAccL1            = Cdbl("0" & GetDiz(dizServRich,"N_IdAccountLivello1"))
      IdAccL2            = Cdbl("0" & GetDiz(dizServRich,"N_IdAccountLivello2"))
      
      IdProcessoElaborativo = GetDiz(dizServRich,"S_IdProcessoElaborativo")    
      IdProdotto         = Cdbl("0" & GetDiz(dizServRich,"N_IdProdotto"))
      
      DescProdotto       = LeggiCampo("select * from Prodotto Where IdProdotto=" & IdProdotto,"DescProdotto")
      ImptTotaleServizio = Cdbl("0" & GetDiz(dizServRich,"N_ImptTotaleServizio"))
      
      'tipo di pagamento richiesto dal richiedente 
      IdTipoCreditoClie = GetDiz(dizServRich,"S_IdTipoCreditoClie")
      IdTipoCreditoRequ = GetDiz(dizServRich,"S_IdTipoCreditoRequ")

      Dim posX
      posX = 0
      Dim IdAccList(10)
      Dim IdAccPaga(10)
      Dim IdAccErro(10)
      for t=1 to 9
         IdAccList(t) = 0
         IdAccPaga(t) = ""
         IdAccErro(t) = ""
      next 
      'se il richiedente non è il cliente devo controllare prima lui 
      if cdbl(IdAccRic)<>cdbl(IdAccCli) then 
         IdAccList(1) = IdAccRic
         IdAccPaga(1) = IdTipoCreditoRequ
         IdAccList(2) = IdAccCli
         IdAccPaga(2) = IdTipoCreditoClie
         posX = 2
      else 
         IdAccList(1) = IdAccCli
         IdAccPaga(1) = IdTipoCreditoRequ
         posX = 1
      end if

      'azzero eventuali segnalatori che non hanno modalità di pagamento 
      'controllo dall'ultimo che dovrebbe essere il referente del cliente
      qSel = "SELECT IdTipoCollaboratore FROM Collaboratore Where IdAccount=" 
      if Cdbl(IdAccL2)>0 and cdbl(IdAccL2)<>cdbl(IdAccRic) then 
         tmpStr = ucase(LeggiCampo(qSel & IdAccL2 ,"IdTipoCollaboratore"))
         if paga_debug then 
            response.write "controllo account SEGN=" & qSel & IdAccL2 & " ->" & tmpStr & "<br>"
         end if          
         if tmpStr<>"SEGN" then 
            posX=posX+1 
            IdAccList(posX) = IdAccL2
         end if 
      end if 
      if Cdbl(IdAccL1)>0 and cdbl(IdAccL1)<>cdbl(IdAccRic) then 
         tmpStr = ucase(LeggiCampo(qSel & IdAccL1 ,"IdTipoCollaboratore"))
         if paga_debug then 
            response.write "controllo account SEGN=" & qSel & IdAccL1 & " ->" & tmpStr & "<br>"
         end if          
         
         if tmpStr<>"SEGN" then 
            posX=posX+1 
            IdAccList(posX) = IdAccL1
         end if 
      end if 
   end if 
   for t=1 to posX
       xx=writeTraceAttivita("Controllo pagamenti " & IdAccList(t) & " " & IdAccPaga(t),IdAttivita,IdNumAttivita)   
   next 
  
   FlagPagaAll = true
   if paga_RetVal="" then
      'si verifica la disponibilità economica di ogni account 
    
      Set DizPag = CreateObject("Scripting.Dictionary")   
   
      qSel = "select * from AccountModPag Where IdAccount = "
      
      for t=1 to posX
         DizPag.RemoveAll
         tmpStr="0" & LeggiCampo(qSel & IdAccList(t),"IdAccount")
         xx=writeTraceAttivita(moduloC & "leggo pagamenti " & qSel & IdAccList(t),IdAttivita,IdNumAttivita)
         'esistono dati per permettere il pagamento
         if cdbl(tmpStr)>0 then 
            xx=writeTraceAttivita(moduloC & "trovato pagamenti ",IdAttivita,IdNumAttivita)
         'carico lo stato dei pagamenti
            xx=GetInfoRecordset(DizPag,qSel & IdAccList(t))   
            FlagBorsellino = TestNumeroPos(GetDiz(DizPag,"FlagBorsellino"))
            FlagFido       = TestNumeroPos(GetDiz(DizPag,"FlagFido"))
            FlagEstratto   = TestNumeroPos(GetDiz(DizPag,"FlagEstratto"))
            FlagEsterno    = 0 

            'devo controllare solo quel tipo di pagamento
            if IdAccPaga(t)<>"" then 
               if IdAccPaga(t)="BORS" then 
                  FlagFido       = 0
                  FlagEstratto   = 0
               end if 
               if IdAccPaga(t)="FIDO" then
                  FlagBorsellino = 0
                  FlagEstratto   = 0
               end if 
               if IdAccPaga(t)="ESTR" then 
                  FlagBorsellino = 0
                  FlagFido       = 0
               end if 
               IdAccPaga(t) = ""
            else
               xx=writeTraceAttivita("non ho scelto un pagamento ",IdAttivita,IdNumAttivita)
            end if 
            FlagPagaOk = false 
            if FlagPagaOk=false and Cdbl(FlagBorsellino)>0 then 
               ImptDisp = TestNumeroPos(GetDiz(DizPag,"ImptBorsellinoDisp"))
               if Cdbl(ImptDisp)>=cdbl(ImptTotaleServizio) then 
                  IdAccPaga(t)="BORS"
                  FlagPagaOk=true
               end if 
            end if 
            if FlagPagaOk=false and Cdbl(FlagFido)>0 then 
               ImptDisp = TestNumeroPos(GetDiz(DizPag,"ImptFidoDisp"))
               if Cdbl(ImptDisp)>=cdbl(ImptTotaleServizio) then 
                  IdAccPaga(t)="FIDO"
                  FlagPagaOk=true
               end if 
            end if 
            if FlagPagaOk=false and Cdbl(FlagEstratto)>0 then 
               ImptDisp = TestNumeroPos(GetDiz(DizPag,"ImptEstrattoDisp"))
               if Cdbl(ImptDisp)>=cdbl(ImptTotaleServizio) then 
                  IdAccPaga(t)="ESTR"
                  FlagPagaOk=true
               end if 
            end if  
            if FlagPagaOk=false and Cdbl(FlagEsterno)>0 then 
               IdAccPaga(t)="ESTE"
               FlagPagaOk=true
            end if             
            if FlagPagaOk=false then
               FlagPagaAll=false 
               IdAccErro(t)="Importo insufficiente al pagamento "
               'forzo uscita : non è possibile pagare
               t = posX               
            end if 
         else
            IdAccErro(t)="Nessuna modalita' di pagamento prevista"
         end if 
      next
   end if   
   
   'tutti ok : inserisco gli importi e aggiorno totali 
   flagInternalError=false 
   'diventa true se si paga con borsellino
   'carico solo pagamenti informativi che servono per fare calcoli
   flagNextInfo     =false 
   if paga_RetVal="" and flagPagaAll=true then   
      'genero il movimento padre    
	  ContaPaga = 0
	  xx=UpdateAccoutPagatoreServizioRichiesto(0,IdAttivita,IdNumAttivita,0,"","",ContaPaga)
      for t=1 to posX
          if IdAccPaga(t)<>"" then 
		     ContaPaga = ContaPaga + 1 
             xx=writeTraceAttivita("creo il movimento economico ",IdAttivita,IdNumAttivita)
             if flagNextInfo=true then 
                IdStatoCredito="INFO"
             else
                IdStatoCredito="IMPE"
             end if 
             idMovEco = nuovoMovEco(IdAccList(t),IdAccPaga(t),IdStatoCredito,"",DescProdotto,ImptTotaleServizio,"ACQ",-1,IdServizioRichiesto,Dtos(),190000)
             if Cdbl(IdMovEco)=0 then 
                xx=writeTraceAttivita(moduloC & "movimento economico non creato ",IdAttivita,IdNumAttivita)
                IdAccErro(t)="Impossibile registare il movimento "
                t = posX 
                flagPagaAll = false 
                flagInternalError=true 
             else 
                xx=writeTraceAttivita(moduloC & "movimento economico creato = " & idMovEco,IdAttivita,IdNumAttivita)
                if IdStatoCredito<>"INFO" then 
                   TmpStr=UpdatePagaAccount(IdAccList(t),IdAccPaga(t))
                end if 
		        xx=UpdateAccoutPagatoreServizioRichiesto(0,IdAttivita,IdNumAttivita,IdAccList(t),IdAccPaga(t),IdStatoCredito,ContaPaga)
                
             end if 
         end if 
         'pagato con borsellino : non devo più scalare
         if IdAccPaga(t)="BORS" then
            flagNextInfo = true 
         end if 
      next 
   end if 
   for t=1 to posX
      xx=writeTraceAttivita("esito pagamenti: " & "IdAcc=" & IdAccList(t) & " tipoPaga=" & IdAccPaga(t) & " err=" & IdAccErro(t),IdAttivita,IdNumAttivita)   
   next 
   if err.number<>0 then 
      xx=writeTraceAttivita("sono in errore 25 : " & err.description,IdAttivita,IdNumAttivita)   
   end if
   
   'in errore per cui cancello se caricato ed aggiorno i totali
   if paga_RetVal="" and flagPagaAll=false then 
      xx=writeTraceAttivita("sono in errore 26 : " & paga_RetVal,IdAttivita,IdNumAttivita)   
      paga_RetVal="KO:pagamento non possibile"
	  
	  'devo stornare tutti i pagamenti
	  xx=stornaPagamenti(IdServizioRichiesto)

      'notifico cosa accade all'account se non è un errore interno
      if flagInternalError=false then 
         'ciclo per decidere cosa notificare 
         flagNotificato = false 
         for t=1 to posX
             if paga_debug then 
                response.write "cerco errore 11:IdAcc=" & IdAccList(t) & " tipoPaga=" & IdAccPaga(t) & " err=" & IdAccErro(t) & "<br>"
             end if 
   
             'notifico a video : è sul richiedente il problema
             if IdAccErro(t)<>"" and cdbl(IdAccList(t))=cdbl(IdAccRic) and flagNotificato = false  then 
                paga_RetVal = IdAccErro(t)
                flagNotificato = true
                if paga_debug then 
                   response.write "Errore sul richiedente:" & IdAccErro(t)
                end if
             end if 
             'notifico a video : è sul cliente il problema
             if IdAccErro(t)<>"" and cdbl(IdAccList(t))=cdbl(IdAccCli) and flagNotificato = false  then 
                paga_RetVal = IdAccErro(t)
                flagNotificato = true
                if paga_debug then 
                   response.write "Errore sul cliente:" & IdAccErro(t)
                end if
             end if 
         next 
         if paga_debug and err.number<>0 then 
            response.write "sono in errore 20 : " & err.description
         end if
         'il problema non è sul richiedente ne sul cliente ma ad un livello superiore
         ' notifico di contattare il suo responsabile. 
         if flagNotificato = false then 
            paga_RetVal = "pagamento al momento non possibile: la preghiamo di contattare il suo gestore."
            'notifico all'utente in errore cosa e' successo 
            for t=1 to posX
                if paga_debug then 
                   response.write "cerco errore 22:IdAcc=" & IdAccList(t) & " tipoPaga=" & IdAccPaga(t) & " err=" & IdAccErro(t) & "<br>"
                end if 
            
                if IdAccErro(t)<>"" then 
                   IdTabella=""
                   IdKey=""
                   if GetDiz(DizMov,"IdAnagServizio")="CAUZ_PROV" then 
                      IdTabella="Cauzione"
                      IdKey="IdCauzione=" & GetDiz(DizMov,"keyServizio")
                   end if 
                   DescEvento = IdAccErro(t) & ":" & DescProdotto
                   if paga_debug then 
                      response.write "creo evento per IdAcc=" & IdAccList(t) & " descr=" & DescEvento & "<br>"
                   end if 
                   'pagamento fallito  
                   xx=createEvento("PAGA","PFAL",IdAccList(t),DescEvento,IdTabella,IdKey,true,IdProdotto)
                   if paga_debug then 
                      response.write "creato evento" & "<br>"
                   end if 
                   err.clear
                end if 
            next
         end if 
      end if 
   end if 
   if paga_debug then 
      response.write "fine : " & err.description & " esito=" & paga_RetVal
   end if
   if paga_RetVal<>"" then 
      xx=writeTraceAttivita("pagaServizioRichiesto esito:" & paga_RetVal,IdAttivita,IdNumAttivita) 
   else 
      xx=UpdateStatoServizioRichiesto(IdAttivita,IdNumAttivita,IdStatoServizio,"")
      xx=writeTraceAttivita("aggiorno stato :" & IdStatoServizio,IdAttivita,IdNumAttivita) 
      ConnMsde.execute qUpd   
	  ConnMsde.execute "ServizioRichiesto_creaMovimentoAcquisto 0,'" & IdAttivita & "'," & IdNumAttivita
   end if 
   pagaServizioRichiesto=paga_RetVal
 
end function 

function stornaPagamenti(IdServizioRichiesto)
Dim MySql,RsQ,retVal,IdProcessoElaborativo,ModuloC,qUpd 
   on error resume next 
   ModuloC = "stornaPagamenti:" 
   MySql = "select * From AccountMovEco Where IdServizioRichiesto=" & IdServizioRichiesto
   xx=writeTraceAttivita(ModuloC & MySql,"IdServizioRichiesto",IdServizioRichiesto)  
   Set RsQ = Server.CreateObject("ADODB.Recordset")
   RsQ.CursorLocation = 3 
   RsQ.Open MySql, ConnMsde
   if err.number<>0 then 
      xx=writeTraceAttivita(ModuloC & Err.Description,"IdServizioRichiesto",IdServizioRichiesto) 
   else
      Do While Not RsQ.EOF
		 qUpd = "update AccountMovEco set IdStatoCredito='ANNU' Where IdAccountMovEco=" & RsQ("IdAccountMovEco")
	     ConnMsde.execute qUpd 
	     TmpStr=UpdatePagaAccount(RsQ("IdAccount"),RsQ("IdTipoCredito"))
         RsQ.MoveNext
     Loop
     RsQ.Close
     ConnMsde.execute "delete From AccountMovEco Where IdStatoCredito='ANNU' and IdServizioRichiesto=" & IdServizioRichiesto
   end if 

end function 


function confermaPagamenti(IdServizioRichiesto)
Dim MySql,RsQ,retVal,IdProcessoElaborativo,ModuloC,qUpd 
   on error resume next 
   ModuloC = "confermaPagamenti:" 
   MySql = "select * From AccountMovEco Where IdStatoCredito = 'IMPE' and IdServizioRichiesto=" & IdServizioRichiesto
   xx=writeTraceAttivita(ModuloC & MySql,"IdServizioRichiesto",IdServizioRichiesto)  
   Set RsQ = Server.CreateObject("ADODB.Recordset")
   RsQ.CursorLocation = 3 
   RsQ.Open MySql, ConnMsde
   if err.number<>0 then 
      xx=writeTraceAttivita(ModuloC & Err.Description,"IdServizioRichiesto",IdServizioRichiesto) 
   else
      Do While Not RsQ.EOF
		 qUpd = "update AccountMovEco set IdStatoCredito='UTIL' Where IdAccountMovEco=" & RsQ("IdAccountMovEco")
	     ConnMsde.execute qUpd 
	     TmpStr=UpdatePagaAccount(RsQ("IdAccount"),RsQ("IdTipoCredito"))
         RsQ.MoveNext
     Loop
     RsQ.Close
   end if 

end function

function cambiaStatoMovEco(IdServizioRichiesto,NewStato)
Dim qSel ,MyRs , tmpInt ,tmpstr , xx 

   on error resume next 
   'cambio stato
   qSel = ""
   qSel = qSel & " update AccountMovEco set IdStatoCredito='" & newStato & "'"
   qSel = qSel & " Where IdServizioRichiesto>0 and IdServizioRichiesto=" & IdServizioRichiesto 
   ConnMsde.execute qSel 
   'leggo i dettagli ed aggiorno  
   qSel = ""
   qSel = qSel & " select * From AccountMovEco Where IdServizioRichiesto>0 and IdServizioRichiesto=" & IdServizioRichiesto

   Set MyRs = Server.CreateObject("ADODB.Recordset")
   MyRs.CursorLocation = 3 
   MyRs.Open qSel, ConnMsde      
   if MyRs.eof = false then 
      do while not MyRs.eof 
         tmpInt = cdbl("0" & MyRs("IdAccount"))
         tmpstr = MyRs("IdTipoCredito")
         if NewStato="ANNU" then 
            connMsde.execute "Delete from AccountMovEco Where IdAccountMovEco = " & MyRs("IdAccountMovEco")
         end if 
         xx=UpdatePagaAccount(tmpInt,tmpstr)
         MyRs.moveNext 
      loop

   end if 
   MyRs.close 
   err.clear
   
end function 

%>