<%

Function GetProdottoByTipoComp(IdTipoServizio,IdCompagnia)
dim retVal,q 
   retVal = 0
   q = ""
   q = q & " select * from Prodotto "
   q = q & " Where IdCompagnia=" & NumForDb(IdCompagnia) 
   q = q & " and IdAnagServizio='" & apici(IdTipoServizio) & "'"
   
   retVal = cdbl("0" & LeggiCampo(q,"IdProdotto"))
   GetProdottoByTipoComp = retVal 
end function 

'si crea un evento per l'account che lo genera 
Function createEvento(IdProcesso,IdTipoEvento,IdAccountEvento,DescEvento,IdTabella,IdKey,sendNotify,IdProdotto)
Dim qSel,qIns,IdAcc,idEvento,dtEv,tmEV
   on error resume next 
   idAcc    = TestNumeroPos(IdAccountEvento) 
   dtEv     = Dtos()
   tmEv     = TimeTos()
   idEvento = 0
   qIns = ""
   qIns = qIns & "INSERT INTO Evento (IdProcesso,IdTipoEvento,IdAccountEvento,DataEvento,TimeEvento"
   qIns = qIns & ",DescEvento,IdTabella,IdKey,IdProdotto)  VALUES"
   qIns = qIns & "('" & apici(idProcesso) & "','" & apici(IdTipoEvento) & "'," & idAcc & "," & dtEv & "," & tmEv
   qIns = qIns & ",'" & apici(descEvento) & "','" & apici(IdTabella) & "','" & apici(idKey) & "'"
   qIns = qIns & ", " & NumForDb(IdProdotto)
   qIns = qIns & ")"
   'response.write qIns 
   ConnMsde.execute qIns 
   if err.number = 0  then
      qSel = ""   
      qSel = qSel & " select IdEvento from Evento "
      qSel = qSel & " Where IdProcesso = '" & Apici(idProcesso) & "'"
      qSel = qSel & " and IdTipoEvento = '" & Apici(IdTipoEvento) & "'"
      qSel = qSel & " and IdAccountEvento = " & idAcc
      qSel = qSel & " and DataEvento = " & dtEv
      qSel = qSel & " and TimeEvento = " & tmEv
      idEvento = cdbl("0" & LeggiCampo(qSel,"IdEvento"))
   else
      err.clear
   end if 
   
   'evento caricato eseguo la logica di storico 
   if cdbl(idEvento)>0 and Cdbl(idAcc)>0 then 
      flagCreateAcc=true 
      'da pagamento arriva a chi notificare via mail 
      if IdProcesso="PAGA" and IdTipoEvento="PFAL" and sendNotify=true then
         flagCreateAcc=false
      end if 
      if flagCreateAcc=true then 
         AddInfoAcc = ""
         xx = createEventoAccount(idAcc,IdEvento,"","","out",dtEv,tmEv,AddInfoAcc)
      end if 
      'creo gli eventi da spedire 
      if sendNotify=true then 
         xx = createEventoMailing(IdAcc,IdEvento,IdProcesso,IdTipoEvento,IdTabella,IdKey,IdProdotto)
         'spedisco gli eventi generato 
         xx = elaboraEventoMailing(IdEvento)
      end if 
   end if 
   err.clear
end Function    

Function createEventoAccount(idAccount,IdEvento,emailNotifica,telNotifica,direzione,dataNotifica,timeNotifica,DescEventoAccount)
Dim qIns 
   on error resume next 
   qIns = ""
   qIns = qIns & " INSERT INTO EventoAccount(IdAccount,IdEvento,emailNotifica,telNotifica,direzione"
   qIns = qIns & ",dataNotifica,timeNotifica,DescEventoAccount)"
   qIns = qIns & " values (" & idAccount & "," & IdEvento
   qIns = qIns & ",'" & apici(emailNotifica) & "','" & apici(telNotifica) & "','" & apici(direzione) & "'"
   qIns = qIns & "," & dataNotifica & "," & timeNotifica & ",'" & apici(DescEventoAccount) & "'"
   qIns = qIns & ")"
   'response.write qIns 
   ConnMsde.execute qIns
   err.clear

end Function 

'in base al tipo di evento si decide la notifica : si cercano tutti gli account 
'oggetto di notifica 
function createEventoMailing(IdAccountEvento,IdEvento,IdProcesso,IdTipoEvento,IdTabella,IdKey,IdProdotto)
dim idP,idT,qElenco,notificaAccount1,notificaAccount2,notificaRichiedente,notificaGestore
dim qSel,myRS,emailNotifica,telNotifica,xx
dim IdAccountCliente,IdAccountGestore,IdAccountRichiedente,IdAccountCollaboratore,IdServizioRichiesto
dim mailNotificaCliente,mailNotificaRichiedente,mailNotificaBackOffice
'sono le mail passate per invio : default CERCA per cercare in base as DB 
dim mail_NotificaAccount1,mail_NotificaAccount2,mail_notificaRichiedente,mail_notificaGestore
dim idStatoServizioPrec 
dim InfoAccountRel,InfoAccountGestore,InfoAccount1,InfoAccount2,InfoAccountRichiedente
   
   Set myRS = Server.CreateObject("ADODB.Recordset")
   myRS.CursorLocation = 3 
      
'        IdTipoEvento    DescTipoEvento
'        ACCE    Accettazione
'        ANNU    Annullamento
'        CANC    Cancellazione
'        CARI    Caricamento
'        CHIU    Chiusura
'        INTE    Integrazione
'        INVI    Invio
'        PREC    Presa in Carico
'        RICH    Richiesta
'        VALI    Validazione
'        VARI    Variazione
'       PAGA    Pagamento
   on error resume next 
   idP=ucase(trim(IdProcesso))
   idT=ucase(trim(IdTipoEvento))
   
   IdAccountCliente       = 0
   IdAccountGestore       = 0
   IdAccountRichiedente   = 0
   IdAccountCollaboratore = 0
   IdServizioRichiesto    = 0
   
   idStatoServizioPrec    = ""
   'sono prelevate da ServizioRichiesto 
   mailNotificaCliente    = ""
   mailNotificaRichiedente= ""
   mailNotificaBackOffice = ""

   mail_NotificaAccount1    = "CERCA"
   mail_NotificaAccount2    = "CERCA"
   mail_notificaRichiedente = "CERCA"
   mail_notificaGestore     = "CERCA"

   InfoAccount1             = ""
   InfoAccount2             = ""
   InfoAccountRichiedente   = ""
   InfoAccountGestore       = ""
   
   ListaAccountBack    = ""
   notificaRichiedente = 0
   notificaAccount1    = 0
   notificaAccount2    = 0
   notificaGestore     = 0
   MailDocumentazione  = ""
   AddInfoAccClie      = ""   
   if idP="AFFI" then
      'recupero cliente,collaboratore,backoffice : la richiesta è per compagnia 
      'escludo IdAccountEvento perchè gia' notificato
      qSel = ""
      qSel = qSel & " select * From AffidamentoRichiesta where IdAffidamentoRichiesta in "
      qSel = qSel & " (select IdAffidamentoRichiesta from AffidamentoRichiestaComp where "
      qSel = qSel & IdKey & ")"
      'response.write Qsel 
      IdAccountCliente     = cdbl("0" & LeggiCampo(qSel,"IdAccountCliente"))
      IdAccountRichiedente = cdbl("0" & LeggiCampo(qSel,"IdAccountRichiedente"))
      if cdbl(IdAccountCliente)=cdbl(IdAccountEvento) then 
         IdAccountCliente=0
      end if 
      if cdbl(IdAccountRichiedente)=cdbl(IdAccountEvento) then 
         IdAccountRichiedente=0
      end if 
      ListaAccountBack=""
      if idT="RICH" then 'richiesta : notifica a tutti i back office 
         notificaAccount1 = IdAccountCliente
         notificaAccount2 = IdAccountRichiedente
         ListaAccountBack = getAccountsBackOffice()
      end if 
      if idT="INTE" or idT="DOCU" then ' notifica al cliente
         notificaAccount1 = IdAccountCliente
         notificaAccount2 = IdAccountRichiedente         
         'response.write "eccomi"
      end if       
      if idT="PREC" or idT="ACCE" then 'presa in carico : notifica al cliente e al richiedente 
         qElenco = "select * from Account Where IdAccount in ("
         qElenco = qElenco & " select IdAccountCliente From " & IdTabella
         qElenco = qElenco & " Where " & IdKey
         qElenco = qElenco & ")"
         'response.write qElenco 
      end if 
   end if 
   'notifica per coobbligati ed ATI : notifico al back office 
   if idP="COOB" or idP="ATI" then
      ListaAccountBack = getAccountsBackOffice()
   end if 
   'processi che utilizzano ServizioRichiesto
   'CAUD = cauzioni definitive 
   if IdP<>"" and instr("CAUD",idP)>0 then
  
      idAttivita             = ""
      IdNumAttivita          = 0
      if idP="CAUD" then 
         idAttivita="CAUZ_DEFI"
         IdNumAttivita=cdbl("0" & LeggiCampo("select * from CauzioneDef Where " & IdKey,"IdCauzioneDef"))
      end if 
      xx=writeTraceAttivita("evento " & IdP & " per " & idT,IdAttivita,IdNumAttivita)
      
      'estraggo idServizioRichiesto 
      qSel = ""
      qSel = qSel & " select * From ServizioRichiesto "
      qSel = qSel & " where IdAttivita='" & apici(IdAttivita) & "'"
      qSel = qSel & " and IdNumAttivita=" & NumForDb(IdNumAttivita)
      myRS.Open qSel, ConnMsde
      if err.number=0 then 
         if Not myRS.EOF then
            IdServizioRichiesto    = myRs("IdServizioRichiesto")
            IdAccountCliente       = myRS("IdAccountCliente")
            IdAccountGestore       = myRS("IdAccountGestore")
            IdAccountRichiedente   = myRS("IdAccountRichiedente")
            mailNotificaCliente    = myRS("mailNotificaCliente")
            mailNotificaRichiedente= myRS("mailNotificaRichiedente")
            mailNotificaBackOffice = myRS("mailNotificaBackOffice")
            idStatoServizioPrec    = myRS("idStatoServizioPrec")

            IdAccountCollaboratore = myRS("IdAccountlivello2")
            if Cdbl(IdAccountCollaboratore)=0 then 
               IdAccountCollaboratore = cdbl("0" & LeggiCampo(qSel,"IdAccountlivello1")) 
            end if 
         end if 
         MyRs.close 
      end if  
      'escludo chi genera l'evento perche' gia notificato 
      if cdbl(IdAccountCliente)=cdbl(IdAccountEvento) then 
         IdAccountCliente=0
      end if 
      if cdbl(IdAccountGestore)=cdbl(IdAccountEvento) then 
         IdAccountGestore=0
      end if 
      if cdbl(IdAccountRichiedente)=cdbl(IdAccountEvento) then 
         IdAccountRichiedente=0
      end if       
      if cdbl(IdAccountCollaboratore)=cdbl(IdAccountEvento) then 
         IdAccountCollaboratore=0
      end if 
      xx=writeTraceAttivita("trovato IdServizioRichiesto = " & IdServizioRichiesto,IdAttivita,IdNumAttivita)
      'eseguo le logiche di invio 
      if Cdbl(IdServizioRichiesto)>0 then 

         'caso richiesta servizio 
         '   notifica back office 
         if idT<>"" and instr("RICH",idT)>0 then 
            ListaAccountBack   = getAccountsBackOfficeByProd(IdProdotto)         
         end if 
      
         'caso presa in carico e lavorazione 
         '   nessuna notifica se non per integrazione documentazione 
         if idT<>"" and instr("LAVO",idT)>0 then 
            if idStatoServizioPrec="INTE" then 
               ListaAccountBack = "select IdAccount from Account Where IdAccount=" & IdAccountGestore
               notificaGestore  = IdAccountGestore
               mail_notificaGestore = mailNotificaBackOffice
			   InfoAccountGestore = "la documentazione e' stata integrata "
            end if 
         end if 

         'caso integrazione documentazione 
         '   notifica richiedente 
         if idT<>"" and instr("DOCF_DOCU_INTE",idT)>0 then 
             notificaRichiedente = IdAccountRichiedente
         end if 

         'caso cancellazione/annullamento/rifiuto 
         '   notifica richiedente 
         '   notifica collaboratore 
         if idT<>"" and instr("CANC_ANNU_RIFI",idT)>0 then 
             notificaRichiedente      = IdAccountRichiedente
             notificaAccount1         = IdAccountCollaboratore
         end if 
         
         'caso richiesta validazione dal cliente VALC o dal collaboratore VALP 
         if idT<>"" and instr("VALC_VALP",idT)>0 then 
            notificaRichiedente      = IdAccountRichiedente
            if idT="VALP" then 
               notificaAccount1 = IdAccountCollaboratore
            else 
               notificaAccount1 = IdAccountCliente
            end if 

         end if 

         'caso richiesta accettada dal cliente 
         '    nessuna notifica 
         if idT<>"" and instr("PACC",idT)>0 then 
         end if 
         
         'caso richiesta validazione da fornitore/back office
         '    nessuna notifica 
         if idT<>"" and instr("VALI",idT)>0 then 
         end if 

         'caso richiesta polizza al fornitore
         '    nessuna notifica 
         if idT<>"" and instr("FORN_RPOL",idT)>0 then 
         end if 

         'caso richiesta approvata dal fornitore
         '    nessuna notifica 
         if idT<>"" and instr("APPF",idT)>0 then 
         end if 
         
         'caso cessazione
         '    notifica richiedente 
         if idT<>"" and instr("CESS",idT)>0 then 
            notificaRichiedente = IdAccountRichiedente
         end if 

         'caso pagamento 
         if idT<>"" and instr("PAGA_PAGC",idT)>0 then 
            ListaAccountBack = "select IdAccount from Account Where IdAccount=" & IdAccountGestore
            notificaGestore  = IdAccountGestore
            mail_notificaGestore = mailNotificaBackOffice
         end if 

         'caso servizio attivato 
         if idT<>"" and instr("ATTI",idT)>0 then 
            notificaRichiedente = IdAccountRichiedente
            notificaAccount1    = IdAccountCollaboratore
         end if 

         'caso servizio affidato
         if idT<>"" and instr("AFFI",idT)>0 then 
         end if 

         'definisco le mail da inviare 
         if notificaRichiedente=IdAccountRichiedente then 
            mail_notificaRichiedente = mailNotificaRichiedente
         end if 
         if notificaAccount1   =IdAccountRichiedente then 
            mail_notificaAccount1    = mailNotificaRichiedente
         end if              
         if notificaAccount2   =IdAccountRichiedente then 
            mail_notificaAccount2    = mailNotificaRichiedente
         end if  
         
         if notificaRichiedente=IdAccountCliente then 
            mail_notificaRichiedente = mailNotificaCliente
         end if 
         if notificaAccount1   =IdAccountCliente then 
            mail_notificaAccount1    = mailNotificaCliente
         end if              
         if notificaAccount2   =IdAccountCliente then 
            mail_notificaAccount2    = mailNotificaCliente
         end if 
    
         xx=writeTraceAttivita("notificaGestore:" & notificaGestore & " " & mail_notificaGestore & " notificaRichiedente = " & notificaRichiedente & " " & mail_notificaRichiedente & " notificaAccount1=" & notificaAccount1 & mail_notificaAccount1 & " notificaAccount2=" & notificaAccount2 & mail_notificaAccount2 & "ListaAccountBack=" & ListaAccountBack ,IdAttivita,IdNumAttivita)
      end if 


   end if 
   
   if idP="FORM" then 
      MailDocForn = ""
      'recupero cliente,collaboratore,backoffice :  
      'escludo IdAccountEvento perchè gia' notificato
      qSel = ""
      qSel = qSel & " select * From Formazione where "
      qSel = qSel & IdKey 
      IdAccountCliente     = 0
      IdAccountRichiedente = 0
      IdAccountFornitore   = 0
      DatiAccessoForm      = ""
      
      myRS.Open qSel, ConnMsde
      if err.number=0 then 
         if Not myRS.EOF then 
            IdAccountCliente     = cdbl("0" & myRS("IdAccountCliente"))
            IdAccountRichiedente = cdbl("0" & myRS("IdAccountRichiedente"))
            IdAccountFornitore   = cdbl("0" & myRS("IdAccountFornitore"))
            'per attivazione metto le credenziali 
            if IdT="ATTI" then 
               if trim(myRS("linkPiattaforma"))<>"" then 
                  DatiAccessoForm = DatiAccessoForm & "Il corso è fruibile al seguente link:" & myRS("linkPiattaforma") & " <br>"
               end if 
               If trim(myRS("userPiattaforma"))<>"" and trim(myRS("passPiattaforma"))<>"" then 
                  DatiAccessoForm = DatiAccessoForm & "per la fruizione utilizzi le seguenti credenziali:<br>"
                  DatiAccessoForm = DatiAccessoForm & "user=<b>" & trim(myRS("userPiattaforma")) & "</b></br>"
                  DatiAccessoForm = DatiAccessoForm & "password=<b>" & trim(myRS("passPiattaforma")) & "</b><br>"
               end if 
            end if
         end if 
         MyRs.close 
      else 
         err.clear
      end if  


      if cdbl(IdAccountCliente)=cdbl(IdAccountRichiedente) then
         IdAccountRichiedente = 0
      end if 
      if cdbl(IdAccountCliente)=cdbl(IdAccountEvento) then 
         IdAccountCliente=0
      end if 
      if cdbl(IdAccountRichiedente)=cdbl(IdAccountEvento) then 
         IdAccountRichiedente=0
      end if 
      if Cdbl(IdAccountFornitore)>0 then 
         MailDocForn = LeggiCampo("select * from AccountProdotto Where IdAccount=" & IdAccountFornitore & " and IdProdotto=" & IdProdotto,"MailDocumentazione")
      end if 
      MailDocumentazione      
      if idT="RICH" then 'richiesta : notifica a tutti i back office 
         notificaAccount1   = IdAccountCliente
         notificaAccount2   = IdAccountRichiedente
         ListaAccountBack   = getAccountsBackOfficeByProd(IdProdotto)
         MailDocumentazione = MailDocForn
      end if 
      if idT="INTE" or idT="ANNU" or idT="ATTI" then ' notifica al cliente
         notificaAccount1 = IdAccountCliente
         notificaAccount2 = IdAccountRichiedente
         if idT<>"INTE" then 
            MailDocumentazione = MailDocForn
         end if 
         'in attivazione mando al cliente le credenziali di accesso
 
      end if
   end if 
   
   if instr("BORS_FIDO_ESTR",idP)>0 then

      qSel = ""
      qSel = qSel & " select * From AccountMovEco where " & IdKey
      'response.write "ecco :" & Qsel 
      'caricamento richiesta : notifico a backoffice tutti se non gia' preso in carico   
      if idT="CARI" then 
         IdAccountGestore = Cdbl("0" & LeggiCampo(qSel,"IdAccountGestore"))
         If Cdbl(IdAccountGestore)=0 then 
            ListaAccountBack = getAccountsBackOffice()
         else
            notificaAccount1 = IdAccountGestore
         end if 
         'response.write "lista: " & ListaAccountBack    
      end if  
      'notifica solo al cliente   
      if instr("ANNU_ACCE_INTE",idT)>0 then 
         notificaAccount1 = Cdbl("0" & LeggiCampo(qSel,"IdAccount"))
      end if      
   end if 
   'pagamento servizio 
   if idP="PAGA" then
      if idT="PFAL" then
         notificaAccount1 = IdAccountEvento
      end if 
   end if 

   if idP="CAUP" then 
      qSel = "select * from Cauzione Where " & IdKey
      IdAccountCliente = cdbl("0" & LeggiCampo(qSel,"IdAccountCliente"))
   
      'polizza annullata : notifica al cliente 
      if idT="ANNU" then
         notificaAccount1 = IdAccountEvento
         notificaAccount2 = IdAccountCliente
      end if   
      'polizza attivata : notifica al cliente 
      if idT="ATTI" then
         notificaAccount1 = IdAccountEvento
         notificaAccount2 = IdAccountCliente
      end if   
      
   end if 
   if idP="ESTR" then 
   end if 
   if idP="FIDO" then 
   end if 
   'se presente query si legge l'account 

   err.clear 
   if MailDocumentazione<>"" then
      xx=createEventoAccount(idAccountFornitore,IdEvento,MailDocumentazione,"","out",0,0,DatiAccessoForm)   
   end if 
   err.clear
   if ListaAccountBack<>"" or cdbl(notificaAccount1)>0 or cdbl(notificaAccount2)>0 or cdbl(notificaRichiedente)>0 then 
      qElenco = ""
      qElenco = qElenco & " Select * from Account"
      qElenco = qElenco & " Where IdAccount in (" & notificaAccount1 & "," & notificaAccount2 & "," & notificaRichiedente & ")" 
      if ListaAccountBack<>"" then 
         qElenco = qElenco & " or IdAccount in (" & ListaAccountBack & ")" 
      end if 
      xx=writeTraceAttivita("ricerca account notifica = " & qElenco,IdAttivita,IdNumAttivita)
      
      myRS.Open qElenco, ConnMsde
      'response.write err.description 
      
      if err.number=0 then 
         Do While Not myRS.EOF 
            'response.write "creo notifica "
            IdAccount      = MyRs("IdAccount")    
            IdTipoAccount  = MyRs("IdTipoAccount")

            InfoAccountRel = ""
            emailNotifica = "CERCA"

            if cdbl(IdAccount) = cdbl(notificaGestore) then 
               emailNotifica   = mail_notificaGestore
               InfoAccountRel  = InfoAccountGestore
            end if 
            if cdbl(IdAccount) = cdbl(notificaRichiedente) then 
               emailNotifica   = mail_notificaRichiedente
               InfoAccountRel  = InfoAccountRichiedente
            end if 
            if cdbl(IdAccount) = cdbl(notificaAccount1) then 
               emailNotifica   = mail_NotificaAccount1
               InfoAccountRel  = InfoAccount1
            end if 
            if cdbl(IdAccount) = cdbl(notificaAccount2) then 
               emailNotifica   = mail_NotificaAccount2
               InfoAccountRel  = InfoAccount2
            end if 
            
            xx=writeTraceAttivita("mailing account = " & IdAccount & " mail=" & emailNotifica ,IdAttivita,IdNumAttivita)
            
            if emailNotifica = "CERCA" then 
               emailNotifica = getMailForAccount(IdAccount,IdTipoAccount,idP,IdProdotto)            
               if emailNotifica = "" and instr(MyRs("UserId"),"@")>0 then 
                  emailNotifica = MyRs("UserId")
               end if    
            end if 

            'response.write "creo notifica :" & emailNotifica & err.description
            telNotifica   = ""
            'response.write "creo notifica "
            ' se gia inviata non la reinvio 
            if ucase(trim(emailNotifica))<>ucase(trim(MailDocumentazione)) then 
               AddInfoAcc = ""
               if cdbl(IdAccount)=cdbl(IdAccountCliente) and idP="FORM" then 
                  AddInfoAcc=DatiAccessoForm
               end if 
               if AddInfoAcc="" then 
                  AddInfoAcc = InfoAccountRel 
               end if 
               xx=createEventoAccount(idAccount,IdEvento,emailNotifica,telNotifica,"out",0,0,AddInfoAcc)
            end if 
            'cerco la mail per TipoAccount
            myRs.moveNext  
         Loop
         myRs.close 
      end if 
   end if   

end function 

function getAccountsBackOffice()
dim q 
   q = ""
   q = q & "Select A.IdAccount "
   q = q & " from Account A, Utente B"
   q = q & " Where A.IdAccount = B.IdAccount"
   q = q & " And A.IdTipoAccount='BackO'"
   q = q & " And A.Abilitato=1" 
   getAccountsBackOffice = q
end function 

function getAccountsBackOfficeByProd(IdProdotto)
dim q 
   q = ""
   q = q & " select distinct A.IdAccount "
   q = q & " from Account A, v_ProdottiAttiviProfiloAccount B "
   q = q & " Where A.IdAccount=B.IdAccount "
   q = q & " and   B.IdProdotto = " &IdProdotto
   q = q & " and   A.IdTipoAccount='BackO'"
   q = q & " And   A.Abilitato=1"
   getAccountsBackOfficeByProd = q   
end function 
  
function getMailForAccount(IdAccount,IdTipoAccount,Processo,IdProdotto)
dim q,mail,qSelContatto
   mail=""
   q = ""
   q = q & " select DescContatto as mail "
   q = q & " from AccountContatto "
   q = q & " Where IdAccount=" & IdAccount
   q = q & " and IdTipoContatto='MAIL'"
   q = q & " and DescContatto like '%@%'"
   q = q & " order by FlagPrincipale desc"
   qSelContatto = q
   'response.write q
   q = ""
   if ucase(IdTipoAccount)=ucase("Admin") then 
   end if 
   if ucase(IdTipoAccount)=ucase("BackO") then 
      q = " select EMail as mail from Utente where IdAccount=" & IdAccount
   end if 
   if ucase(IdTipoAccount)=ucase("Clie") then 
      q = qSelContatto
   end if 
   if ucase(IdTipoAccount)=ucase("Coll") then 
      q = qSelContatto
   end if 
   if ucase(IdTipoAccount)=ucase("Forn") then 
      if Processo="AFFI" then 
         q = ""
         q = q & " select MailDocumentazione as Mail "
         q = q & " from AccountProdotto "
         q = q & " Where IdAccount = " & IdAccount
         q = q & " and IdProdotto = " & IdProdotto
         mail = LeggiCampo(q,"mail")
         if mail="" then 
            q = qSelContatto
         else
            q = ""
         end if 
      else
         q = qSelContatto
      end if 
   end if    
   if q<>"" then 
      mail = LeggiCampo(q,"mail")
   end if
   'se non l'ho trovata provo a vedere se l'account l'ha sul prodotto
   if mail="" then 
      q = ""
      q = q & " select MailDocumentazione as Mail "
      q = q & " from AccountProdotto "
      q = q & " Where IdAccount = " & IdAccount
      q = q & " and IdProdotto = " & IdProdotto
      mail = LeggiCampo(q,"mail")
   end if 
   getMailForAccount=mail 
end function 


'elaboro tutti gli invii da fare  
function elaboraEventoMailing(IdEvento)
dim qSel,myRS,qUpd,IdAccount,toAddress,addInfo,xx, Nominativo  
   err.clear 
   on error resume next 
   qSel = ""
   qSel = qSel & " select * from EventoAccount a , Evento b"
   qSel = qSel & " Where A.IdEvento = " & IdEvento
   qSel = qSel & " and A.IdEvento = B.IdEvento "   
   qSel = qSel & " and A.DataNotifica=0"

   'response.write qSel 
   
   Set myRS = Server.CreateObject("ADODB.Recordset")
   myRS.CursorLocation = 3 
   myRS.Open qSel, ConnMsde
   if err.number=0 then 
      Do While Not myRS.EOF 
         toAddress  = trim(myRS("emailNotifica"))
         addInfo    = trim(myRS("DescEvento"))
         AddInfoAcc = trim(myRS("DescEventoAccount"))
         if AddInfoAcc<>"" then 
            addInfo = addInfo & "<br>" & AddInfoAcc
         end if 
 
         if toAddress<>"" and addInfo<>"" then 
            Nominativo = LeggiCampo("Select * from Account Where idAccount=" & myRS("IdAccount"),"Nominativo")
            TextHtml=""
            TextText=""
            xx=CreaTestoMail(Nominativo,addInfo,TextHtml,TextText)
            'response.write "sending : " & Nominativo & " " & ToAddress
            xx=SendMailMessageHTMLWithAttach("", ToAddress, "","notifiche ", TextText, TextHtml, "", false)    
            'response.write xx & " " & err.description 
         end if 
         myRS.MoveNext
      Loop
      myRS.close 
      'si chiudono gli invii
      qUpd = ""  
      qUpd = qUpd & " update EventoAccount set "
      qUpd = qUpd & " dataNotifica = " & DtoS() 
      qUpd = qUpd & ",timeNotifica = " & TimeToS() 
      qUpd = qUpd & " Where IdEvento = " & IdEvento
      qUpd = qUpd & " and dataNotifica = 0"
      ConnMsde.execute qUpd  
   end if 
   err.clear 
   
end function 
%>




