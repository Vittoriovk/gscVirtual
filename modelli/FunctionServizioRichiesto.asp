<%

'contiene i nomi dei campi della tabella ServizioRichiesto
'non devono essere presenti campi con lo stesso nome sulle due tabelle 
Dim DizCampiServizioRichiesto
Const tipoImporto_Premio                     = "PREMIO"
Const tipoImporto_PremioAssicurativo         = "PREMIO_ASS"
Const tipoImporto_DirittiEmissione           = "DIR_EMI"
Const tipoImporto_DirittiIntermediazione     = "DIR_INT"
Const tipoImporto_DirittiIntermediazioneColl = "DIR_INT_COLL"
Const tipoImporto_Spese                      = "SPESE"
Const tipoImporto_AltroImporto               = "ALTRO"
Const tipoImporto_Provvigioni                = "PROV"


Function GetDizServizioRichiesto(IdServizioRichiesto,IdAttivita,IdNumAttivita)

   Dim MyRs,MySql,IdStato,K,nome
   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
   
   xx=SetDiz(DizDatabase,"N_IdServizioRichiesto",0)
   xx=SetDiz(DizDatabase,"S_IdAttivita",IdAttivita)
   xx=SetDiz(DizDatabase,"N_IdNumAttivita",IdNumAttivita)
   xx=SetDiz(DizDatabase,"S_DescServizioRichiesto","")
   xx=SetDiz(DizDatabase,"N_IdAccountRichiedente",0)
   xx=SetDiz(DizDatabase,"N_IdAccountCliente",0)
   xx=SetDiz(DizDatabase,"N_IdAccountGestore",0)
   xx=SetDiz(DizDatabase,"S_IdStatoServizio","")
   xx=SetDiz(DizDatabase,"S_IdStatoServizioPrec","")
   xx=SetDiz(DizDatabase,"S_DescStatoServizio","")
   xx=SetDiz(DizDatabase,"N_FlagStatoFinale",0)
   xx=SetDiz(DizDatabase,"S_NoteStatoServizio","")
   xx=SetDiz(DizDatabase,"N_IdProdotto",0)  
   xx=SetDiz(DizDatabase,"N_IdProdottoTemplate",0)     
   xx=SetDiz(DizDatabase,"S_IdAnagServizio","")
   xx=SetDiz(DizDatabase,"N_IdAnagCaratteristica",0)
   xx=SetDiz(DizDatabase,"S_DescAnagCaratteristica","")
   xx=SetDiz(DizDatabase,"N_IdCompagnia",0)
   xx=SetDiz(DizDatabase,"N_IdFornitore",0)
   xx=SetDiz(DizDatabase,"N_IdAccountFornitore",0)
   xx=SetDiz(DizDatabase,"N_DataRegistrazione",0)
   xx=SetDiz(DizDatabase,"N_TimeRegistrazione",0)
   xx=SetDiz(DizDatabase,"N_DataChiusura",0)
   xx=SetDiz(DizDatabase,"N_TimeChiusura",0)
   xx=SetDiz(DizDatabase,"N_ValidoDal",0)
   xx=SetDiz(DizDatabase,"N_ValidoAl",0)
   xx=SetDiz(DizDatabase,"N_GiorniValidita",0)
   xx=SetDiz(DizDatabase,"N_ImptTotaleServizio",0)
   xx=SetDiz(DizDatabase,"N_ImptServizio",0)
   xx=SetDiz(DizDatabase,"N_ImptDirittiEmissione",0)
   xx=SetDiz(DizDatabase,"N_ImptIntermediazione",0)
   xx=SetDiz(DizDatabase,"N_ImptIntermediazioneColl",0)
   xx=SetDiz(DizDatabase,"N_ImptAltro",0)
   xx=SetDiz(DizDatabase,"N_ImptProvvigioni",0)
   xx=SetDiz(DizDatabase,"N_ImptProvvigioniRete",0)
   xx=SetDiz(DizDatabase,"S_NoteServizio","")
   xx=SetDiz(DizDatabase,"S_NoteServizioCliente","")   
   xx=SetDiz(DizDatabase,"S_IdTipoCreditoClie","")
   xx=SetDiz(DizDatabase,"S_IdTipoCreditoRequ","")
   xx=SetDiz(DizDatabase,"S_IdProcessoElaborativo","")
   xx=SetDiz(DizDatabase,"N_FlagContabilizzato",0)   
   xx=SetDiz(DizDatabase,"N_FlagCalcoloImporti",0)   
   xx=SetDiz(DizDatabase,"N_IdAccountLivello1",0)   
   xx=SetDiz(DizDatabase,"N_IdAccountLivello2",0)   
   xx=SetDiz(DizDatabase,"N_IdAccountLivello3",0)      
   xx=SetDiz(DizDatabase,"N_DataMovimento",0)      
   xx=SetDiz(DizDatabase,"N_TimeMovimento",0)      
   xx=SetDiz(DizDatabase,"N_ImptTasse",0)         
   xx=SetDiz(DizDatabase,"S_mailNotificaCliente","")
   xx=SetDiz(DizDatabase,"S_mailNotificaRichiedente","")
   xx=SetDiz(DizDatabase,"S_mailNotificaBackOffice","")
   xx=SetDiz(DizDatabase,"S_NoteServizioFornitore","")
   xx=SetDiz(DizDatabase,"N_IdRamo",0)
   xx=SetDiz(DizDatabase,"N_IdDirittiEmissione",0)
   xx=SetDiz(DizDatabase,"N_ImptDistribuzione",0)
   xx=SetDiz(DizDatabase,"N_ImptListino",0)
   xx=SetDiz(DizDatabase,"N_ImptFornitore",0)
   xx=SetDiz(DizDatabase,"S_IdFlussoProcesso","")
   xx=SetDiz(DizDatabase,"N_DataAnnullabile",0)
   xx=SetDiz(DizDatabase,"S_IdStatoServizioBackO","")
   xx=SetDiz(DizDatabase,"N_DataInvioFornitore",0)
   xx=SetDiz(DizDatabase,"N_IdMovimentoAcquisto",0)
   xx=SetDiz(DizDatabase,"N_IdMovimentoStorno",0)
   
   xx=SetDiz(DizDatabase,"N_IdAccPaga1",0)
   xx=SetDiz(DizDatabase,"S_IdTipoCreditoAcc1","")
   xx=SetDiz(DizDatabase,"S_IdStatoCreditoAcc1","")
   xx=SetDiz(DizDatabase,"N_IdAccPaga2",0)
   xx=SetDiz(DizDatabase,"S_IdTipoCreditoAcc2","")
   xx=SetDiz(DizDatabase,"S_IdStatoCreditoAcc2","")
   xx=SetDiz(DizDatabase,"N_IdAccPaga3",0)
   xx=SetDiz(DizDatabase,"S_IdTipoCreditoAcc3","")
   xx=SetDiz(DizDatabase,"S_IdStatoCreditoAcc3","")

   'copio i nomi dei campi nel dizionario dei campi 
   Set DizCampiServizioRichiesto = CreateObject("Scripting.Dictionary") 
   For Each K In DizDatabase
       xx=SetDiz(DizCampiServizioRichiesto,k,k) 
   next    

   Set MyRs = Server.CreateObject("ADODB.Recordset")
   if Cdbl(IdServizioRichiesto)>0 or (IdAttivita<>"" and Cdbl(IdNumAttivita)>0) then 
      MySql = ""
      MySql = MySql & " select * from ServizioRichiesto "
      if Cdbl(IdServizioRichiesto)>0 then 
         MySql = MySql & " Where IdServizioRichiesto=" & IdServizioRichiesto  
      else
         MySql = MySql & " Where IdAttivita='" & IdAttivita & "'"
         MySql = MySql & " and   IdNumAttivita=" & NumforDb(IdNumAttivita)
      end if 

      MyRs.CursorLocation = 3 
	  
      MyRs.Open MySql, ConnMsde      
      if MyRs.eof = false then 
         For Each K In DizDatabase
            if isOfServizioRichiesto(k) then 
               nome = mid(k,3,99)
               xx=SetDiz(DizDatabase,k,MyRs(nome)) 
               'response.write k & " " & MyRs(nome) & "<br>"
            end if 
         next 
      end if 
      MyRs.close 
   end if 
   set GetDizServizioRichiesto = DizDatabase
End function 

Function isOfServizioRichiesto(key)
   isOfServizioRichiesto = DizCampiServizioRichiesto.Exists(key)
End function 

Function GetNewServizioRichiesto(DizDatabase)
Dim MySql,v_id,Campi,Valori,IdAttivita,IdNumAttivita 
   on error resume next 
   v_id=0
   MySql = ""
   Campi = ""
   Valori= ""
   'controllo valori predefiniti 
   'leggo tutti i parametri ad accezione di IdServizioRichiesto
   For Each K In DizDatabase
      if isOfServizioRichiesto(k) then 
         Valo = GetDiz(DizDatabase,K)
         Tipo = mid(k,1,2)
         nome = mid(k,3,99)
         'response.write nome 
		 if nome=ucase("IdAttivita") then 
		    IdAttivita = valo
		 end if 
		 if nome=ucase("IdNumAttivita") then 
		    IdNumAttivita = valo
		 end if 
		 
         if nome<>ucase("IdServizioRichiesto") then 
            if Campi <> "" then 
               Campi  = Campi & ","
               Valori = Valori & ","
            end if 
            Campi  = Campi & nome 

           if Tipo="N_" then 
              Valori = Valori & NumForDb(valo)
           else
              Valori = Valori & "'" & apici(valo) & "'"
           end if 
        end if         
     end if 
   Next
   MySql = MySql & " Insert into ServizioRichiesto (" & Campi &  ") Values (" & Valori & ")"
   
   xx=writeTraceAttivita("GetNewServizioRichiesto:" & MySql ,IdAttivita,NumForDb(IdNumAttivita)) 
   
   ConnMsde.execute MySql 
   if err.number=0 then 
      v_id = GetTableIdentity("ServizioRichiesto")  
      xx   = SetDiz(DizDatabase,"N_IdServizioRichiesto",v_id)
   else 
      xx=writeTrace("GetNewServizioRichiesto:" & MySql & ":" & err.description)
   end if 
   GetNewServizioRichiesto=v_id
end function 


Function UpdateServizioRichiesto(DizDatabase)
Dim v_id , MySql , Esito , v_ret,Nome,IdAz,IdServizioRichiesto,IdAttivita,IdNumAttivita
Dim tmpId,tmpDe,IdProdotto,IdAccountLivello1,IdProdottoTemplate
on error resume next 
   Esito = ""
  MySql = ""
   
   'recupero dati di descrizione 
   tmpId=GetDiz(DizDatabase,"S_IdStatoServizio")
   'response.write TmpId
   tmpDe=LeggiCampo("select * from StatoServizio Where IdStatoServizio='" & tmpId & "'","DescStatoServizio")
   'response.write TmpDe
   xx=SetDiz(DizDatabase,"S_DescStatoServizio",tmpDe)
   tmpDe=LeggiCampo("select * from StatoServizio Where IdStatoServizio='" & tmpId & "'","FlagStatoFinale")
   xx=SetDiz(DizDatabase,"N_FlagStatoFinale",tmpDe) 

   idProdotto         = cdbl("0" & GetDiz(DizDatabase,"N_IdProdotto"))
   IdProdottoTemplate = cdbl("0" & GetDiz(DizDatabase,"N_IdProdottoTemplate"))
   if Cdbl(IdProdottoTemplate)=0 then 
      IdProdottoTemplate = LeggiCampo("select * from Prodotto Where IdProdotto=" & IdProdotto,"IdProdottoTemplate")
	  IdProdottoTemplate = Cdbl("0" & IdProdottoTemplate)
      xx=SetDiz(DizDatabase,"N_IdProdottoTemplate",IdProdottoTemplate)   
   end if 
   tmpId=GetDiz(DizDatabase,"S_IdAnagServizio")
   if tmpId="" then 
      tmpId = LeggiCampo("select * from ProdottoTemplate Where IdProdottoTemplate=" & IdProdottoTemplate,"IdAnagServizio")
      xx=SetDiz(DizDatabase,"S_IdAnagServizio",tmpId) 
   end if 
   tmpId=cdbl("0" & GetDiz(DizDatabase,"N_IdAnagCaratteristica"))
   if tmpId=0 then 
      tmpId = LeggiCampo("select * from ProdottoTemplate Where IdProdottoTemplate=" & IdProdottoTemplate,"IdAnagCaratteristica")
      if Cdbl(tmpId)>0 then 
         xx=SetDiz(DizDatabase,"N_IdAnagCaratteristica",tmpId) 
         tmpDe=LeggiCampo("select * from AnagCaratteristica Where IdAnagCaratteristica=" & tmpId,"DescAnagCaratteristica")
         xx=SetDiz(DizDatabase,"S_DescAnagCaratteristica",tmpDe) 

      end if 
   end if 
   tmpId=cdbl("0" & GetDiz(DizDatabase,"N_IdRamo"))
   if tmpId=0 then 
      tmpId = LeggiCampo("select * from ProdottoTemplate Where IdProdottoTemplate=" & IdProdottoTemplate,"IdRamo")
      if Cdbl(tmpId)>0 then 
         xx=SetDiz(DizDatabase,"N_IdRamo",tmpId) 
      end if 
   end if   
   tmpId=cdbl("0" & GetDiz(DizDatabase,"N_IdFornitore"))
   tmpDe=cdbl("0" & GetDiz(DizDatabase,"N_IdAccountFornitore"))
   if tmpId=0 and tmpDe>0 then 
      tmpId = LeggiCampo("select * from Fornitore Where IdAccount=" & tmpDe,"IdFornitore")
      if Cdbl(tmpId)>0 then 
         xx=SetDiz(DizDatabase,"N_IdFornitore",tmpId) 
      end if 
   end if    
   if tmpId>0 and tmpDe=0 then 
      tmpDe = LeggiCampo("select * from Fornitore Where IdFornitore=" & tmpDe,"IdAccount")
      if Cdbl(tmpDe)>0 then 
         xx=SetDiz(DizDatabase,"N_IdAccountFornitore",tmpDe) 
      end if 
   end if    
   'leggo tutti i parametri ad accezione di IdServizioRichiesto
   IdServizioRichiesto=0
   IdAttivita=""
   IdNumAttivita=0
   IdAccountLivello1=0
   For Each K In DizDatabase
      if isOfServizioRichiesto(k) then
        'response.write 
        Valo = DizDatabase.item(ucase(K))
        Tipo = mid(k,1,2)
        nome = mid(k,3,99)
        'response.write nome & "=" & valo & " " & K
        if nome<>ucase("IdServizioRichiesto") and nome<>ucase("IdAttivita") and nome<>ucase("IdNumAttivita")  then 
           if MySql <> "" then 
              MySql = MySql & ","
           end if 
           MySql = MySql & nome & "="
           if Tipo="N_" then 
              MySql = MySql & NumForDb(valo)
           else
              MySql = MySql & "'" & apici(valo) & "'"
           end if            
        else
           if nome=ucase("IdServizioRichiesto")  then 
              IdServizioRichiesto=Cdbl(Valo)
           end if 
           if nome=ucase("IdAttivita") then 
              IdAttivita=Valo
           end if 
           if nome=ucase("IdNumAttivita")  then
              IdNumAttivita=Cdbl(Valo)
           end if 
           if nome=ucase("IdAccountLivello1")  then
              IdAccountLivello1=Cdbl(Valo)
           end if 
        end if 
     end if   
   Next
   MySql = " Update ServizioRichiesto Set " & MySql 
   if cdbl(IdServizioRichiesto)>0 then 
      MySql = MySql & " Where IdServizioRichiesto = " & IdServizioRichiesto
   else
      MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita
   end if 
   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateServizioRichiesto:" & MySql & ":" & err.description)
   else 
      xx=aggiornaTotaleServizioRichiesto(IdAttivita,IdNumAttivita,IdServizioRichiesto)
      if cdbl(IdAccountLivello1)=0 then 
         xx=verifyAccountServizioRichiesto(IdAttivita,IdNumAttivita)
	  end if 
   end if    
  
   UpdateServizioRichiesto=Esito 
End function 

Function getIdServizioRichiestoByAttivita(IdAttivita,IdNumAttivita)
Dim MySql,retV
   MySql = ""
   MySql = MySql & " select * from ServizioRichiesto  " 
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita

   retV = cdbl("0" & LeggiCampo(MySql,"IdServizioRichiesto"))
   getIdServizioRichiestoByAttivita = retV 
End function 

Function getStatoServizioRichiesto(IdServizioRichiesto,IdAttivita,IdNumAttivita)
Dim MySql,retV
   MySql = ""
   MySql = MySql & " select * from ServizioRichiesto "
   if Cdbl(IdServizioRichiesto)>0 then 
      MySql = MySql & " Where IdServizioRichiesto=" & IdServizioRichiesto  
   else
      MySql = MySql & " Where IdAttivita='" & IdAttivita & "'"
      MySql = MySql & " and   IdNumAttivita=" & NumforDb(IdNumAttivita)
   end if 

   retV = LeggiCampo(MySql,"IdStatoServizio")
   getStatoServizioRichiesto = retV 
End function 

Function UpdateStatoPrecServizioRichiesto(IdAttivita,IdNumAttivita,IdStatoServizio)
Dim MySql
   on error resume next 
 
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " IdStatoServizioPrec = '" & apici(IdStatoServizio) & "'"
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateStatoPrecServizioRichiesto:" & MySql & ":" & err.description)
   end if 

End function 

Function UpdateFlussoProcessoServizioRichiesto(IdServizioRichiesto,IdAttivita,IdNumAttivita,IdFlussoProcesso)
Dim MySql
   on error resume next 
 
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " IdFlussoProcesso = '" & apici(IdFlussoProcesso) & "'"
   if Cdbl(IdServizioRichiesto)>0 then 
      MySql = MySql & " Where IdServizioRichiesto= " & NumForDb(IdServizioRichiesto)
   else 
      MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita
   end if 

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateStatoPrecServizioRichiesto:" & MySql & ":" & err.description)
   end if 

End function 

Function UpdateAccoutPagatoreServizioRichiesto(IdServizioRichiesto,IdAttivita,IdNumAttivita,IdAccount,IdTipoCredito,IdStatoCredito,Progressivo)
Dim MySql
   on error resume next 
 
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   if cdbl(Progressivo)=0 then 
      MySql = MySql & " IdAccPaga1=0,IdAccPaga2=0,IdAccPaga3=0 " 
	  MySql = MySql & ",IdTipoCreditoAcc1='',IdTipoCreditoAcc2='',IdTipoCreditoAcc3='' " 
	  MySql = MySql & ",IdStatoCreditoAcc1='',IdStatoCreditoAcc2='',IdStatoCreditoAcc3='' " 
   else
      if cdbl(Progressivo)=1 then 
         MySql = MySql & " IdAccPaga1 = " & IdAccount
         MySql = MySql & ",IdTipoCreditoAcc1 = '" & apici(IdTipoCredito) & "'"
		 MySql = MySql & ",IdStatoCreditoAcc1 = '" & apici(IdStatoCredito) & "'"
      end if 
      if cdbl(Progressivo)=2 then 
         MySql = MySql & " IdAccPaga2 = " & IdAccount
         MySql = MySql & ",IdTipoCreditoAcc2 = '" & apici(IdTipoCredito) & "'"
		 MySql = MySql & ",IdStatoCreditoAcc2 = '" & apici(IdStatoCredito) & "'"
      end if 
      if cdbl(Progressivo)=3 then 
         MySql = MySql & " IdAccPaga3 = " & IdAccount
         MySql = MySql & ",IdTipoCreditoAcc3 = '" & apici(IdTipoCredito) & "'"
		 MySql = MySql & ",IdStatoCreditoAcc3 = '" & apici(IdStatoCredito) & "'"
      end if 
   end if 
   
   if Cdbl(IdServizioRichiesto)>0 then 
      MySql = MySql & " Where IdServizioRichiesto= " & NumForDb(IdServizioRichiesto)
   else 
      MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita
   end if 

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateAccoutPagatoreServizioRichiesto:" & MySql & ":" & err.description)
   end if 

End function 

Function UpdateStatoServizioRichiestoById(IdServizioRichiesto,IdStatoServizio,NoteStatoServizio)
Dim IdAttivita,IdNumAttivita, retVal 
 
   IdAttivita    = LeggiCampoServizioRichiestoById(IdServizioRichiesto,"IdAttivita")
   IdNumAttivita = LeggiCampoServizioRichiestoById(IdServizioRichiesto,"IdNumAttivita")
   retVal = UpdateStatoServizioRichiesto(IdAttivita,IdNumAttivita,IdStatoServizio,NoteStatoServizio)
   UpdateStatoServizioRichiestoById = retVal 
   
End function 

Function UpdateStatoServizioRichiesto(IdAttivita,IdNumAttivita,IdStatoServizio,NoteStatoServizio)
Dim MySql,tmpDe,tmpFl,xx,IdStatoServizioPrec 
   on error resume next 
   tmpDe=LeggiCampo("select * from StatoServizio Where IdStatoServizio='" & IdStatoServizio & "'","DescStatoServizio")
   tmpFl=LeggiCampo("select * from StatoServizio Where IdStatoServizio='" & IdStatoServizio & "'","FlagStatoFinale")
   
   MySql = ""
   MySql = MySql & " select IdStatoServizio from ServizioRichiesto " 
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita
   IdStatoServizioPrec  = leggiCampo(MySql,"IdStatoServizio")
   
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " IdStatoServizio = '" & apici(IdStatoServizio) & "'"
   MySql = MySql & ",DescStatoServizio = '" & apici(tmpDe) & "'"
   MySql = MySql & ",FlagStatoFinale = " & numForDb(tmpFl)
   MySql = MySql & ",NoteStatoServizio = '" & apici(NoteStatoServizio) & "'"
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateStatoServizioRichiesto:" & MySql & ":" & err.description)
   elseif IdStatoServizio<>IdStatoServizioPrec then 
      xx=notificaEventoServizioRichiesto(0,IdAttivita,IdNumAttivita)   
   end if 

End function 

Function setProcessoElabServizioRichiesto(IdAttivita,IdNumAttivita)
Dim MySql,IdProcessoElaborativo 
    on error resume next 
    MySql = ""
    MySql = MySql & " select c.IdProcessoElaborativo "
    MySql = MySql & " from ServizioRichiesto A,Cliente B, Collaboratore C "
    MySql = MySql & " Where A.IdAttivita='" & IdAttivita & "' and A.IdNumAttivita = " & IdNumAttivita
    MySql = MySql & " and   A.IdAccountCliente = B.IdAccount"
    MySql = MySql & " and   B.IdAccountLivello1 = C.IdAccount"
	'response.write MySql 
    IdProcessoElaborativo = LeggiCampo(MySql,"IdProcessoElaborativo")
    
    MySql = ""
    MySql = MySql & " Update ServizioRichiesto Set " 
    MySql = MySql & " IdProcessoElaborativo = '" & apici(IdProcessoElaborativo) & "'"
    MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita
	'response.write MySql
    ConnMsde.Execute MySql 
    if err.number<>0 then 
       writeTrace("setProcessoElabServizioRichiesto:" & MySql & ":" & err.description)
    end if 
End function 


Function UpdateNoteFornitoreServizioRichiesto(IdAttivita,IdNumAttivita,NoteServizioFornitore)
Dim MySql 
   on error resume next 
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " NoteServizioFornitore = '" & apici(NoteServizioFornitore) & "'"
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateNoteFornitoreServizioRichiesto:" & MySql & ":" & err.description)
   end if 
   
End function 

Function UpdateNoteServizioRichiesto(IdAttivita,IdNumAttivita,NoteServizio,NoteServizioCliente)
Dim MySql 
   on error resume next 
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " NoteServizio = '" & apici(NoteServizio) & "'"
   MySql = MySql & ",NoteServizioCliente = '" & apici(NoteServizioCliente) & "'"
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateNoteServizioRichiesto:" & MySql & ":" & err.description)
   end if 
   
End function 

Function UpdateNoteFornServizioRichiesto(IdAttivita,IdNumAttivita,DataInvioFornitore,IdStatoServizioBackO,NoteServizioFornitore)
Dim MySql 
   on error resume next 
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " DataInvioFornitore = " & NumForDb(DataInvioFornitore)
   MySql = MySql & ",IdStatoServizioBackO = '" & apici(IdStatoServizioBackO) & "'"
   MySql = MySql & ",NoteServizioFornitore = '" & apici(NoteServizioFornitore) & "'"   
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita
   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateNoteFornServizioRichiesto:" & MySql & ":" & err.description)
   end if 
   
End function 

Function UpdateGestoreServizioRichiesto(IdAttivita,IdNumAttivita,IdAccountGestore)
Dim MySql,mail 
   
   on error resume next 
   mail = LeggiCampo("select * from Utente Where IdAccount=" & IdAccountGestore,"email")
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " IdAccountGestore = " & NumforDb(IdAccountGestore)
   MySql = MySql & ",mailNotificaBackOffice='" & apici(mail) & "'"
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita
   MySql = MySql & " and IdAccountGestore=0"

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateGestoreServizioRichiesto:" & MySql & ":" & err.description)
   end if 
   
End function

Function UpdateCompagniaServizioRichiesto(IdAttivita,IdNumAttivita,IdCompagnia)
Dim MySql 
   on error resume next 
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " IdCompagnia = " & NumforDb(IdCompagnia)
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateCompagniaServizioRichiesto:" & MySql & ":" & err.description)
   end if 
   
End function

Function UpdateFornitoreServizioRichiesto(IdAttivita,IdNumAttivita,IdFornitore,IdAccountFornitore)
Dim MySql 
   on error resume next 
   if cdbl(IdAccountFornitore)=0 and Cdbl(IdFornitore)>0 then 
      IdAccountFornitore = cdbl("0" & LeggiCampo("select * from Fornitore where IdFornitore=" & IdFornitore,"IdAccount"))
   end if 
   if cdbl(IdAccountFornitore)>0 and Cdbl(IdFornitore)=0 then 
      IdFornitore = cdbl("0" & LeggiCampo("select * from Fornitore where IdAccount=" & IdAccountFornitore,"IdFornitore"))
   end if    
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " IdFornitore = " & NumforDb(IdFornitore)
   MySql = MySql & ",IdAccountFornitore = " & NumforDb(IdAccountFornitore)
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("UpdateFornitoreServizioRichiesto:" & MySql & ":" & err.description)
   end if 
   
End function

Function LeggiCampoServizioRichiesto(IdAttivita,IdNumAttivita,campo)
dim retV 
   retV = ""
   retV = LeggiCampo("select * from ServizioRichiesto Where IdAttivita='" & apici(IdAttivita) & "' and IdNumAttivita=" & NumforDb(IdNumAttivita),campo)
   LeggiCampoServizioRichiesto = retV
end function 

function LeggiCampoServizioRichiestoById(IdServizioRichiesto,campo)
dim retV 
   retV = ""
   retV = LeggiCampo("select * from ServizioRichiesto Where IdServizioRichiesto=" & NumforDb(IdServizioRichiesto),campo)
   LeggiCampoServizioRichiestoById = retV
end function 



function UpdateImportiServizioRichiesto(IdAttivita,IdNumAttivita,ImptServizio,ImptProvvigioni,ImptDirittiEmissione,ImptIntermediazione,ImptIntermediazioneColl,ImptAltro)
Dim ImptTotaleServizio
   on error resume next 
   ImptTotaleServizio = 0
   ImptTotaleServizio = cdbl(ImptTotaleServizio) + cdbl(ImptServizio)
   ImptTotaleServizio = cdbl(ImptTotaleServizio) + cdbl(ImptDirittiEmissione)
   ImptTotaleServizio = cdbl(ImptTotaleServizio) + cdbl(ImptIntermediazione)
   ImptTotaleServizio = cdbl(ImptTotaleServizio) + cdbl(ImptIntermediazioneColl)
   ImptTotaleServizio = cdbl(ImptTotaleServizio) + cdbl(ImptAltro)
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " ImptTotaleServizio = " & NumforDb(ImptTotaleServizio)
   MySql = MySql & ",ImptServizio = " & NumforDb(ImptServizio)
   MySql = MySql & ",ImptProvvigioni = " & NumforDb(ImptProvvigioni)
   MySql = MySql & ",ImptDirittiEmissione = " & NumforDb(ImptDirittiEmissione)
   MySql = MySql & ",ImptIntermediazione = " & NumforDb(ImptIntermediazione)
   MySql = MySql & ",ImptIntermediazioneColl = " & NumforDb(ImptIntermediazioneColl)
   MySql = MySql & ",ImptAltro = " & NumforDb(ImptAltro)
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      xx=writeTrace("UpdateImportiServizioRichiesto:" & MySql & ":" & err.description)
   end if 

end function 

function UpdateInterCollServizioRichiesto(IdAttivita,IdNumAttivita,ImptIntermediazioneColl)
Dim IdServizioRichiesto
   on error resume next 
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " ImptIntermediazioneColl = " & NumforDb(ImptIntermediazioneColl)
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita
   ConnMsde.Execute MySql

   IdServizioRichiesto = getIdServizioRichiestoByAttivita(IdAttivita,IdNumAttivita)
   xx = aggiornaTotaleServizioRichiesto(IdAttivita,IdNumAttivita,IdServizioRichiesto)
   
   if err.number<>0 then 
      xx=writeTrace("UpdateImportiServizioRichiesto:" & MySql & ":" & err.description)
   end if 

end function 

function UpdateInterServizioRichiesto(IdAttivita,IdNumAttivita,ImptIntermediazione)
Dim IdServizioRichiesto
   on error resume next 
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " ImptIntermediazione = " & NumforDb(ImptIntermediazione)
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita
   ConnMsde.Execute MySql

   IdServizioRichiesto = getIdServizioRichiestoByAttivita(IdAttivita,IdNumAttivita)
   xx = aggiornaTotaleServizioRichiesto(IdAttivita,IdNumAttivita,IdServizioRichiesto)
   
   if err.number<>0 then 
      xx=writeTrace("UpdateInterServizioRichiesto:" & MySql & ":" & err.description)
   end if 

end function 

function UpdateEmisServizioRichiesto(IdAttivita,IdNumAttivita,ImptEmissione)
Dim IdServizioRichiesto
   on error resume next 
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " ImptDirittiEmissione = " & NumforDb(ImptEmissione)
   MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita
   ConnMsde.Execute MySql

   IdServizioRichiesto = getIdServizioRichiestoByAttivita(IdAttivita,IdNumAttivita)
   xx = aggiornaTotaleServizioRichiesto(IdAttivita,IdNumAttivita,IdServizioRichiesto)
   
   if err.number<>0 then 
      xx=writeTrace("UpdateEmisServizioRichiesto:" & MySql & ":" & err.description)
   end if 

end function 

Function aggiornaTotaleServizioRichiesto(IdAttivita,IdNumAttivita,IdServizioRichiesto)

   if cdbl(IdServizioRichiesto)=0 then 
      IdServizioRichiesto = getIdServizioRichiestoByAttivita(IdAttivita,IdNumAttivita)
   end if 
   
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " ImptTotaleServizio = ImptServizio+ImptDirittiEmissione+ImptIntermediazione+ImptIntermediazioneColl+ImptAltro"
   MySql = MySql & " Where IdServizioRichiesto= " & IdServizioRichiesto

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      xx=writeTrace("aggiornaTotaleServizioRichiesto:" & MySql & ":" & err.description)
   end if 
end function 

Function aggiornaAnnullabileServizioRichiesto(IdServizioRichiesto,IdAttivita,IdNumAttivita,dataAnnullabile)

   if cdbl(IdServizioRichiesto)=0 then 
      IdServizioRichiesto = getIdServizioRichiestoByAttivita(IdAttivita,IdNumAttivita)
   end if 
   
   MySql = ""
   MySql = MySql & " Update ServizioRichiesto Set " 
   MySql = MySql & " dataAnnullabile = " & dataAnnullabile
   MySql = MySql & " Where IdServizioRichiesto = " & IdServizioRichiesto

   ConnMsde.Execute MySql 
   if err.number<>0 then 
      xx=writeTrace("aggiornaAnnullabileServizioRichiesto:" & MySql & ":" & err.description)
   end if 
end function 

'valorizza gli account del cliente  
Function verifyAccountServizioRichiesto(IdAttivita,IdNumAttivita)
Dim MySql,IdAccountLivello1,IdAccountLivello2,IdAccountLivello3,IdAccountCliente
Dim ClSql 
   IdAccountCliente =0
   IdAccountLivello1=0
   IdAccountLivello2=0
   IdAccountLivello3=0
   IdAccountCliente  = cdbl("0" & LeggiCampoServizioRichiesto(IdAttivita,IdNumAttivita,"IdAccountCliente"))
   IdAccountLivello1 = cdbl("0" & LeggiCampoServizioRichiesto(IdAttivita,IdNumAttivita,"IdAccountLivello1"))
   if Cdbl(IdAccountCliente)>0 and cdbl(IdAccountLivello1)=0 then 
      ClSql = "select * from Cliente Where IdAccount=" & IdAccountCliente
      IdAccountLivello1=cdbl("0" & LeggiCampo(ClSql ,"IdAccountLivello1"))
      IdAccountLivello2=cdbl("0" & LeggiCampo(ClSql ,"IdAccountLivello2"))
      IdAccountLivello3=cdbl("0" & LeggiCampo(ClSql ,"IdAccountLivello3"))
      
      MySql = ""
      MySql = MySql & " Update ServizioRichiesto Set " 
      MySql = MySql & " IdAccountLivello1 = " & NumforDb(IdAccountLivello1)
      MySql = MySql & ",IdAccountLivello2 = " & NumforDb(IdAccountLivello2)
      MySql = MySql & ",IdAccountLivello3 = " & NumforDb(IdAccountLivello3)
      MySql = MySql & " Where IdAttivita='" & IdAttivita & "' and IdNumAttivita = " & IdNumAttivita  
      ConnMsde.Execute MySql 
      xx=writeTraceAttivita("aggiorno livelli " & MySql & Err.description ,IdAttivita,IdNumAttivita)
   end if 
   
end function 

Function CancellaServizioRichiesto(idC,idAccountCliente)
Dim MsgErrore,qDel,recordsAffected,xx
MsgErrore=""
err.clear
qDel = ""
qDel = qDel & " update ServizioRichiesto "
qDel = qDel & " set    IdStatoServizio='CANC'"
qDel = qDel & " Where IdServizioRichiesto=" & idC
if Cdbl(idAccountCliente)>0 then 
   qDel = qDel & "and IdAccountCliente = " & NumForDb(idAccountCliente)
end if 
recordsAffected=0
connMsde.execute qDel , recordsAffected
if recordsAffected=1 then 
   xx=notificaEventoServizioRichiesto(idC,"",0)
end if 
CancellaServizioRichiesto=MsgErrore
end function 

Function deleteImportiServizioRichiesto(IdServizioRichiesto)
Dim MySql 
   on error resume next 
   MySql = ""
   MySql = MySql & " delete from ServizioRichiestoImporti "
   MySql = MySql & " Where IdServizioRichiesto = " & NumforDb(IdServizioRichiesto)
   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("deleteImportiServizioRichiesto:" & MySql & ":" & err.description)
   end if 

end function 

Function addImportoServizioRichiesto(IdServizioRichiesto,IdTipoImporto,DescImporto,Importo)
Dim MySql 
   on error resume next 
   MySql = ""
   MySql = MySql & " delete from ServizioRichiestoImporti "
   MySql = MySql & " where IdServizioRichiesto = " & numForDb(IdServizioRichiesto)
   MySql = MySql & " and IdTipoImporto = '" & apici(IdTipoImporto) & "'"
   ConnMsde.Execute MySql 
 
   if cdbl(Importo)<>0 then 
      MySql = ""
      MySql = MySql & " insert into ServizioRichiestoImporti (IdServizioRichiesto,IdTipoImporto,DescImporto,Importo) values " 
      MySql = MySql & " (" & NumforDb(IdServizioRichiesto)
      MySql = MySql & ",'" & apici(IdTipoImporto) & "'"
      MySql = MySql & ",'" & apici(DescImporto) & "'"   
      MySql = MySql & ", " & NumforDb(Importo)
      MySql = MySql & ") "
	  'response.write MySql 
   end if 
   ConnMsde.Execute MySql 
   if err.number<>0 then 
      writeTrace("addImportoServizioRichiesto:" & MySql & ":" & err.description)
   end if 

end function 

Function addPremioAssiServizioRichiesto(IdServizioRichiesto,DescImporto,Importo)
   if cdbl("0" & Importo)>0 then 
      xx = addImportoServizioRichiesto(IdServizioRichiesto,tipoImporto_PremioAssicurativo,DescImporto,Importo)
   end if 
end function
Function addDirEmisServizioRichiesto(IdServizioRichiesto,DescImporto,Importo)
   if cdbl("0" & Importo)>0 then 
      xx = addImportoServizioRichiesto(IdServizioRichiesto,tipoImporto_DirittiEmissione,DescImporto,Importo)
   end if 
end function
Function addDirInteServizioRichiesto(IdServizioRichiesto,DescImporto,Importo)
   if cdbl("0" & Importo)>0 then 
      xx = addImportoServizioRichiesto(IdServizioRichiesto,tipoImporto_DirittiIntermediazione,DescImporto,Importo)
   end if 
end function
Function addDirInteCollServizioRichiesto(IdServizioRichiesto,DescImporto,Importo)
   if cdbl("0" & Importo)>0 then 
      xx = addImportoServizioRichiesto(IdServizioRichiesto,tipoImporto_DirittiIntermediazioneColl,DescImporto,Importo)
   end if 
end function
Function addSpeseServizioRichiesto(IdServizioRichiesto,DescImporto,Importo)
   if cdbl("0" & Importo)>0 then 
      xx = addImportoServizioRichiesto(IdServizioRichiesto,tipoImporto_Spese,DescImporto,Importo)
   end if 
end function
Function addAltroServizioRichiesto(IdServizioRichiesto,DescImporto,Importo)
   if cdbl("0" & Importo)>0 then 
      xx = addImportoServizioRichiesto(IdServizioRichiesto,tipoImporto_AltroImporto,DescImporto,Importo)
   end if 
end function
Function addProvvServizioRichiesto(IdServizioRichiesto,DescImporto,Importo)
   if cdbl("0" & Importo)>0 then 
      xx = addImportoServizioRichiesto(IdServizioRichiesto,tipoImporto_Provvigioni,DescImporto,Importo)
   end if 
end function

'attiva un evento ; invocata al cambio stato 
Function notificaEventoServizioRichiesto(IdServizioRichiesto,IdAttivita,IdNumAttivita)
Dim MySql,MyRs,idProcesso   
on error resume next 
   MySql = ""
   MySql = MySql & " select * from ServizioRichiesto "
   if Cdbl(IdServizioRichiesto)>0 then 
      MySql = MySql & " Where IdServizioRichiesto=" & IdServizioRichiesto  
   else
      MySql = MySql & " Where IdAttivita='" & IdAttivita & "'"
      MySql = MySql & " and   IdNumAttivita=" & NumforDb(IdNumAttivita)
   end if 
   
   Set MyRs = Server.CreateObject("ADODB.Recordset")
   MyRs.CursorLocation = 3 
   MyRs.Open MySql, ConnMsde    
   if err.number = 0 then 
      if MyRs.eof = false then 
         IdAttivita      = MyRs("IdAttivita")
         IdNumAttivita   = MyRs("IdNumAttivita")
         IdStatoServizio = MyRs("IdStatoServizio")
         IdProdotto      = MyRs("IdProdotto")
		 DescEvento      = ""
         if IdAttivita="CAUZ_DEFI" then 
            idProcesso = "CAUD"
            IdTabella  = "CauzioneDef"
            IdKey      = "IdCauzioneDef=" & IdNumAttivita
            DescEvento = MyRs("DescServizioRichiesto")
         end if 
         if IdProcesso<>"" then 
            xx = createEvento(IdProcesso,IdStatoServizio,Session("LoginIdAccount"),DescEvento,IdTabella,IdKey,true,IdProdotto)
         end if 
      end if 
      MyRs.close 
   end if 
   err.clear 
   
end function 

function CalcolaProvvigioniForn(IdServizioRichiesto)
   on error resume next 
   ConnMsde.execute "calcolaProvvigioneForn" & IdServizioRichiesto
   err.clear
   
end function 

function creaStrutturaDati(IdServizioRichiesto)
Dim retVal 
Dim IdSessione,IdAccountCliente,IdProdotto,IdProdottoTemplate,IdFornitore,IdCompagnia,ImptBaseCalcolo,PrezzoServizio,ImptProvvigioni,Giorni    
Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
   Set DizDatabase = GetDizServizioRichiesto(IdServizioRichiesto,"",0)
    'si trova in FunProcedure 
   retVal=""

   if Cdbl(IdServizioRichiesto)>0 then 
     
      IdSessione         = Session("SessionId")
      IdAccountCliente   = Cdbl("0" & GetDiz(DizDatabase,"N_IdAccountCliente"))
      IdProdotto         = Cdbl("0" & GetDiz(DizDatabase,"N_IdProdotto"))
	  IdProdottoTemplate = Cdbl("0" & GetDiz(DizDatabase,"N_IdProdottoTemplate"))
      IdFornitore        = Cdbl("0" & GetDiz(DizDatabase,"N_IdFornitore"))
      IdCompagnia        = Cdbl("0" & GetDiz(DizDatabase,"N_IdCompagnia"))
      ImptBaseCalcolo    = 0 '????
      PrezzoServizio     = 0 '????
      ImptProvvigioni    = 0 '????
      Giorni             = Cdbl("0" & GetDiz(DizDatabase,"N_GiorniValidita"))
      RetVal=ServizioRichiesto_leggiPrezziProdotto(IdServizioRichiesto,IdSessione,IdAccountCliente,IdProdotto,IdProdottoTemplate,IdFornitore,IdCompagnia,ImptBaseCalcolo,PrezzoServizio,ImptProvvigioni,Giorni) 
   end if    
end function 

   
%>