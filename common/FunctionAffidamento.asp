<%
'query per prendere tutti i documenti necessari 
Function funDoc_getDocAffidamento(IdCompagnia,andCond)   
Dim MySqlDoc
   MySqlDoc = "" 
   MySqlDoc = MySqlDoc & " select C.IdDocumento,C.DescDocumento "
   MySqlDoc = MySqlDoc & ",Max(FlagObbligatorio) as FlagObbligatorio"
   MySqlDoc = MySqlDoc & ",Max(FlagDataScadenza) as FlagDataScadenza"
   MySqlDoc = MySqlDoc & " From AccountProdottoDocAff a, Prodotto B, Documento C"
   MySqlDoc = MySqlDoc & " Where A.IdProdotto = b.IdProdotto "
   MySqlDoc = MySqlDoc & " And   B.IdCompagnia = " & IdCompagnia
   MySqlDoc = MySqlDoc & " and   A.IdDocumento = C.IdDocumento "
   MySqlDoc = MySqlDoc & " and   B.IdAnagServizio = 'CAUZ_PROV'"
   MySqlDoc = MySqlDoc & andCond
   MySqlDoc = MySqlDoc & " group by C.IdDocumento,C.DescDocumento"
   funDoc_getDocAffidamento = MySqlDoc
end Function 

Function funDoc_getDocAccount(IdDocumento,IdAccount,AllaData)
Dim MySqlDoc
   MySqlDoc = ""
   MySqlDoc = MySqlDoc & " select A.IdAccountDocumento,A.IdDocumento,A.IdUpload,A.IdTipoValidazione"
   MySqlDoc = MySqlDoc & ",isNull(b.descBreve,'') as DescBreve"
   MySqlDoc = MySqlDoc & ",isNull(b.PathDocumento,'') as PathDocumento"
   MySqlDoc = MySqlDoc & ",isNull(b.ValidoDal,10000101) as ValidoDal"
   MySqlDoc = MySqlDoc & ",isNull(b.ValidoAl ,99991231) as ValidoAl"
   MySqlDoc = MySqlDoc & " from AccountDocumento A "
   MySqlDoc = MySqlDoc & " left join Upload b on a.idUpload = b.idupload"
   MySqlDoc = MySqlDoc & " Where A.IdDocumento = " & IdDocumento 
   MySqlDoc = MySqlDoc & " And   A.IdAccount = " & IdAccount 
   MySqlDoc = MySqlDoc & " And   isNull(b.ValidoDal,10000101) <= " & AllaData
   MySqlDoc = MySqlDoc & " And   isNull(b.ValidoAl ,99991231) >= " & AllaData
   funDoc_getDocAccount = MySqlDoc
end Function 

Function funDoc_ValutaStatoDoc(IdTipoValidazione)
Dim opAmm
   opAmm=""
   if IdTipoValidazione="VALIDO" then
      opAmm="" 
   elseif IdTipoValidazione="" then 
      opAmm="I"
   elseif IdTipoValidazione="NONRIC" then
      opAmm="UD"
   elseif IdTipoValidazione="NONVAL" then
      opAmm="D"   
   end if 
   funDoc_ValutaStatoDoc = opAmm

end function 

Function funDoc_DescrizioneStatoDoc(IdTipoValidazione)
Dim opAmm
   opAmm=""
   if IdTipoValidazione="" or IdTipoValidazione="NONRIC" then 
      OpAmm = "Documento caricato"
   else 
      OpAmm = LeggiCampo("Select * from TipoValidazione Where IdTipoValidazione='" & IdTipoValidazione & "'","DescTipoValidazione")
   end if 
   funDoc_DescrizioneStatoDoc = opAmm

end function 

function funDoc_DelAccDoc(IdAccountDocumento)
Dim KK,Acc,esi,Esito,qSel
    Esito = ""
	KK    = Cdbl("0" & IdAccountDocumento)
	Acc   = Cdbl("0" & LeggiCampo("select IdUpload from AccountDocumento where IdAccountDocumento=" & kk,"IdUpload"))

	'verifico se ci sono documenti legati in affidamento 
	
	if Cdbl(kk>0) and esito="" then 
	   qSel = "select top 1 * from AffidamentoRichiestaCompDoc Where IdAccountDocumento = " & kk
	   'response.write qSel 
	   esi = LeggiCampo(qSel,"IdAccountDocumento")
	   esi = Cdbl("0" & esi)
	   if esi>0 then 
	      Esito = "non cancellabile esistono operazioni in corso"
	   end if 
	end if 
	
	if Cdbl(kk>0) and esito="" then 
	   ConnMsde.execute "Delete From AccountDocumento where IdAccountDocumento=" & kk
	   if Cdbl(Acc)>0 then 
	      ConnMsde.execute "Delete From Upload where IdUpload=" & acc
       end if 
    end if 
    funDoc_DelAccDoc = Esito 
end function 

function funDoc_VerificaDocAff(IdAccount)
Dim MySqlDoc,IdDocumento,Qsel
Dim RsTmp,retVal
    'prelevo i documenti di tutte le compagnie 
    retVal="OK"
    Set RsTmp = Server.CreateObject("ADODB.Recordset")
    MySqlDoc = funDoc_getDocAffidamento("")
    RsTmp.CursorLocation = 3 
    RsTmp.Open MySqlDoc, ConnMsde
    if err.number=0 then 
       Do While Not RsTmp.EOF 
          IdDocumento=RsTmp("IdDocumento")
          Qsel = funDoc_getDocAccount(IdDocumento,IdAccount,Dtos())
		  'response.write Qsel
          IdTipoValidazione = LeggiCampo(Qsel,"IdTipoValidazione")
		  IdUpload          = Cdbl("0" & LeggiCampo(Qsel,"IdUpload"))
		  if IdTipoValidazione="NONVAL" or IdTipoValidazione="" or Cdbl(IdUpload)=0 then 
		     retVal="KO"
          end if 
          rsTmp.MoveNext
		  'response.end 
       Loop
    End if 
    rsTmp.close 
    err.clear 
	funDoc_VerificaDocAff=RetVal
end function

'restituisce 
'  OK    = documentazione completa e validata 
'  COMPL = documentazione completa : ci sono tutti i documenti richiesti
'  INTEG = documentazione da integrare 
'        = documentazione incompleta di default 
Function funDoc_StatoDocum(IdRichiestaAffidamentoComp)
dim retVal,IdRicComp,q
Dim MySqlDoc,Qsel,RsTmpl,allValidi,allPresenti,daIntegrare,haDocumenti
    haDocumenti=false 
    allValidi  =true 
	allPresenti=true 
    daIntegrare=false 

    IdRicComp=IdRichiestaAffidamentoComp 
	if Cdbl(IdRicComp)>0 then 
	   'leggo i documenti per la compagnia 
       retVal="OK"
       Set RsTmp = Server.CreateObject("ADODB.Recordset")
	   MySqlDoc = ""
       MySqlDoc = MySqlDoc  & " select A.*,isnull(B.IdTipoValidazione,'') as IdTipoValidazione"
       MySqlDoc = MySqlDoc  & " from AffidamentoRichiestaCompDoc A "
	   MySqlDoc = MySqlDoc  & " left join AccountDocumento B on A.idAccountDocumento = b.idAccountDocumento"
	   MySqlDoc = MySqlDoc  & " Where a.IdAffidamentoRichiestaComp=" & IdRicComp
	   'response.write MySqldoc
       RsTmp.CursorLocation = 3 
       RsTmp.Open MySqlDoc, ConnMsde
       if err.number=0 then 
          Do While Not RsTmp.EOF 
		     haDocumenti=true 
			 pathdocumento=""
			 if cdbl(RsTmp("idAccountDocumento"))>0 then
			    qSel = "select * from Upload Where IdUpload=(select idUpload from AccountDocumento Where IdAccountDocumento=" & RsTmp("idAccountDocumento") & "  )"
                pathdocumento=LeggiCampo(qSel,"PathDocumento")
			 end if 
		     if cdbl(RsTmp("idAccountDocumento"))=0 or pathDocumento="" then 
			    if RsTmp("FlagObbligatorio")=1 then 
			       allPresenti=false
				   allValidi=false 
			    end if 
			 else
			    'response.write "ecco 2 " & RsTmp("IdTipoValidazione")
			    if RsTmp("IdTipoValidazione")<>"VALIDO" then 
				   allValidi=false
				end if 
				if RsTmp("IdTipoValidazione")="NONVAL" then 
				   daIntegrare=true 
                end if 
				
			 end if 
             rsTmp.MoveNext
             
          Loop
       End if 
       rsTmp.close 
       err.clear 	   
	end if 
	'default non posso fare nulla
	retVal=""
	if haDocumenti then 
	   'response.write "ecco 3 " & allValidi & " " & daIntegrare
       if allValidi=true then 
	      retVal="OK"
       elseif daIntegrare=true then 
	      retVal="INTEG"
	  elseif allPresenti=true then 
	      retVal="COMPL"
	   end if 
	end if 
	funDoc_StatoDocum=RetVal
end function

Function funDoc_getIdRichiestaComp(IdAffidamentoRichiesta,IdCompagnia)
dim IdRicComp,q 
    q = ""
    q = q & " Select * from AffidamentoRichiestaComp"
    q = q & " Where IdAffidamentoRichiesta= " & IdAffidamentoRichiesta
    q = q & " and IdCompagnia = " & IdCompagnia  
	
    IdRicComp=Cdbl("0" & Leggicampo(q,"IdIdAffidamentoRichiestaComp"))
	funDoc_getIdRichiestaComp=IdRicComp
end function

'oper = V per sola verifica
'oper = I per verifica ed inserimento
function caricaImportoAffidamento(oper,IdAccount,IdCompagnia,IdFornitore,idRecMod,ValidoDal,ValidoAl,imptComplessivo,ImptSingolaPolizza,AffidamentoUsato)
dim ImptMinimo,mySql,MsgErrore,qSel,xx,qIns,qUpd,ImptStornato,ImptUsato,ImptImpegnato  
   MsgErrore = ""

   if len(ValidoDal)<>8 or len(ValidoAl)<>8 or ValidoDal > ValidoAl then 
      MsgErrore=MsgErrore & "Date inserite non valide;"
   end if    
   if MsgErrore="" then 
      qSel = ""
      qSel = qSel & " select top 1 * from AccountCreditoAffi "
	  qSel = qSel & " Where IdAccount = " & IdAccount
	  qSel = qSel & " and   IdAccountCreditoAffi <> " & IdRecMod
	  qSel = qSel & " and   IdCompagnia = " & IdCompagnia 
	  qSel = qSel & " and ("
	  qSel = qSel & "     (ValidoDal<= " & ValidoDal & " and ValidoAl>=" & ValidoDal &") "
	  qSel = qSel & "  or (ValidoDal<= " & ValidoAl  & " and ValidoAl>=" & ValidoAl  &") "
	  qSel = qSel & "  or (ValidoDal>= " & ValidoDAl & " and ValidoAl<=" & ValidoAl  &") "
	  qSel = qSel & "     )"
	  'response.write qSel 
      xx=cdbl("0" & LeggiCampo(qSel,"IdAccount"))
	  if Cdbl(xx)>0 then 
         MsgErrore= "Date incongruenti con altro periodo"
      end if 
   end if 
   if MsgErrore="" and oper="I" then 
      err.clear
      if cdbl(IdRecMod)=0 then 
         qIns = ""
         qIns = qins & "INSERT INTO AccountCreditoAffi("
         qIns = qins & " IdAccount,IdTipoCredito,ValidoDal,ValidoAl,IdFornitore"
         qIns = qins & ",IdCompagnia,ImptComplessivo,ImptSingolaPolizza"
         qIns = qins & ") VALUES (" & IdAccount & ",'AFFI'," & ValidoDal & "," & ValidoAl & "," & IdFornitore
         qIns = qins & "," & IdCompagnia & ",0,0)"
	     'response.write qIns
	     ConnMsde.execute qIns 
	     if err.number<>0 then 
	        msgErrore=err.description 
	     else
	        IdRecMod = GetTableIdentity("AccountCreditoAffi")
	     end if 
	  end if 
	  qUpd = ""
	  qUpd = qUpd & "update AccountCreditoAffi set "
	  qUpd = qUpd & " ValidoDal=" & ValidoDal
	  qUpd = qUpd & ",ValidoAl=" & ValidoAl
	  qUpd = qUpd & ",IdCompagnia=" & IdCompagnia 
	  qUpd = qUpd & ",IdFornitore=" & IdFornitore
	  qUpd = qUpd & ",ImptComplessivo=" & NumForDb(ImptComplessivo)
	  qUpd = qUpd & ",ImptSingolaPolizza=" & NumForDb(ImptSingolaPolizza)
	  qUpd = qUpd & " Where IdAccountCreditoAffi = " & IdRecMod
	  'response.write qUpd 
	  connMsde.execute qUpd 
	  
	  'controllo se devo registrare su AccountCreditoAffiTotali
      qSel = ""
	  qSel = qSel & " select * from AccountCreditoAffiTotali "
	  qSel = qSel & " Where IdCompagnia = " & IdCompagnia
	  qSel = qSel & " and IdAccount = " & IdAccount
	  idC = "0" & LeggiCampo(qSel,"IdCompagnia")
	  'non esiste la riga la devo inserire 
	  if Cdbl(idC)=0 then 
         qIns = ""
		 qIns = qIns & " Insert into AccountCreditoAffiTotali "
		 qIns = qIns & "(IdAccount,IdCompagnia,ImptUtilizzato,ImptUtilizzatoStornato"
		 qIns = qIns & ",ImptIniziale,ImptInizialeStornato,ImptAffidatoImpegnato) "
		 qIns = qIns & " values "
		 qIns = qIns & "(" & IdAccount & "," & IdCompagnia & ",0,0"
		 qIns = qIns & ",0,0,0) "
		 ConnMsde.execute qIns 
	  end if 
   end if 
   if Cdbl(AffidamentoUsato)>0 then 
      impt = "0" & LeggiCampo(qSel,"ImptIniziale")
	  if Cdbl(impt)<>Cdbl(AffidamentoUsato) then 
	     qUpd = ""
	     qUpd = qUpd & "update AccountCreditoAffiTotali set "
	     qUpd = qUpd & " ImptIniziale=" & NumForDb(AffidamentoUsato)
	     qUpd = qUpd & " Where IdAccount = " & IdAccount
		 qUpd = qUpd & " and IdCompagnia = " & IdCompagnia
	     'response.write qUpd 
	     connMsde.execute qUpd 
	     
	  end if 
   end if 
   
   caricaImportoAffidamento=MsgErrore 

end function 
'IdTipoValidazione	DescTipoValidazione
'DAVALI	Da Validare
'INVALI	In Validazione
'NONRIC	Non Richiesta
'NONVAL	Non Valido
'VALIDO	Validato

Function AggiornaRichiestaAffidamento(IdAffidamentoRichiesta)
Dim MySqlDoc,Qsel,descStato,qUpd
Dim RsTmp,retVal,oneAffi,allClos 
    'prelevo i documenti di tutte le compagnie 
	err.clear
    Set RsTmp = Server.CreateObject("ADODB.Recordset")
	MySqlDoc = ""
    MySqlDoc = MySqlDoc & " select A.*,B.flagStatoFinale,B.DescStatoServizio,C.DescCompagnia  "
    MySqlDoc = MySqlDoc & " from AffidamentoRichiestaComp A,StatoServizio B, Compagnia C"
    MySqlDoc = MySqlDoc & " Where a.IdAffidamentoRichiesta=" & IdAffidamentoRichiesta
    MySqlDoc = MySqlDoc & " and A.IdStatoAffidamento = B.IdStatoServizio"
	MySqlDoc = MySqlDoc & " and A.IdCompagnia = C.IdCompagnia"
	'response.write MySqlDoc
    RsTmp.CursorLocation = 3 
    RsTmp.Open MySqlDoc, ConnMsde
	
	oneAffi = false 
	allClos = false 
    if err.number=0 then 
	   allClos = true 
	   descStato = ""
       Do While Not RsTmp.EOF 
	      descStato = descStato & rsTmp("DescCompagnia") & ":" & rsTmp("DescStatoServizio") & ";"
          if rsTmp("flagStatoFinale")<>1 then 
		     'response.write "qui"
		     allClos = false 
		  end if 
	   
          if rsTmp("IdStatoAffidamento")="AFFI" then 
		     oneAffi = true 
		  end if 
          rsTmp.MoveNext
		  'response.end 
       Loop
    End if 
    rsTmp.close 
	retVal = "LAVO"
	if allClos then
       if oneAffi=true then 
	      retVal = "AFFI"
       else
	      retVal = "ANNU"
	   end if 
	end if 
	qUpd = ""
    qUpd = qUpd & " Update AffidamentoRichiesta set "
    qUpd = qUpd & " IdStatoAffidamento='" & retVal & "'"
    qUpd = qUpd & ",NoteAffidamento='" & apici(descStato) & "'"
	if allClos=true then 
	   qUpd = qUpd & ",DataChiusura= " & Dtos()
	end if 
    qUpd = qUpd & " Where IdAffidamentoRichiesta=" & IdAffidamentoRichiesta
	'response.write qUpd
    ConnMsde.execute qUpd 
end function 

Function GetDettaglioAffComp(Dizionario,IdAffidamentoRichiestaComp,IdAffidamentoRichiesta,IdCompagnia)
Dim q,esito
  
   q = ""
   q = q & " select * from AffidamentoRichiestaComp "
   q = q & " Where ( IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
   q = q & " or (   IdAffidamentoRichiesta=" & IdAffidamentoRichiesta 
   q = q & "        and IdCompagnia=" & IdCompagnia 
   q = q & "    )"
   q = q & "       )"
   'response.write q
   esito = GetInfoRecordset(Dizionario,q)
   GetDettaglioAffComp = esito 
   
end function

Function GetTotaliAffidamentoComp(Dizionario,IdAccount,IdCompagnia)
Dim q,esito,totU
   q = ""
   q = q & " select * from AccountCreditoAffiTotali "
   q = q & " Where IdAccount = " & IdAccount 
   q = q & " and IdCompagnia = " & IdCompagnia
   esito = GetInfoRecordset(Dizionario,q)
   'calcolo il totale usato 
   if esito=true then 
      totU = 0
	  totU = totU + Cdbl(getDiz(Dizionario,"ImptUtilizzato")) + Cdbl(getDiz(Dizionario,"ImptIniziale"))
	  totU = totU - Cdbl(getDiz(Dizionario,"ImptUtilizzatoStornato")) - Cdbl(getDiz(Dizionario,"ImptInizialeStornato"))
	  xx=Setdiz(Dizionario,"TotaleImpegnato",totU)
   end if 
   'xx=DumpDic(Dizionario,"xx")
   GetTotaliAffidamentoComp = esito 
   
end function

Function getInfoProcessoAffi(Dizionario,IdStato,IdFlussoProcesso)
Dim rVal,MySql,esito  
   rVal = ""
   MySql = ""
   MySql = MySql & " Select * "
   MySql = MySql & " from " & getStatoFlusso() 
   MySql = MySql & " Where StFl.IdStatoSorgente='" & IdStato & "'"
   MySql = MySql & " and   StFl.IdFlussoProcesso  in ('*','" & IdFlussoProcesso & "')"
   esito = GetInfoRecordset(Dizionario,MySql)
   'xx=DumpDic(Dizionario,"xx")
   getInfoProcessoAffi = esito 
   
end function

Function GetDettaglioCauzioneComp(Dizionario,IdCauzioneCompagnia,IdCauzione,IdCompagnia)
Dim q,esito
  
   q = ""
   q = q & " select * from CauzioneCompagnia "
   q = q & " Where ( IdCauzioneCompagnia=" & IdCauzioneCompagnia
   q = q & " or (   IdCauzione=" & IdCauzione 
   q = q & "        and IdCompagnia=" & IdCompagnia 
   q = q & "    )"
   q = q & "       )"
   'response.write q
   esito = GetInfoRecordset(Dizionario,q)
   GetDettaglioCauzioneComp = esito 
   
end function


%>