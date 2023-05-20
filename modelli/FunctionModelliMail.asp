<%
function inviaMailFormazione(IdFormazione)
   Set Dizionario = Server.CreateObject("Scripting.Dictionary")
   xx=GetInfoRecordset(Dizionario,"select * from Formazione Where IdFormazione=" & IdFormazione)
   
   eMailNotificaCliente = getValueOfDic(Dizionario,"eMailNotificaCliente")
   linkPiattaforma      = getValueOfDic(Dizionario,"linkPiattaforma")
   userPiattaforma      = getValueOfDic(Dizionario,"userPiattaforma")
   passPiattaforma      = getValueOfDic(Dizionario,"passPiattaforma") 

   idProdotto         = getValueOfDic(Dizionario,"IdProdotto") 
   IdAccountFornitore = getValueOfDic(Dizionario,"IdAccountFornitore") 
   descProdotto       = getInfoProdotto(idProdotto,"DescProdotto")
   MailDocForn = LeggiCampo("select * from AccountProdotto Where IdAccount=" & IdAccountFornitore & " and IdProdotto=" & IdProdotto,"MailDocumentazione")
  
   addInfo = ""
   addInfo = addInfo  & " Conferma attivazione corso di formazione: <b>" & DescProdotto & "</b><br><br>"

   if linkPiattaforma<>"" then 
	  addInfo = addInfo  & " Il corso e' fruibile al seguente indirizzo: " & linkPiattaforma &  " <br>"
      if userPiattaforma<>"" and passPiattaforma<>"" then 
	     addInfo = addInfo  & " acceda con le seguenti le credenziali : <br>"
	     addInfo = addInfo  & " utenza = " & userPiattaforma & "<br>"
	     addInfo = addInfo  & " password = " & passPiattaforma & "<br>"
	     addInfo = addInfo  & " <br>"
     else 
	     addInfo = addInfo  & " contatti il suo referente per ricevere le credenziali di accesso. <br>"
	     addInfo = addInfo  & " <br>"
     end if 
   else
      addInfo = addInfo  & " Per fruire del corso si rivolga al suo referente che le le fornira' i dati di accesso. <br>"
   end if    
 
  addInfo = addInfo  & " Buona formazione .."
  Nominativo = getValueOfDic(Dizionario,"Cognome") & " " & getValueOfDic(Dizionario,"Nome")
  TextHtml=""
  TextText=""
  xx=CreaTestoMail(Nominativo,addInfo,TextHtml,TextText)
  if eMailNotificaCliente<>"" then 
     xx=SendMailMessageHTMLWithAttach("", eMailNotificaCliente, "","Attivazione Utenza per il Corso:" & descProdotto, TextText, TextHtml, "", false) 
  end if 
  if MailDocForn<>"" then 
     xx=SendMailMessageHTMLWithAttach("", MailDocForn, "","Attivazione Utenza per il Corso:" & descProdotto, TextText, TextHtml, "", false)   
  end if 

end function 

function inviaMailFormazioneServizio(IdServizioRichiesto)
Dim q
   Set Dizionario = Server.CreateObject("Scripting.Dictionary")
   q = ""
   q = q & " select * "
   q = q & " from formazione a, ServizioRichiesto b "
   q = q & " Where B.IdServizioRichiesto = " & IdServizioRichiesto
   q = q & " and   b.IdAttivita = 'FORMAZ'"
   q = q & " and   b.IdNumAttivita = A.IdFormazione"
   
   xx=GetInfoRecordset(Dizionario,q)
   
   eMailNotificaCliente = getValueOfDic(Dizionario,"MailNotificaUtilizzatore")
   linkPiattaforma      = getValueOfDic(Dizionario,"linkPiattaforma")
   userPiattaforma      = getValueOfDic(Dizionario,"userPiattaforma")
   passPiattaforma      = getValueOfDic(Dizionario,"passPiattaforma") 

   idProdotto         = getValueOfDic(Dizionario,"IdProdotto") 
   IdAccountFornitore = getValueOfDic(Dizionario,"IdAccountFornitore") 
   descProdotto       = getInfoProdotto(idProdotto,"DescProdotto")
   MailDocForn = LeggiCampo("select * from AccountProdotto Where IdAccount=" & IdAccountFornitore & " and IdProdotto=" & IdProdotto,"MailDocumentazione")
  
   addInfo = ""
   addInfo = addInfo  & " Conferma attivazione corso di formazione: <b>" & DescProdotto & "</b><br><br>"

   if linkPiattaforma<>"" then 
	  addInfo = addInfo  & " Il corso e' fruibile al seguente indirizzo: " & linkPiattaforma &  " <br>"
      if userPiattaforma<>"" and passPiattaforma<>"" then 
	     addInfo = addInfo  & " acceda con le seguenti le credenziali : <br>"
	     addInfo = addInfo  & " utenza = " & userPiattaforma & "<br>"
	     addInfo = addInfo  & " password = " & passPiattaforma & "<br>"
	     addInfo = addInfo  & " <br>"
     else 
	     addInfo = addInfo  & " contatti il suo referente per ricevere le credenziali di accesso. <br>"
	     addInfo = addInfo  & " <br>"
     end if 
   else
      addInfo = addInfo  & " Per fruire del corso si rivolga al suo referente che le le fornira' i dati di accesso. <br>"
   end if    
 
  addInfo = addInfo  & " Buona formazione .."
  Nominativo = getValueOfDic(Dizionario,"Cognome") & " " & getValueOfDic(Dizionario,"Nome")
  TextHtml=""
  TextText=""
  xx=CreaTestoMail(Nominativo,addInfo,TextHtml,TextText)
  if eMailNotificaCliente<>"" then 
     xx=SendMailMessageHTMLWithAttach("", eMailNotificaCliente, "","Attivazione Utenza per il Corso:" & descProdotto, TextText, TextHtml, "", false) 
  end if 
  if MailDocForn<>"" then 
     xx=SendMailMessageHTMLWithAttach("", MailDocForn, "","Attivazione Utenza per il Corso:" & descProdotto, TextText, TextHtml, "", false)   
  end if 

end function 

%>