<%

Function creaMovimento(IdProdotto,IdCompagnia,IdFornitore,DescProdotto,IdAccountCliente,IdAccountRichiedente,IdTipoCreditoClie,IdTipoCreditoRequ,Importo,DataMovimento,TimeMovimento,keyServizio,IdAttivita,IdNumAttivita)
dim IdMovimento,q,IdAnagServ,IdAcc1,IdAcc2,IdAcce
   IdMovimento = 0
   
   DataMovimento=TestNumeroPos(DataMovimento)
   if Cdbl(DataMovimento)=0 then 
      DataMovimento = DtoS()
   end if 
   TimeMovimento=TestNumeroPos(TimeMovimento)
   if Cdbl(TimeMovimento)=0 then 
      TimeMovimento = TimeToS()
   end if 
   IdCompagnia = TestNumeroPos(IdCompagnia)
   if Cdbl(IdCompagnia)=0 then 
      IdCompagnia = LeggiCampo("Select * from Prodotto Where IdProdotto=" & IdProdotto,"IdCompagnia")
   end if 
   IdAnagServ = LeggiCampo("Select * from Prodotto Where IdProdotto=" & IdProdotto,"IdAnagServizio")
   Idacc1     = LeggiCampo("Select * from Cliente Where IdAccount=" & IdAccountCliente,"IdAccountLivello1")
   Idacc2     = LeggiCampo("Select * from Cliente Where IdAccount=" & IdAccountCliente,"IdAccountLivello2")
   Idacc3     = LeggiCampo("Select * from Cliente Where IdAccount=" & IdAccountCliente,"IdAccountLivello3")   
   q = ""
   q = q & " insert into Movimento ("
   q = q & " IdProdotto,DescProdotto,IdFornitore,IdCompagnia"
   q = q & ",IdAccountCliente,IdAccountRichiedente"
   q = q & ",IdAccountLivello1,IdAccountLivello2,IdAccountLivello3"
   q = q & ",DataMovimento,TimeMovimento"
   q = q & ",FlagCalcoloImporti,FlagContabilizzato"
   q = q & ",IdTipoCreditoClie,ImptMovimentoLordo,ImptMovimentoNetto,ImptMovimentoTasse"
   q = q & ",IdTipoCreditoRequ"
   q = q & ",IdAnagServizio,keyServizio,IdAttivita,IdNumAttivita"
   q = q & ") values ("
   q = q & " " & NumForDb(IdProdotto) & ",' " & Apici(DescProdotto) & "'," & NumForDB(IdFornitore) & "," & NumForDB(IdCompagnia)
   q = q & "," & numForDb(IdAccountCliente) & "," & NumForDb(IdAccountRichiedente)
   q = q & "," & NumForDb(Idacc1) & "," & NumForDb(Idacc2) & "," & NumForDb(Idacc3) 
   q = q & "," & NumForDb(DataMovimento) & "," & NumForDb(TimeMovimento)
   q = q & ",0,0"
   q = q & ",'" & apici(IdTipoCreditoClie) & "'," & NumForDB(Importo) & ",0,0"
   q = q & ",'" & apici(IdTipoCreditoRequ) & "'"
   q = q & ",'" & apici(IdAnagServ) & "','" & apici(keyServizio) & "'"   
   q = q & ",'" & apici(IdAttivita) & "'," & NumForDb(IdNumAttivita)   
   q = q & ")"
   'response.write q
   on error resume next 
   ConnMsde.execute q
   'response.write "ee:" & err.number 
   'response.write "ed:" & err.description 
   if err.number=0 then 
      IdMovimento = GetTableIdentity("Movimento")
   else
      xx=writeTrace(q & "::" & err.description)
   end if 
   creaMovimento = idMovimento 
end function 
%>


