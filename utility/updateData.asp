<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
flagDebug=false
Oper = ""
IdCompagnia = 0
IdFornitore = 0
IdProdotto  = 0
IdAccount   = 0

operInfo = ucase(trim(Request("op")))
sendData = request("sendData")
if flagDebug=true then 
   response.write sendData & "<br>"
end if 
sendData = DecryptWithKey(sendData,Session("CryptKey"))
if flagDebug=true then 
   response.write operInfo & "<br>"
   response.write sendData & "<br>"
end if 
arD=split(sendData,"|")
for J=lbound(arD) to ubound(arD)
   campo=arD(j)
   ptr=instr(campo,"=")
   if flagDebug=true then 
      response.write "campo " & Campo & "<br>"
	  response.write "ptr   " & ptr & "<br>"
   end if 
   
   if ptr>0 then 
      k=trim(mid(campo,1,ptr-1))
	  v=trim(mid(campo,ptr+1))
	  k=ucase(trim(k))
      if flagDebug=true then 
         response.write "k " & k & "<br>"
	     response.write "v " & v & "<br>"
      end if 	  
	  if k="OPER" then 
	     Oper=ucase(V)
      elseif k="IDCOMPAGNIA" then 
	     IdCompagnia = V
      elseif k="IDFORNITORE" then 
	     IdFornitore = V
      elseif k="IDACCOUNT" then 
	     IdAccount = V		 
      elseif k="IDPRODOTTO"  then 
	     IdProdotto = V
	  end if 
   end if 
next 

IdCompagnia = Cdbl("0" & IdCompagnia)
IdFornitore = Cdbl("0" & IdFornitore)
IdAccount   = Cdbl("0" & IdAccount  )
IdProdotto  = Cdbl("0" & IdProdotto )

If Oper=ucase("UpdCompForn") and operInfo="D" and Cdbl(IdCompagnia)>0 and Cdbl(IdAccount)>0 then 
   MyQ = "" 
   MyQ = MyQ & " delete from AccountCompagnia "
   MyQ = MyQ & " where IdCompagnia = " & IdCompagnia
   MyQ = MyQ & " and   IdAccount = "   & IdAccount
   ConnMsde.execute MyQ 

   InSql = InSql & " Select a.IdProdotto "
   InSql = InSql & " From AccountProdotto a, Prodotto B  "
   InSql = InSql & " Where a.IdAccount = " & IdAccount
   InSql = InSql & " and   a.IdProdotto = B.IdProdotto "
   InSql = InSql & " and   b.IdCompagnia = " & IdCompagnia
   
   MyQ = "" 
   MyQ = MyQ & " delete from AccountProdotto "
   MyQ = MyQ & " where IdAccount   = " & IdAccount
   MyQ = MyQ & " and   IdProdotto  in  (" &  inSql & ")"   
   ConnMsde.execute MyQ
   
end if 
If Oper=ucase("UpdCompForn") and operInfo="I" and Cdbl(IdCompagnia)>0 and Cdbl(IdAccount)>0 then
   MyQ = "" 
   MyQ = MyQ & " Insert into AccountCompagnia ("
   MyQ = MyQ & " IdAccount,IdCompagnia"
   MyQ = MyQ & ") values ("   
   MyQ = MyQ & "  " & IdAccount     
   MyQ = MyQ & " ," & IdCompagnia 
   MyQ = MyQ & ")"
   ConnMsde.execute MyQ 
end if 

If Oper=ucase("UpdProdAccount") and operInfo="D" and Cdbl(IdProdotto)>0 and Cdbl(IdAccount)>0 then 
   MyQ = "" 
   MyQ = MyQ & " delete from AccountProdotto "
   MyQ = MyQ & " where IdAccount   = " & IdAccount
   MyQ = MyQ & " and   IdProdotto  = " & IdProdotto
   ConnMsde.execute MyQ 
   if err.number = 0 then 
	   MyQ = "" 
	   MyQ = MyQ & " delete from AccountProdottoDocAff "
	   MyQ = MyQ & " where IdAccount   = " & IdAccount
	   MyQ = MyQ & " and   IdProdotto  = " & IdProdotto
	   ConnMsde.execute MyQ 
   end if 
  
end if 
If Oper=ucase("UpdProdAccount") and operInfo="I" and Cdbl(IdProdotto)>0 and Cdbl(IdAccount)>0 then
   err.clear
   codiceProdotto = Request("CodiceProdotto")
   MyQ = "" 
   MyQ = MyQ & " Insert into AccountProdotto ("
   MyQ = MyQ & " IdAccount,IdProdotto,CodiceProdotto"
   MyQ = MyQ & ") values ("   
   MyQ = MyQ & "  " & IdAccount
   MyQ = MyQ & ", " & IdProdotto 
   MyQ = MyQ & ",'" & apici(codiceProdotto) & "'" 
   MyQ = MyQ & ")"
   ConnMsde.execute MyQ 
   if err.number = 0 then 
      IdLista = Cdbl("0" & LeggiCampo("Select IdListaDocumento from Prodotto Where IdProdotto=" & IdProdotto ,"IdListaDocumento"))
	  if Cdbl(IdLista)>0 then 
         MyQ = "" 
         MyQ = MyQ & " Insert into AccountProdottoDocAff ("
         MyQ = MyQ & " IdAccount,IdProdotto,IdDocumento,FlagObbligatorio,FlagDataScadenza,DITT,PEFI,PEGI,TIPODOC)"
         MyQ = MyQ & " select " & idAccount & " as IdAccount, " & IdProdotto & " as IdProdotto,IdDocumento"
         MyQ = MyQ & ",FlagObbligatorio,FlagDataScadenza,DITT,PEFI,PEGI,'PROD' as TipoDoc"      
         MyQ = MyQ & " from ServizioDocumento Where IdAnagServizio='LISTA' and IdTipoUtenza='" & IdLista & "'"
         MyQ = MyQ & " and IdDocumento not in (select IdDocumento from AccountProdottoDocAff Where IdAccount=" & IdAccount & " and IdProdotto=" & IdProdotto & " ) "
   	     ConnMsde.execute MyQ
		 'xx=writeTrace(MyQ)
	  end if 
      IdLista = Cdbl("0" & LeggiCampo("Select IdListaAffidamento from Prodotto Where IdProdotto=" & IdProdotto ,"IdListaAffidamento"))
	  if Cdbl(IdLista)>0 then 
         MyQ = "" 
         MyQ = MyQ & " Insert into AccountProdottoDocAff ("
         MyQ = MyQ & " IdAccount,IdProdotto,IdDocumento,FlagObbligatorio,FlagDataScadenza,DITT,PEFI,PEGI,TIPODOC)"
         MyQ = MyQ & " select " & idAccount & " as IdAccount, " & IdProdotto & " as IdProdotto,IdDocumento"
         MyQ = MyQ & ",FlagObbligatorio,FlagDataScadenza,DITT,PEFI,PEGI,'AFFI' as TipoDoc"      
         MyQ = MyQ & " from ServizioDocumento Where IdAnagServizio='LISTA' and IdTipoUtenza='" & IdLista & "'"
         MyQ = MyQ & " and IdDocumento not in (select IdDocumento from AccountProdottoDocAff Where IdAccount=" & IdAccount & " and IdProdotto=" & IdProdotto & " ) "
   	     ConnMsde.execute MyQ
		 'xx=writeTrace(MyQ)
	  end if 
	  
   end if 
   
end if
if flagDebug=true then 
   response.write MyQ
end if 
%>