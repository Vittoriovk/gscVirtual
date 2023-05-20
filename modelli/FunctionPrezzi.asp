<%

Function GetPrezzoFissoAccount(dizPrezzi,IdProdotto,IdAccountFornitore,IdAccount)
Dim DizAcc,conta,xx,chiave,q,idAcc 
Dim ReadRs,Trovato
    on error resume next 
    xx=LeggiCampo("select * from prodotto Where IdProdotto=" & IdProdotto,"FlagPrezzoFisso")
    if Cdbl("0" & xx)=1 then 
       Set DizAcc = CreateObject("Scripting.Dictionary")
       xx=GetGerarchiaAccount(DizAcc,IdAccount)
	   trovato=false 
	   'se non trovo il listino per account prendo quello definito al fornitore
	   xx=SetDiz(DizAcc,"G4",IdAccountFornitore)
	   for conta=1 to 4 
	      chiave = "G" & conta
	      If DizAcc.Exists(chiave) and Trovato=false Then
	         idAcc = DizAcc(chiave)
	         q = GetQueryListinoAccount(IdProdotto,IdAccountFornitore,IdAcc)
			 'response.write q
             set ReadRs = ConnMsde.execute(q)
             if err.number=0 then 
                if not ReadRs.eof then 
			       trovato = true 
		           xx=SetDiz(dizPrezzi,"PrezzoCompagnia"    ,ReadRs("PrezzoCompagnia"))
		           xx=SetDiz(dizPrezzi,"PrezzoFornitore"    ,ReadRs("PrezzoFornitore"))
		           xx=SetDiz(dizPrezzi,"PrezzoDistribuzione",ReadRs("PrezzoDistribuzione"))
			       xx=SetDiz(dizPrezzi,"PrezzoListino"      ,ReadRs("PrezzoListino"))
			    end if 
             end if 
             ReadRs.close		  
	      end if 
       next 
    end if 
end function 


Function GetQueryListinoAccount(IdProdotto,IdAccountFornitore,IdAccount)
Dim q
  q = ""
  q = q & " select * from AccountProdottoListino " 
  q = q & " Where IdProdotto = " & IdProdotto
  q = q & " and   IdAccountFornitore = " & IdAccountFornitore
  q = q & " and   IdAccount = " & IdAccount 
  q = q & " and   ValidoDal <= " & Dtos()
  q = q & " and   ValidoAl >= " & Dtos()
  GetQueryListinoAccount = q
end function 

Function GetGerarchiaAccount(DizAccount,IdAccount)
Dim tipoAcc,q,conta,chiave 
Dim ReadRs

   conta=1
   xx=SetDiz(DizAccount,"G1",IdAccount)
   TipoAcc=ucase(LeggiCampo("select * from Account Where IdAccount=" & IdAccount,"IdTipoAccount"))
   q = ""
   if TipoAcc="CLIE" then 
      q = "select * from cliente Where IdAccount = " & IdAccount 
   end if 
   if TipoAcc="COLL" then 
      q = "select * from Collaboratore Where IdAccount = " & IdAccount 
   end if 
   if q<>"" then
      set ReadRs = ConnMsde.execute(q)
      if err.number=0 then 
         if not ReadRs.eof then 
		    
            if Cdbl(ReadRs("IdAccountLivello2"))>0 then 
			   conta=conta+1
			   xx=SetDiz(DizAccount,"G" & conta,ReadRs("IdAccountLivello2"))
			end if 
            if Cdbl(ReadRs("IdAccountLivello1"))>0 then 
			   conta=conta+1
			   chiave="G" & conta 
			   xx=SetDiz(DizAccount,chiave,ReadRs("IdAccountLivello1"))
			end if 
         end if 
      end if 
      ReadRs.close
   end if 
   err.clear
end function 

%>