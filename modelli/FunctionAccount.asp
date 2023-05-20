<%
Function GetEmptyAccount()
Dim MySql,v_id
   v_id=0
   'controllo esistenza di uno vuoto 
   MySql = ""
   MySql = MySql & " select top 1 idAccount From Account "
   MySql = MySql & " where IdAzienda = 0" 
   v_ret = "0" & LeggiCampo(MySql,"IdAccount")
   if Cdbl(v_ret)=0 then 
	   MySql = ""
	   MySql = MySql & " insert into Account ("
	   MySql = MySql & " IdAzienda,IdTipoAccount,UserId,Password,Abilitato)"
	   MySql = MySql & " values (0,'','','',0)"
	   ConnMsde.execute MySql 
	   if err.number=0 then 
	      v_id=GetTableIdentity("Account")   
	   end if 
   else
       v_id=cdbl(v_ret)
   end if 
   GetEmptyAccount=v_id
end function 

Function GetTempAccount()
Dim MySql,v_id,flagClean,oggi
   Oggi = Dtos()
   v_id=0
   flagClean=0
   'controllo esistenza di uno vuoto 
   MySql = ""
   MySql = MySql & " select top 1 idAccount From Account "
   MySql = MySql & " where IdAzienda = 0" 
   MySql = MySql & " and   IdTipoAccount like 'WORK%'" 
   MySql = MySql & " and   IdTipoAccount < 'WORK" & oggi & "'"
   
   v_ret = "0" & LeggiCampo(MySql,"IdAccount")
   if Cdbl(v_ret)=0 then 
	   MySql = ""
	   MySql = MySql & " insert into Account ("
	   MySql = MySql & " IdAzienda,IdTipoAccount,UserId,Password,Abilitato)"
	   MySql = MySql & " values (0,'WORK" & oggi & "','','',0)"
	   ConnMsde.execute MySql 
	   if err.number=0 then 
	      v_id=GetTableIdentity("Account")   
	   end if 
   else
       flagClean=1
       v_id=cdbl(v_ret)
   end if 
   'cancello le tabelle collegate all'account 
   if flagClean=1 then 
	connMsde.execute "delete from AccountSede Where idAccount=" & v_id
	connMsde.execute "delete from AccountContatto Where idAccount=" & v_id
	connMsde.execute "delete from AccountCompagnia Where idAccount=" & v_id
   end if 
   GetTempAccount=v_id
end function 

Function UpdateLoginAccount(IdAccount,UserId,Password,Attivo,Blocco)
Dim MySql , Esito , v_ret,Nome
on error resume next 
   Esito = ""

   'controllo duplicazione 
   if Attivo="S" then 
	   MySql = ""
	   MySql = MySql & " select top 1 idAccount From Account "
	   MySql = MySql & " where IdAccount<>" & IdAccount
	   MySql = MySql & " and   FlagAttivo='S'"  
	   MySql = MySql & " and   UserId='" & apici(UserId) & "'" 
	   v_ret = "0" & LeggiCampo(MySql,"IdAccount")   
	   if cdbl(v_ret)>0 then 
          Esito  = "UserId Esistente : utente non attivato"
          Attivo = "N"
       end if 
   end if 
   
   MySql = ""
   MySql = MySql & " update Account set "
   MySql = MySql & " UserId='"      & apici(UserId)    & "'"
   MySql = MySql & ",PassWord ='"   & apici(PassWord)  & "'"
   MySql = MySql & ",FlagAttivo ='" & apici(Attivo)    & "'"
   if Attivo = "S" then
      MySql = MySql & ",Abilitato = 1 " 
   else
      MySql = MySql & ",Abilitato = 0 " 
   end if 
   MySql = MySql & ",DescBlocco ='" & apici(Blocco)    & "'"
   MySql = MySql & " where IdAccount = " & IdAccount
   'response.write MySql
   ConnMsde.execute MySql 
   if err.number<>0 then 
      Esito = Err.description 
   end if 
   
   UpdateLoginAccount=Esito
   
end function 

Function UpdateAccount()
Dim v_id , MySql , Esito , v_ret,Nome,IdAz
on error resume next 
   Esito = ""
   v_id = cdbl("0" & GetDiz(DizDatabase,"IdAccount"))
   IdAz = cdbl("0" & GetDiz(DizDatabase,"IdAzienda"))
   Nome = GetDiz(DizDatabase,"Nominativo")
   if Nome="" then 
      Nome=trim(GetDiz(DizDatabase,"Cognome") & " " & GetDiz(DizDatabase,"Nome"))
   end if 
   
   'controllo duplicazione 
   MySql = ""
   MySql = MySql & " select top 1 idAccount From Account "
   MySql = MySql & " where IdAccount<>" & v_id 
   MySql = MySql & " and   IdAzienda = " & IdAz 
   MySql = MySql & " and   FlagAttivo='S'"  
   MySql = MySql & " and   UserId='" & apici(GetDiz(DizDatabase,"UserId")) & "'" 
   v_ret = LeggiCampo(MySql,"IdAccount")
   
   if v_ret<>"" then 
      Esito = "Codice Utente Esistente "
   else
      if cdbl(v_id)=0 then 
         MySql = ""
         MySql = MySql & " insert into Account ("
		 MySql = MySql & " IdAzienda,IdTipoAccount,UserId"
		 MySql = MySql & ",Password,Abilitato,Nominativo,PartitaIva,CodiceFiscale"
		 MySql = MySql & ",Indirizzo1,Indirizzo2,Cap,Comune,Provincia,Settore,email1,email2"
		 MySql = MySql & ",Telefono,FlagAttivo,DescBlocco,Cognome,Nome,IdTipoUsoServizio,IdProfiloAbilitazione)"		 
		 MySql = MySql & " values ("
		 MySql = MySql & "  " & GetDiz(DizDatabase,"IdAzienda")
		 MySql = MySql & ",'" & apici(GetDiz(DizDatabase,"IdTipoAccount")) & "'"
		 MySql = MySql & ",'" & apici(GetDiz(DizDatabase,"UserId")) & "'"
		 MySql = MySql & ",'',1,'','',''"
		 MySql = MySql & ",'','','','','','','',''"
		 MySql = MySql & ",'','S','','','','',0" 
		 MySql = MySql & " )"
		 
		 ConnMsde.execute MySql 
		 if err.number<>0 then 
		    Esito = Err.description 
		 else
		    v_id=GetTableIdentity("Account")
			xx=SetDiz(DizDatabase,"IdAccount",v_id)
		 end if 
      end if 	  
      if cdbl(v_id)>0 then 
         MySql = ""
         MySql = MySql & " update Account set "
         MySql = MySql & " IdTipoAccount='" & apici(GetDiz(DizDatabase,"IdTipoAccount")) & "'"
         MySql = MySql & ",PassWord     ='" & apici(GetDiz(DizDatabase,"PassWord"))      & "'"
         MySql = MySql & ",Abilitato    = " &       GetDiz(DizDatabase,"Abilitato")
         MySql = MySql & ",Nominativo   ='" & apici(Nome)                                & "'"
         MySql = MySql & ",PartitaIva   ='" & apici(GetDiz(DizDatabase,"PartitaIva"))    & "'"
         MySql = MySql & ",CodiceFiscale='" & apici(GetDiz(DizDatabase,"CodiceFiscale")) & "'"
         MySql = MySql & ",Indirizzo1   ='" & apici(GetDiz(DizDatabase,"Indirizzo1"))    & "'"
         MySql = MySql & ",Indirizzo2   ='" & apici(GetDiz(DizDatabase,"Indirizzo2"))    & "'"
         MySql = MySql & ",Cap          ='" & apici(GetDiz(DizDatabase,"Cap"))           & "'"
         MySql = MySql & ",Comune       ='" & apici(GetDiz(DizDatabase,"Comune"))        & "'"
         MySql = MySql & ",Provincia    ='" & apici(GetDiz(DizDatabase,"Provincia"))     & "'"
         MySql = MySql & ",Settore      ='" & apici(GetDiz(DizDatabase,"Settore"))       & "'"
         MySql = MySql & ",email1       ='" & apici(GetDiz(DizDatabase,"email1"))        & "'"
         MySql = MySql & ",email2 ='"            & apici(GetDiz(DizDatabase,"email2"))        & "'"
         MySql = MySql & ",Telefono ='"          & apici(GetDiz(DizDatabase,"Telefono"))      & "'"
         MySql = MySql & ",FlagAttivo ='"        & apici(GetDiz(DizDatabase,"FlagAttivo"))    & "'"
         MySql = MySql & ",DescBlocco ='"        & apici(GetDiz(DizDatabase,"DescBlocco"))    & "'"
         MySql = MySql & ",Cognome ='"           & apici(GetDiz(DizDatabase,"Cognome"))       & "'"
         MySql = MySql & ",Nome ='"              & apici(GetDiz(DizDatabase,"Nome"))          & "'"
		 MySql = MySql & ",IdTipoUsoServizio ='" & apici(GetDiz(DizDatabase,"IdTipoUsoServizio")) & "'"
		 MySql = MySql & ",IdProfiloAbilitazione=" & GetDiz(DizDatabase,"IdProfiloAbilitazione")
		 MySql = MySql & " where IdAccount = " & v_id
		 ConnMsde.execute MySql 
		 if err.number<>0 then 
		    Esito = Err.description 
		 end if 		 
      end if 	  
   end if 
 
   UpdateAccount=Esito 
End function 
%>