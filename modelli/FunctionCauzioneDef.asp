<%

Function GetDizCauzioneDef(IdCauzione,IdAccountCliente)
   Dim MyRs,MySql,IdStato,DescStato,K,nome
   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
   
   Set DizDatabase = GetDizServizioRichiesto(0,"CAUZ_DEFI",IdCauzione)
   if Cdbl(IdCauzione)=0 then 
      xx=SetDiz(DizDatabase,"N_IdAccountCliente",IdAccountCliente)
   end if 
   xx=SetDiz(DizDatabase,"N_IdCauzioneDef",0)
   xx=SetDiz(DizDatabase,"S_Indirizzo","")
   xx=SetDiz(DizDatabase,"S_Cap","")
   xx=SetDiz(DizDatabase,"S_Comune","")
   xx=SetDiz(DizDatabase,"S_Provincia","")
   xx=SetDiz(DizDatabase,"S_DescStato","")
   xx=SetDiz(DizDatabase,"S_Civico","")
   xx=SetDiz(DizDatabase,"S_Pec","")   
   xx=SetDiz(DizDatabase,"S_DescCauzione","")
   xx=SetDiz(DizDatabase,"S_Beneficiario","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioCF","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioPI","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioSede","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioIndirizzo","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioProvincia","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioCap","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioPec","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioContatti","")
   xx=SetDiz(DizDatabase,"S_DescCoobbligati","")  
   xx=SetDiz(DizDatabase,"S_DescATI","")
   xx=SetDiz(DizDatabase,"S_OggettoAppalto","")
   xx=SetDiz(DizDatabase,"N_ImportoLotto",0)
   xx=SetDiz(DizDatabase,"N_ImportoSicurezza",0)
   xx=SetDiz(DizDatabase,"S_PathDocumentoZip","")
   xx=SetDiz(DizDatabase,"S_PathDocumentoRichiesta","")

   Set MyRs = Server.CreateObject("ADODB.Recordset")
   if Cdbl(IdCauzione)>0 then 
      MySql = "select * from CauzioneDef Where IdCauzioneDef=" & IdCauzione  
      MyRs.CursorLocation = 3 
      MyRs.Open MySql, ConnMsde      
      if MyRs.eof = false then 
         For Each K In DizDatabase
            if isOfServizioRichiesto(k)=false then 
               nome = mid(k,3,99)			
               xx=SetDiz(DizDatabase,k,MyRs(nome))
            end if 
         Next
	  end if 
	  MyRs.close 
   end if 
   
   if Cdbl(IdCauzione)=0 and GetDiz(DizDatabase,"S_Pec")="" and GetDiz(DizDatabase,"N_IdAccountCliente")<>"0" then 
	  MySql = ""
      MySql = MySql & " select * from AccountContatto "
	  MySql = MySql & " Where IdAccount=" & GetDiz(DizDatabase,"N_IdAccountCliente") 
	  MySql = MySql & " And IdTipoContatto='PECC' order by FlagPrincipale Desc"  
	  
      MyRs.CursorLocation = 3 
      MyRs.Open MySql, ConnMsde 
      if MyRs.eof = false then 
	     'response.write MyRs("DescContatto")
         xx=SetDiz(DizDatabase,"S_Pec",MyRs("DescContatto"))
	  end if 
	  MyRs.close    
   end if
   
   if Cdbl(IdCauzione)=0 and GetDiz(DizDatabase,"S_mailNotificaCliente")="" and GetDiz(DizDatabase,"N_IdAccountCliente")<>"0" then 
      mail=getMailForAccount(GetDiz(DizDatabase,"N_IdAccountCliente"),"CLIE","",0)
	  xx=SetDiz(DizDatabase,"S_mailNotificaCliente"   ,mail)
   end if 
   if Cdbl(IdCauzione)=0 and GetDiz(DizDatabase,"S_mailNotificaRichiedente")="" then 
      mail=getMailForAccount(Session("LoginIdAccount"),Session("LoginTipoUtente"),"",0)
	  xx=SetDiz(DizDatabase,"S_mailNotificaRichiedente"   ,mail)
   end if 
      
   if Cdbl(IdCauzione)=0 and GetDiz(DizDatabase,"S_Indirizzo")="" and GetDiz(DizDatabase,"N_IdAccountCliente")<>"0" then 
      
      MySql = "select * from AccountSede Where IdAccount=" & GetDiz(DizDatabase,"N_IdAccountCliente")  
	  'response.write MySql
      MyRs.CursorLocation = 3 
      MyRs.Open MySql, ConnMsde 
      if MyRs.eof = false then 
         xx=SetDiz(DizDatabase,"S_Indirizzo",MyRs("Indirizzo"))
         xx=SetDiz(DizDatabase,"S_Cap"      ,MyRs("CAP"))
         xx=SetDiz(DizDatabase,"S_Comune"   ,MyRs("Comune"))
         xx=SetDiz(DizDatabase,"S_Provincia",MyRs("Provincia"))
         xx=SetDiz(DizDatabase,"S_Civico"   ,MyRs("Civico"))

         IdStato  = MyRs("IdStato")
         if IdStato="" then 
            IdStato="IT"
         end if 
         DescStato  = LeggiCampo("Select * from Stato Where IdStato='" & IdStato & "'","DescStato")
         if DescStato="" then 
            DescStato=IdStato
         end if 
         xx=SetDiz(DizDatabase,"S_DescStato",DescStato)
	  end if 
	  MyRs.close 
   end if   
   
   set GetDizCauzioneDef = DizDatabase
End function 

Function GetNewCauzioneDef(DizDatabase)
Dim MySql,v_id,Campi,Valori

   v_id=0
   MySql = ""
   Campi = ""
   Valori= ""
   'leggo tutti i parametri ad accezione di IdCauzione
   For Each K In DizDatabase
      if isOfServizioRichiesto(k)=false then 
         Valo = GetDiz(DizDatabase,K)
         Tipo = mid(k,1,2)
         nome = mid(k,3,99)
         'response.write nome 
	     if nome<>ucase("IdCauzioneDef") then 
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
   MySql = MySql & " Insert into CauzioneDef (" & Campi &  ") Values (" & Valori & ")"
   
   ConnMsde.execute MySql 
   if err.number=0 then 
      v_id=GetTableIdentity("CauzioneDef") 
	  xx=SetDiz(DizDatabase,"N_IdCauzioneDef",v_id)
	  xx=SetDiz(DizDatabase,"N_IdNumAttivita",v_id)
      xx=GetNewServizioRichiesto(DizDatabase)
   else
      xx=writeTraceAttivita("Errore  GetNewCauzioneDef :" & Sql & ":" & Err.description,GetDiz(DizDatabase,"S_IdAttivita"),GetDiz(DizDatabase,"N_IdNumAttivita"))   
   end if 
   GetNewCauzioneDef=v_id
end function 


Function UpdateCauzioneDef(DizDatabase)
Dim v_id , MySql , Esito , v_ret,Nome,IdAz,IdCauzione,IdAttivita
on error resume next 
   Esito = ""
  MySql = ""
   
   'leggo tutti i parametri ad accezione di IdCauzione
   IdCauzione = 0
   IdAttivita = ""
   For Each K In DizDatabase
      if isOfServizioRichiesto(k)=false then 
        'response.write 
        Valo = DizDatabase.item(ucase(K))
        Tipo = mid(k,1,2)
		nome = mid(k,3,99)
		'response.write nome & "=" & valo & " " & K
		if nome=ucase("IdAttivita") then 
		   IdAttivita=valo
		end if 
		
		if nome<>ucase("IdCauzioneDef") then 
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
		   IdCauzione=Cdbl(Valo)
		end if 
	  end if 
   Next
   MySql = " Update CauzioneDef Set " & MySql 
   MySql = MySql & " Where IdCauzioneDef = " & IdCauzione
   ConnMsde.Execute MySql 
   if err.number <> 0 then 
      xx = writeTraceAttivita("UpdateCauzioneDef:" & MySql & " " & Err.description,IdAttivita,IdCauzione)
   end if 
   
   'response.write MySql 
   
   xx=SetDiz(DizDatabase,"S_DescServizioRichiesto",GetDiz(DizDatabase,"S_DescCauzione"))
   xx=UpdateServizioRichiesto(DizDatabase)
   UpdateCauzioneDef=Esito 
End function 

%>