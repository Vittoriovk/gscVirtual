<!--#include virtual="/gscVirtual/common/FunCallOtherPage.asp"-->
<%
Function GetDizCauzione(IdCauzione,IdAccountCliente)
  
   Dim MyRs,MySql,IdStato,DescStato,K,nome
   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
   
   Set DizDatabase = GetDizServizioRichiesto(0,"CAUZ_PROV",IdCauzione)
   if Cdbl(IdCauzione)=0 then 
      xx=SetDiz(DizDatabase,"N_IdAccountCliente",IdAccountCliente)
   end if 
   
   xx=SetDiz(DizDatabase,"N_IdCauzione",0)
   xx=SetDiz(DizDatabase,"N_DataCessazione",0)
   xx=SetDiz(DizDatabase,"S_DescCauzione","")
   xx=SetDiz(DizDatabase,"N_NumCoobbligati",0)
   xx=SetDiz(DizDatabase,"S_DescCoobbligati","")
   xx=SetDiz(DizDatabase,"S_CIG","")
   xx=SetDiz(DizDatabase,"S_CUG","")
   xx=SetDiz(DizDatabase,"S_CPV","")
   xx=SetDiz(DizDatabase,"S_Beneficiario","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioCF","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioPI","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioSede","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioIndirizzo","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioProvincia","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioCap","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioPec","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioContatti","")
   xx=SetDiz(DizDatabase,"S_OggettoAppalto","")
   xx=SetDiz(DizDatabase,"S_TipologiaAppalto","")
   xx=SetDiz(DizDatabase,"S_Categoria","")
   xx=SetDiz(DizDatabase,"S_Classe","")
   xx=SetDiz(DizDatabase,"S_Responsabile","")
   xx=SetDiz(DizDatabase,"S_Settore","")
   xx=SetDiz(DizDatabase,"S_ModalitaRealizzazione","")
   xx=SetDiz(DizDatabase,"S_LuogoEsecuzione","")
   xx=SetDiz(DizDatabase,"S_Procedura","")
   xx=SetDiz(DizDatabase,"S_CriterioAggiudicazione","")
   xx=SetDiz(DizDatabase,"N_ImportoLotto",0)
   xx=SetDiz(DizDatabase,"N_ImportoSicurezza",0)
   xx=SetDiz(DizDatabase,"N_PercGarantita",0)
   xx=SetDiz(DizDatabase,"N_ImptGarantito",0)
   xx=SetDiz(DizDatabase,"N_DataPubblicazione",0)
   xx=SetDiz(DizDatabase,"N_DataScadenza",0)
   xx=SetDiz(DizDatabase,"N_DataApertura",0)
   xx=SetDiz(DizDatabase,"S_Indirizzo","")
   xx=SetDiz(DizDatabase,"S_Cap","")
   xx=SetDiz(DizDatabase,"S_Comune","")
   xx=SetDiz(DizDatabase,"S_Provincia","")
   xx=SetDiz(DizDatabase,"S_DescStato","")
   xx=SetDiz(DizDatabase,"S_Civico","")
   xx=SetDiz(DizDatabase,"S_Pec","")
   xx=SetDiz(DizDatabase,"S_IdTipoFirma","")
   xx=SetDiz(DizDatabase,"N_ImptFirma",0)
   xx=SetDiz(DizDatabase,"S_ElencoATI","")
   xx=SetDiz(DizDatabase,"N_NumATI",0)
   xx=SetDiz(DizDatabase,"N_CapogruppoATI",0)
   xx=SetDiz(DizDatabase,"S_DescATI","")
   xx=SetDiz(DizDatabase,"S_ListaCertificazioni","")
   xx=SetDiz(DizDatabase,"N_NumLotti",0)
   xx=SetDiz(DizDatabase,"S_DescLotti","")
   xx=SetDiz(DizDatabase,"S_DescCasoParticolare","")
   xx=SetDiz(DizDatabase,"N_RichiediCond",0)
   
   Set MyRs = Server.CreateObject("ADODB.Recordset")
   if Cdbl(IdCauzione)>0 then 
      MySql = "select * from Cauzione Where IdCauzione=" & IdCauzione  
      MyRs.CursorLocation = 3 
      MyRs.Open MySql, ConnMsde      
      if MyRs.eof = false then 
         For Each K In DizDatabase
            if isOfServizioRichiesto(k)=false then 
               nome = mid(k,3,99)
               xx=SetDiz(DizDatabase,k,MyRs(nome)) 
            end if 
         next 
      end if 
      MyRs.close 
   end if 
   if cdbl(IdCauzione)=0 and GetDiz(DizDatabase,"")="" then 
      xx=SetDiz(DizDatabase,"S_DescCauzione"    ,"--Nuova Cauzione--" )
      xx=SetDiz(DizDatabase,"S_IdStatoServizio" ,"COMP" )
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
   
   if Cdbl(IdCauzione)=0 and GetDiz(DizDatabase,"S_mailNotificaCliente")="" and GetDiz(DizDatabase,"N_IdAccountCliente")<>"0" then 
      mail=getMailForAccount(GetDiz(DizDatabase,"N_IdAccountCliente"),"CLIE","",0)
      xx=SetDiz(DizDatabase,"S_mailNotificaCliente"   ,mail)
   end if 
   if Cdbl(IdCauzione)=0 and GetDiz(DizDatabase,"S_mailNotificaRichiedente")="" then 
      mail=getMailForAccount(Session("LoginIdAccount"),Session("LoginTipoUtente"),"",0)
      xx=SetDiz(DizDatabase,"S_mailNotificaRichiedente"   ,mail)
   end if 
   set GetDizCauzione = DizDatabase
End function 

Function GetNewCauzione(DizDatabase)
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
         if nome<>ucase("IdCauzione") and trim(nome)<>"" then 
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
   MySql = MySql & " Insert into Cauzione (" & Campi &  ") Values (" & Valori & ")"
   'response.write MySql 
   
   ConnMsde.execute MySql 
   if err.number=0 then 
      v_id=GetTableIdentity("Cauzione") 
      xx=SetDiz(DizDatabase,"N_IdCauzione",v_id)
      xx=SetDiz(DizDatabase,"N_IdNumAttivita",v_id)
      xx=GetNewServizioRichiesto(DizDatabase)  
   else 
      xx=writeTraceAttivita("functionCauzione:GetNewCauzione:" & MySql & err.description,"CAUZ_PROV",0)   
   end if 
   GetNewCauzione=v_id
end function 


Function UpdateCauzione(DizDatabase)
Dim v_id , MySql , Esito , v_ret,Nome,IdAz,IdCauzione
on error resume next 
   Esito = ""
  MySql = ""
   
   'leggo tutti i parametri ad accezione di IdCauzione
   IdCauzione=0
   For Each K In DizDatabase
      if isOfServizioRichiesto(k)=false then
        'response.write 
         Valo = Pulisci(DizDatabase.item(ucase(K)))
         Tipo = mid(k,1,2)
         nome = mid(k,3,99)
         'response.write nome & "=" & valo & " " & K & "<br>"
		 if nome = ucase("OggettoAppalto") then 
		    Valo = mid(Valo,1,990)
		 end if 
		 if nome = ucase("DescCauzione") then 
		    Valo = mid(Valo,1,990)
		 end if 
         if nome<>ucase("IdCauzione") and trim(nome)<>"" then 
            if MySql <> "" then 
               MySql = MySql & ","
            end if 
            MySql = MySql & nome & "="
            if Tipo="N_" then 
               MySql = MySql & NumForDb(valo)
            else
               MySql = MySql & "'" & apici(valo) & "'"
            end if            
         elseif nome=ucase("IdCauzione") then 
            IdCauzione=Cdbl(Valo)
         end if 
      end if   
   Next
   MySql = " Update Cauzione Set " & MySql 
   MySql = MySql & " Where IdCauzione = " & IdCauzione
   ConnMsde.Execute MySql 
   if err.Number <> 0 then 
      xx=writeTraceAttivita("functionCauzione:UpdateCauzione:" & MySql & err.description,"CAUZ_PROV",IdCauzione)
   end if 
   
   xx=SetDiz(DizDatabase,"S_DescServizioRichiesto",GetDiz(DizDatabase,"S_DescCauzione"))
   xx=UpdateServizioRichiesto(DizDatabase)
   
   UpdateCauzione=Esito 
End function 

Function CancellaCauzione(idC,idAccount)
Dim MsgErrore,qDel,recordsAffected
MsgErrore=""
err.clear
qDel = ""
qDel = qDel & " Delete From Cauzione Where IdCauzione=" & idC
qDel = qDel & " And   IdStatoServizio in ('COMP','PAGA')"
if Cdbl(IdAccount)>0 then 
   qDel = qDel & "and IdAccountCliente = " & NumForDb(IdAccount)
end if 
recordsAffected=0
connMsde.execute qDel , recordsAffected
if recordsAffected=1 then 
   connMsde.execute "Delete From CauzioneATI Where IdCauzione=" & idC
   connMsde.execute "Delete From CauzioneCGI Where IdCauzione=" & idC
   connMsde.execute "Delete From CauzioneCoobbligato Where IdCauzione=" & idC
else 
   MsgErrore = "cauzione non modificabile"
end if 
CancellaCauzione=MsgErrore
end function 

Function CalcolaPercGarantita(ListaCertificazioni)
Dim RetVal,Lista,PercRid 
Dim MyRs,MySql
   RetVal = 2
   'ciclo sulle certificazioni 
   Lista = trim(ListaCertificazioni)
   if Lista<>"" then 
      if Mid(lista,1,1)="|" then 
         Lista = trim(Mid(Lista,2,99))
      end if 
      if right(lista,1)="|" then 
         Lista = trim(mid(Lista,1,len(Lista)-1))
      end if 
      Lista = replace(Lista,"|",",")
      Set MyRs = Server.CreateObject("ADODB.Recordset")
 
      MySql = "select * from Certificazione Where IdCertificazione in (" & Lista & ")"
      'response.write MySql 
      MyRs.CursorLocation = 3 
      MyRs.Open MySql, ConnMsde      
      if Err.number = 0 then 
         do While not MyRs.eof
            PercRid = MyRs("PercRiduzioneCauzione")
            'response.write "PercR:" & PercRid  
            RetVal = cdbl(RetVal) * (1 - cdbl(PercRid)/100)
            MyRs.Movenext 
         loop 
      end if 
      MyRs.close 

   end if 
   'response.write "dd:" & RetVal 
   CalcolaPercGarantita = RetVal
end Function 

Function generaRiepilogo(IdCauzione) 
on error resume next 
Dim opUrl,opMethod,opData,opType,OpResp
   opUrl     = getDomain() & virtualPath & "/pdf/PdfCauzioneProvvisoriaBozza.asp"
   opMethod  = "POST"
   opData    = ""  
   opData    = opData & "IdCauzione=" & IdCauzione 
   opType    = ""
   opReferer = "http://www.mysite.com"
   opResp    = ""
   xx = CallOtherPage(opUrl,opMethod,opData,opType,opReferer,opResp)	
   'response.write opUrl & "  xx=" & xx & err.description
   'response.end  
   err.clear 
   
end function 

%>