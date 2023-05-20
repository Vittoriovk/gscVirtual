<%
Function getAziendaMaster()
dim q
   q="select IdAzienda from Azienda Where IdAziendaMaster = 0"
   getAziendaMaster = cdbl("0" & LeggiCampo(q,"IdAzienda"))
end Function 

Function getInfoAzienda(idAzienda)
dim D,xx
   Set D = Server.CreateObject("Scripting.Dictionary")
   xx=GetInfoRecordset(D,"select * from Azienda Where IdAzienda = " & IdAzienda)
   set getInfoAzienda=D
end Function 

Function getInfoClienteByAccount(id)
dim D,xx
   Set D = Server.CreateObject("Scripting.Dictionary")
   xx=GetInfoRecordset(D,"select * from Cliente Where IdAccount = " & Id)
   set getInfoClienteByAccount=D
end Function 
Function getInfoClienteById(id)
dim D,xx
   Set D = Server.CreateObject("Scripting.Dictionary")
   xx=GetInfoRecordset(D,"select * from Cliente Where IdCliente = " & Id)
   set getInfoClienteById=D
end Function 

Function getStatoRichiesta(IdStato)
dim q
   q="select * from StatoRichiesta Where IdStatoRichiesta='" & apici(IdStato) & "'"
   getStatoRichiesta = LeggiCampo(q,"DescStatoRichiesta") 
end Function 


Function VerificaDel(entita,Id)
Dim retV,keyType,j,q
   retV=""
   dim tabDef(10)
   dim keyDef(10)
   dim keyTyp
   j=0
   if ucase(entita)="RAMO" then 
	  j=j+1
      tabDef(j)="Prodotto"
	  keyDef(j)="IdRamo"   
	  j=j+1
      tabDef(j)="RegolaProvvigione"
	  keyDef(j)="IdRamo"  	  
   END IF 
   if ucase(entita)="SUBRAMO" then 
	  j=j+1
      tabDef(j)="Prodotto"
	  keyDef(j)="IdSubRamo"   
   END IF    
   if ucase(entita)="COMPAGNIA" then 
	  j=j+1
      tabDef(j)="Prodotto"
	  keyDef(j)="IdCompagnia"   
   END IF    
   if ucase(entita)="FORNITORE" then 
	  j=j+1
      tabDef(j)="CauzioneDefRichiestaComp"
	  keyDef(j)="IdAccountFornitore"   
	  j=j+1
      tabDef(j)="DirittiEmissione"
	  keyDef(j)="IdAccountFornitore"  
	  j=j+1
      tabDef(j)="ProdottoAttivo"
	  keyDef(j)="IdAccountFornitore"  
	  j=j+1
      tabDef(j)="ServizioRichiesto"
	  keyDef(j)="IdAccountFornitore"  
  
   END IF 
   if ucase(entita)="FORNITORE_1" then 
	  j=j+1
      tabDef(j)="AffidamentoRichiestaComp"
	  keyDef(j)="IdFornitore"   
	  j=j+1
      tabDef(j)="CauzioneCompagnia"
	  keyDef(j)="IdFornitore" 
	  j=j+1
      tabDef(j)="Movimento"
	  keyDef(j)="IdFornitore" 
	  j=j+1
      tabDef(j)="RegolaProvvigione"
	  keyDef(j)="IdFornitore" 
   END IF
   if ucase(entita)="DOCUMENTO" then 
	  j=j+1
      tabDef(j)="ServizioDocumento"
	  keyDef(j)="IdDocumento"   
 
	  
   END IF  
   if ucase(entita)="PRODOTTOTEMPLATE" then 
	  j=j+1
      tabDef(j)="ServizioRichiesto"
	  keyDef(j)="IdProdottoTemplate"   
	  j=j+1
      tabDef(j)="Prodotto"
	  keyDef(j)="IdProdottoTemplate" 	  
	  
   END IF    
   if ucase(entita)="ELENCO" then 
      'usare apice per una stringa
      keySep=""
	  j=j+1
      tabDef(j)="DatoTecnico"
	  keyDef(j)="IdElenco"
	  j=j+1
      tabDef(j)="ElencoValore"
	  keyDef(j)="IdElenco"
   end if
   for k=1 to j
      'leggo la tabella 
	  q="select top 1 * from " & tabDef(k) & " where " & keyDef(k) & "=" & keysep & Id & keysep
      retV = LeggiCampo(q,keyDef(k))
      if retV<>"" then 
	     retV = " Cancellazione non possibile : dati associati su " & tabDef(k)
	     exit for
      end if 
   next    

   VerificaDel = retV
end Function

Function SelezionaServiziAccount(Cond)
Dim Q, IdTipoUsoServizio,Oggi
   Q = ""
   Oggi = DtoS()

   IdTipoUsoServizio=LeggiCampo("Select * from Account Where IdAccount=" & Session("LoginIdAccount"),"IdTipoUsoServizio")
		
   if IdTipoUsoServizio="TUTTI" then 
	   Q = "" 
	   Q = Q & " Select * From Servizio B "
	   Q = Q & " Where  B.IdServizio <> '' "
	   Q = Q & " And    B.DataInizioValidita <= " & Oggi
	   Q = Q & " And    B.DataFineValidita   >= " & Oggi	   
	   Q = Q & Cond 
	   Q = Q & " order By DescServizio"   
   end if 
   
   if IdTipoUsoServizio="TRANNE" then 
   	   Q = "" 
	   Q = Q & " Select * From Servizio B "
	   Q = Q & " Where  B.IdServizio not in ("
	   Q = Q & " select IdServizio From AccountServizio where IdAccount = " & Session("LoginIdAccount")
	   Q = Q & ") "
	   Q = Q & " And    B.DataInizioValidita <= " & Oggi
	   Q = Q & " And    B.DataFineValidita   >= " & Oggi	   
	   Q = Q & Cond 
	   Q = Q & " order By DescServizio"
   end if 
   
   if IdTipoUsoServizio="SOLO" then
	   Q = "" 
	   Q = Q & " Select * From AccountServizio A, Servizio B "
	   Q = Q & " Where  A.IdAccount = " & Session("LoginIdAccount")
	   Q = Q & " And    A.IdServizio = B.IdServizio "
	   Q = Q & " And    B.DataInizioValidita <= " & Oggi
	   Q = Q & " And    B.DataFineValidita   >= " & Oggi
       Q = Q & " And    A.ValidoDal  <= " & Oggi
	   Q = Q & " And    A.ValidoAl   >= " & Oggi	   
	   Q = Q & Cond 
	   Q = Q & " order By DescServizio"
   END IF 

   if IdTipoUsoServizio="NESSUNO" then
	   Q = "" 
	   Q = Q & " Select * From Servizio B "
	   Q = Q & " Where  B.IdServizio = '' "
	   Q = Q & " And    B.DataInizioValidita <= " & Oggi
	   Q = Q & " And    B.DataFineValidita   >= " & Oggi	   
	   Q = Q & Cond 
	   Q = Q & " order By DescServizio"
   END IF 
   
   SelezionaServiziAccount=Q		
end Function 

function addLog(descErr,descQuery)
on error resume next 
connmsde.execute "insert into log(descErrore,query) values ('" &  apici(mid(descErr,1,450)) & "','" & apici(mid(descQuery,1,950)) & "') "
err.clear 
		
end Function 

function addAudit(IdTabella,KeyTabella,IdAccount,DescAudit)
Dim qIns,lAccount
on error resume next 
lAccount=idAccount
if cdbl(idAccount)=0 then 
   lAccount=Session("LoginIdAccount")
end if 
qIns=""
qIns = qIns & " Insert into Audit (IdTabella,IdTabellaKeyString,DataAudit,TimeAudit,IdAccount,DescAudit) "
qIns = qIns & "values ("
qIns = qIns & " '" & Apici(IdTabella)  & "'"
qIns = qIns & ",'" & Apici(KeyTabella) & "'"
qIns = qIns & ", " & Dtos()
qIns = qIns & ", " & TimeToS()
qIns = qIns & ", " & lAccount
qIns = qIns & ",'" & Apici(DescAudit) & "'"
qIns = qIns & ")"
'response.write qIns
connmsde.execute qIns
err.clear 
end Function 


Function GetOperAbilitate(IdProfilo,IdAnagFunzione)
Dim MyQ,MyRs,retVal 
   MyQ = ""
   MyQ = MyQ & " select * "
   MyQ = MyQ & " from  ProfiloAbilitazioneFunzione "
   MyQ = MyQ & " where IdProfiloAbilitazione = " & IdProfilo 
   MyQ = MyQ & " and IdAnagFunzione='" & apici(IdAnagFunzione) & "'"
   retVal=""
   
   on Error resume next 
   Set MyRs = Server.CreateObject("ADODB.Recordset")
   MyRs.CursorLocation = 3
   MyRs.Open MyQ, ConnMsde 

   if MyRs.eof = false then 
      if MyRs("FlagCreate")=1 then 
         retVal=retVal & "C"
      end if 
      if MyRs("FlagRead")=1 then 
         retVal=retVal & "R"
      end if 
      if MyRs("FlagUpdate")=1 then 
         retVal=retVal & "U"
      end if 
      if MyRs("FlagDelete")=1 then 
         retVal=retVal & "D"
      end if 
   end if 
   MyRs.close
   err.clear 
   
   GetOperAbilitate=retVal
end Function 		

Function GetSubAziende(IdAzienda)
Dim MyQ,MyRs,retVal,ArDat,maxI,curI,idA

   Set MyRs = Server.CreateObject("ADODB.Recordset")
   MyRs.CursorLocation = 3
   
   maxI=1
   redim ArDat(maxI+1)
   ArDat(maxI)=IdAzienda 
   curI=1
   retVal=IdAzienda
   
   on Error resume next 
   do while curI<=maxI
      idA=ArDat(curI)
	  
      MyQ = "select * from Azienda Where IdAziendaMaster=" & IdA 
	  'response.write MyQ
      MyRs.Open MyQ, ConnMsde 
	  if err.number = 0 then 
	     response.write MyRs.eof
	     do while not MyRs.eof 
		    if MyRs("IdAzienda")>0 then 
			   maxI=maxI+1
               redim preserve ArDat(maxI+1)
               ArDat(maxI)=MyRs("IdAzienda") 
			   retVal=RetVal & "," & MyRs("IdAzienda") 
			end if
            MyRs.moveNext 			
         loop
         MyRs.close 
      end if 
      err.clear 
      curI=curI+1
   loop
   GetSubAziende=retVal
   'response.write retVal
   'response.end 
end Function 		

%>