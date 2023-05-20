<!--#include virtual="/gscVirtual/include/includeStd.asp"-->

<%
   StructBase        = Request("ns")
   StructHeader      = Request("ss")
   ArRigaD           = Request("vs")

   IdAccount=0
   struttura=""
   if mid(StructBase,1,len("CONTATTI"))="CONTATTI" then 
      struttura="CONTATTO"
   end if 

   if len(ArRigaD)=0 or len(struttura)=0 then 
      response.end 
   end if 
   
   ' ho una stringa in ingresso per una struttura nota  
   'elenco dei campi in prima posizione e li metto in un dizionario 
   Set d = CreateObject("Scripting.Dictionary")
   ArRigaC = Split(StructHeader,";")
   for conta=LBound(ArRigaC) to Ubound(ArRigaC)
      x=ArRigaC(conta)
      if len(trim(x))>0 then 
         tmpRiga = split(trim(x),",")
         d.item(tmpRiga(0))=conta
      end if 
   next   
   Azioni=""
   'indica che devo restituire la nuova riga 
   FlagRetUpd=false
   newRow=0

   
   'gestisco il contatto 
   if struttura="CONTATTO" then 
      Ordine = 99
      DtRiga = split(ArRigaD,"~~")
      Azioni            = DtRiga(d("Azioni"))
      IdAccount         = cdbl("0" & DtRiga(d("IdAccount")))
      IdAccountContatto = cdbl("0" & DtRiga(d("IdAccountContatto")))
      IdTipoContatto    = DtRiga(d("IdTipoContatto"))
      DescContatto      = DtRiga(d("DescContatto"))
      NoteContatto      = DtRiga(d("NoteContatto")) 
      FlagPrincipale    = DtRiga(d("ValFlagPrincipale"))

      MyQ = "" 
      if Azioni="DEL" and cdbl(IdAccountContatto) > 0  then 
         MyQ = "" 
         MyQ = MyQ & " Delete From AccountContatto "
         MyQ = MyQ & " where IdAccountContatto = " & IdAccountContatto
         MyQ = MyQ & " and   IdAccount = "         & IdAccount
      end if 
      if Azioni="MOD" and cdbl(IdAccountContatto) > 0  then 
         MyQ = "" 
         MyQ = MyQ & " update AccountContatto set "
         MyQ = MyQ & " IdTipoContatto = '" & Apici(idTipoContatto) & "'"
         MyQ = MyQ & ",DescContatto = '"   & Apici(DescContatto) & "'"
         MyQ = MyQ & ",NoteContatto = '"   & Apici(NoteContatto) & "'"
         MyQ = MyQ & ",FlagPrincipale = '" & Apici(FlagPrincipale) & "'"			 
         MyQ = MyQ & " where IdAccountContatto = " & IdAccountContatto
         MyQ = MyQ & " and   IdAccount = "         & IdAccount
      end if 		  
      if Azioni="NEW" and cdbl(IdAccount) > 0  then 
         MyQ = "" 
         MyQ = MyQ & " Insert into AccountContatto ("
         MyQ = MyQ & " IdAccount,IdTipoContatto,DescContatto,NoteContatto,FlagPrincipale,Ordine"
         MyQ = MyQ & ") values ("			
         MyQ = MyQ & "  " & IdAccount
         MyQ = MyQ & ",'" & Apici(idTipoContatto) & "'"
         MyQ = MyQ & ",'" & Apici(DescContatto)   & "'"
         MyQ = MyQ & ",'" & Apici(NoteContatto)   & "'"
         MyQ = MyQ & ",'" & Apici(FlagPrincipale) & "'"
         MyQ = MyQ & ", " & Ordine
         MyQ = MyQ & ")"
		 response.write MyQ
		 ConnMsde.execute MyQ
		 if err.number = 0 then 
            newRow = GetTableIdentity("AccountContatto")
		    DtRiga(d("IdAccountContatto"))=newRow
		    DtRiga(d("Azioni"))="OLD"			
            FlagRetUpd=true 
		 end if 
		 MyQ = ""
      end if 
      if MyQ<>"" then 
        'response.write MyQ	
         ConnMsde.execute MyQ
      end if 
   end if 
   
   'ricostruisco la riga e la restituisco 
   if FlagRetUpd=true then
      ArRigaD=""
      for conta=LBound(DtRiga) to Ubound(DtRiga)-1
         ArRigaD=ArRigaD + DtRiga(conta) + "~~"
      next 
      response.write ArRigaD
   end if    
%>