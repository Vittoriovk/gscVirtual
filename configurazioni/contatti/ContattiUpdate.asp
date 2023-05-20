<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
err.clear

sendData          = Request("sendData")
StructHeader      = Request("campiStruttura")
rowData           = Request("rowData")
IdAccount         = cdbl("0" & Request("IdAccount"))

Set d = CreateObject("Scripting.Dictionary")
ArRigaC = Split(StructHeader,";")

for conta=LBound(ArRigaC) to Ubound(ArRigaC)
    x=ArRigaC(conta)
    if len(trim(x))>0 then 
      ArRigaD = split(trim(x),",")
      d.item(ArRigaD(0))=conta
   end if 
next   

Azioni=""
if len(rowData)>0 and cdbl(IdAccount)>0 then 
   DtRiga = split(rowData,"~~")
   Azioni            = DtRiga(d("Azioni"))
   IdAccountContatto = cdbl("0" & DtRiga(d("IdAccountContatto")))
   IdTipoContatto    = DtRiga(d("IdTipoContatto"))
   DescContatto      = DtRiga(d("DescContatto"))
   NoteContatto      = DtRiga(d("NoteContatto")) 
   FlagPrincipale    = DtRiga(d("ValFlagPrincipale"))
end if 
 

MyQ = "" 
if Azioni="DEL" and cdbl(IdAccountContatto) > 0  then 
   MyQ = "" 
   MyQ = MyQ & " Delete From AccountContatto "
   MyQ = MyQ & " where IdAccountContatto = " & IdAccountContatto
   MyQ = MyQ & " and   IdAccount = "         & IdAccount
end if 
if (Azioni="MOD" or Azioni="OLD") and cdbl(IdAccountContatto) > 0  then 
   MyQ = "" 
   MyQ = MyQ & " update AccountContatto set "
   MyQ = MyQ & " IdTipoContatto = '" & Apici(idTipoContatto) & "'"
   MyQ = MyQ & ",DescContatto = '"   & Apici(DescContatto) & "'"
   MyQ = MyQ & ",NoteContatto = '"   & Apici(NoteContatto) & "'"
   MyQ = MyQ & ",FlagPrincipale = '" & Apici(FlagPrincipale) & "'"			 
   MyQ = MyQ & " where IdAccountContatto = " & IdAccountContatto
   MyQ = MyQ & " and   IdAccount = "         & IdAccount
end if 		  
if Azioni="NEW" then 
   MyQ = "" 
   MyQ = MyQ & " Insert into AccountContatto ("
   MyQ = MyQ & " IdAccount,IdTipoContatto,DescContatto,NoteContatto,FlagPrincipale,Ordine"
   MyQ = MyQ & ") values ("			
   MyQ = MyQ & "  " & IdAccount
   MyQ = MyQ & ",'" & Apici(idTipoContatto) & "'"
   MyQ = MyQ & ",'" & Apici(DescContatto)   & "'"
   MyQ = MyQ & ",'" & Apici(NoteContatto)   & "'"
   MyQ = MyQ & ",'" & Apici(FlagPrincipale) & "'"
   MyQ = MyQ & ",99" 
   MyQ = MyQ & ")"
end if 	

if MyQ<>"" then 
   ConnMsde.execute MyQ
end if 
  
%>