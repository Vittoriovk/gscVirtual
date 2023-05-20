<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
err.clear
flagDebug=true
sendData          = Request("sendData")
StructHeader      = Request("campiStruttura")
rowData           = Request("rowData")
IdAccount         = cdbl("0" & Request("IdAccount"))

sendData = DecryptWithKey(sendData,Session("CryptKey"))
arD=split(sendData,"|")
Trovato=false 
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
      if k="IDACCOUNT" then 
	     if cdbl(v)>0 then 
		    Trovato = true
	     end if  
	  end if 
   end if 
next 
if trovato=false then 
   return
end if 

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

   IdAccountSede = cdbl("0" & DtRiga(d("IdAccountSede")))
   IdTipoSede    = DtRiga(d("IdTipoSede"))
   IdStato       = DtRiga(d("IdStato"))
   Indirizzo     = DtRiga(d("Indirizzo")) 
   Civico        = DtRiga(d("Civico"))
   Cap           = DtRiga(d("Cap"))
   Comune        = DtRiga(d("Comune"))
   Provincia     = DtRiga(d("Provincia"))
   if IdStato="" then 
      IdStato="IT"
   end if 
end if 

MyQ = "" 
if Azioni="DEL" and cdbl(IdAccountSede) > 0  then 
   MyQ = "" 
   MyQ = MyQ & " Delete From AccountSede "
   MyQ = MyQ & " where IdAccountSede = " & IdAccountSede
   MyQ = MyQ & " and   IdAccount = "     & IdAccount
end if 
if (Azioni="MOD" or Azioni="OLD") and cdbl(IdAccountSede) > 0  then 
   MyQ = "" 
   MyQ = MyQ & " update AccountSede set "
   MyQ = MyQ & " IdStato = '"     & Apici(idStato) & "'"
   MyQ = MyQ & ",IdTipoSede = '"  & Apici(idTipoSede) & "'"
   MyQ = MyQ & ",Indirizzo = '"   & Apici(Indirizzo) & "'"
   MyQ = MyQ & ",Civico = '"      & Apici(civico) & "'"
   MyQ = MyQ & ",Cap = '"         & Apici(Cap) & "'"
   MyQ = MyQ & ",Comune = '"      & Apici(Comune) & "'"
   MyQ = MyQ & ",Provincia = '"   & Apici(Provincia) & "'"   
   MyQ = MyQ & " where IdAccountSede = " & IdAccountSede
   MyQ = MyQ & " and   IdAccount = "     & IdAccount
end if 		  

if Azioni="NEW" then 
   MyQ = "" 
   MyQ = MyQ & " Insert into AccountSede ("
   MyQ = MyQ & " IdAccount,IdStato,IdTipoSede,Indirizzo,Civico,Cap,Comune,Provincia,Ordine"
   MyQ = MyQ & ") values ("			
   MyQ = MyQ & "  " & IdAccount
   MyQ = MyQ & ",'" & Apici(IdStato)    & "'"
   MyQ = MyQ & ",'" & Apici(idTipoSede) & "'"
   MyQ = MyQ & ",'" & Apici(Indirizzo)  & "'"
   MyQ = MyQ & ",'" & Apici(Civico)     & "'"
   MyQ = MyQ & ",'" & Apici(Cap)        & "'"
   MyQ = MyQ & ",'" & Apici(Comune)     & "'"
   MyQ = MyQ & ",'" & Apici(Provincia)  & "'"
   MyQ = MyQ & ",99" 
   MyQ = MyQ & ")"
end if 	

if MyQ<>"" then 
   ConnMsde.execute MyQ
   err.clear
end if 
 
%>