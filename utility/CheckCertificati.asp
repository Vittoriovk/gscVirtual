<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
flagDebug=true
Oper = ""
IdCert = cdbl("0" & Request("ce"))
action = ucase(Request("action"))
allSel = Request("sel")
retVal = allsel 
if Cdbl(IdCert)=0 or instr("AR",action)=0 then 
   response.write retVal
   response.end 
end if 

'nessuna selezione : restituisco Id passato
if allSel="" then 
   if action="A" then 
      retVal="|" & IdCert & "|"
      response.write retVal
      response.end   
   end if 
end if 

idxCert=0
dim ArCert(100)
for j=1 to 99
    ArCert(j)=0
next 

if action="A" then 
   idxCert=1
   ArCert(idxCert)=IdCert
end if 

Lista=Split(allSel,"|")
for each x in Lista 
    elem=trim(x)
    if elem<>"" then
	   if action="R" then 
	      if cdbl(elem)<>cdbl(IdCert) then 
	         idxCert=idxCert+1
	         ArCert(idxCert)=cdbl(elem)
          end if 
       else 
          if Cdbl(elem)<Cdbl(IdCert) then
	         ArCert(idxCert)=cdbl(elem)
             idxCert=idxCert+1
             ArCert(idxCert)=IdCert
          else 
	         idxCert=idxCert+1
		     ArCert(idxCert)=cdbl(elem)
          end if
       end if 
    end if 
next

'restituisco la lista 
retVal="" 
for j=1 to idxCert
   elem=ArCert(j)
   if RetVal="" then 
      RetVal=RetVal & "|"
   end if 
   RetVal=RetVal & elem & "|"
next 
if action="A" then 
   test=LeggiCampo("select * from CertificazioneMatrice where CertificazioneCompatibile='" & RetVal & "'","IdCertificazioneMatrice")
   test=Cdbl("0" & Test)
   if test=0 then 
      retVal="|" & IdCert & "|"
   end if 
end if 

response.write RetVal 
Response.end 

%>