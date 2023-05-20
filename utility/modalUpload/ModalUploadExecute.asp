<!--#include virtual="/gscVirtual/common/clsupload.asp"-->
<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<%
set o = new clsUpload
idUp = o.valueOf("idUp")
NomeFilFull = o.FileNameOf("modalUploadFile")
if trim(NomeFilFull)="" then 
   response.write "ERR: file non valorizzato"
   response.end 
end if 
'response.write "files:" & NomeFilFull
'response.end 

'response.write "title:" & o.valueOf("DescDocumento" & idUp)

sFileSplit = split(NomeFilFull, "\")
sFile = sFileSplit(Ubound(sFileSplit))

sFileWrite = "CX" & IdUp & "_" & Year(Now()) & Month(Now()) & Day(Now()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now()) &  "_" & sFile
  
o.FileInputName = "modalUploadFile"
o.FileFullPath = PathBaseUpload  & sFileWrite
o.save

   if o.Error <> ""  then
       MsgErrore= "ERR:Caricamento Fallito: " & o.Error & o.FileFullPath
   elseif err.number<>0 then
       MsgErrore= "ERR:Caricamento Errore : " & Err.Description
   else
       MSGErrore= PathBaseUpload  & sFileWrite
   end if 
   response.write msgErrore
%>
