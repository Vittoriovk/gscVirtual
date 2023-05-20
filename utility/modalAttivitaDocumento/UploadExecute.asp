<!--#include virtual="/gscVirtual/common/clsupload.asp"-->
<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<%
   set o = new clsUpload
   
   IdAttivitaDocumento = o.valueOf("MAD_IdAttivitaDocumento")
   IdAttivitaDocumento = Cdbl("0" & IdAttivitaDocumento)
   if Cdbl(IdAttivitaDocumento)=0 then 
      response.write "ERR: id non valorizzato"
      response.end 
   end if

   NomeFilFull = o.FileNameOf("MAD_FileToUpload0")
   'file obbligatorio per caricare 
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
  
   o.FileInputName = "MAD_FileToUpload0"
   o.FileFullPath = PathBaseUpload  & sFileWrite
   o.save

   if o.Error <> ""  then
       MsgErrore= "ERR:Caricamento Fallito: " & o.Error & o.FileFullPath
   elseif err.number<>0 then
       MsgErrore= "ERR:Caricamento Errore : " & Err.Description
   else
       MSGErrore= "OK:" & sFileWrite
   end if 
   response.write msgErrore
%>
