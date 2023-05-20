<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<%
' Recupero il file da scaricare
Dim download, file
linkDocumento = Request("tt")

pathFile = Server.MapPath(Virtualpath)

file = pathFile & DirectoryUpload & linkDocumento

'controllo se esiste ed è pieno 
filesize=0
Set fso = Server.CreateObject("Scripting.FileSystemObject")
if fso.fileExists(file) then 
   Set fileObject = fso.GetFile(file)
   filesize = fileObject.Size
   Set fileObject = Nothing
end if 
Set fso = Nothing

if filesize > 0 then 
   Set download = Server.CreateObject("ADODB.Stream")
' Apro la connessione e carico il file
   download.Type = 1
   download.Open
   download.LoadFromFile file
   allData = download.Read
end if 
' Aggiungo le intestazioni del tipo di file
Response.AddHeader "Content-Disposition", "attachment; filename=" & linkDocumento
Response.ContentType = "application/octet-stream"
if filesize > 0 then
   Response.BinaryWrite allData
   download.Close
end if 
Set download = Nothing

%>
