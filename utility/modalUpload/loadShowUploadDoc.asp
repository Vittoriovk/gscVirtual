<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
IdModUpload    = Request("IdModUpload")
Linkdocumento  = Request("Linkdocumento")
DescFileUpload = Request("DescDocumento")
NomeFileUpload = Request("nomeDocumento")
canUpload      = Request("canUpload") 
idTableUpload  = Request("idTableUpload")
if canUpload   = "N" then 
   canUpload = false 
else
   canUpload = true
end if 
InsertRow      = false 
IsFileUploadMd = "S"
%>
<!--#include virtual="/gscVirtual/utility/modalUpload/showUploadDoc.asp"-->
