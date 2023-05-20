<%@language=vbscript%>

<% 
' response.ContentType ="application/pdf" 
' xx=Response.AddHeader("Content-Disposition","inline")

%> 
<!--#include virtual="/gscVirtual/common/function.asp"-->
<!--#include virtual="/gscVirtual/common/functionNew.asp"-->
<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<!--#include file="fpdf.asp"-->
<%

'verifico se il file esiste nel caso lo leggo e lo invio al browser
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")

nome="BaseTemplate.pdf"
filename=Server.MapPath(VirtualPath & nome )

' Modifiche per Android
NomePdfOut = nome
PathC=Server.MapPath(VirtualPath)

response.write filename
'response.end 

If ScriptObject.FileExists(filename) = true Then
   ScriptObject.DeleteFile(filename) 
End If

Set pdf=CreateJsObject("FPDF")
pdf.CreatePDF()
pdf.SetPath("fpdf/")
pdf.SetFont "Arial","",8
pdf.Open()
pdf.AddPage()

MaxRow   =270
MaxCol   =195
BottomRow=280

for x = 0 to MaxCol step 10 
    pdf.Text x, 5, x
next 

for y = 10 to BottomRow step 5
   pdf.SetXY  0 , y  
   pdf.Write  4 , y
   for x = 5 to 195 step 5 
       pdf.SetXY  x , y  
       pdf.Write  4 , "."
   next 
next 

' NomePdfOut = "prova.pdf"
pdf.Output (filename), false
response.redirect replace(virtualpath & NomePdfOut,"\","/")
response.end

pdf.Close()
pdf.Output(filename)
pdf.Output()
  
%> 