<%@language=vbscript%>

<% 
' response.ContentType ="application/pdf" 
' xx=Response.AddHeader("Content-Disposition","inline")

%> 
<!--#include virtual="/gscVirtual/common/function.asp"-->
<!--#include virtual="/gscVirtual/common/functionNew.asp"-->
<!--#include virtual="/gscVirtual/common/connDb.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAccessoDb.asp"-->
<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<!--#include virtual="/gscVirtual/common/NumToLet.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionServizioRichiesto.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionCauzione.asp"-->
<!--#include virtual="/gscVirtual/modelli/functionOpzioni.asp"-->
<!--#include file="fpdf.asp"-->
<!--#include file="writebox.asp"-->
<%

Set Rs = Server.CreateObject("ADODB.Recordset")

IdCauzione  = "0" & Request("IdCauzione")
SendBrowser = Request("SendBrowser")

if IsNumeric(IdCauzione)=false then 
   IdCauzione=0
else
   IdCauzione=cdbl(IdCauzione)
end if 

if Cdbl(IdCauzione)=0 then 
   response.end 
end if 
'verifico se devo ricreare il pdf 

FlagRewrite=1

'verifico se il file esiste nel caso lo leggo e lo invio al browser
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")

flnm="Cauzione_" & IdCauzione & ".pdf"
nome="PdfCauzioni/" & flnm
nmpt=VirtualPath & DirectoryUpload & nome
nmpt=replace(nmpt,"//","/")
filename=Server.MapPath(nmpt)

' Modifiche per Android
NomePdfOut = nome
PathC=Server.MapPath(VirtualPath)

'response.write filename
'response.end 

If ScriptObject.FileExists(filename) = true Then

    if FlagRewrite=1 then 
        ScriptObject.DeleteFile(filename) 
    else
        Const adTypeBinary = 1
        Dim strFilePath

        strFilePath = filename

        Set objStream = Server.CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = adTypeBinary
        objStream.LoadFromFile strFilePath 

        Response.BinaryWrite objStream.Read

        objStream.Close
        Set objStream = Nothing
        
        response.end 
    end if 
End If

Set DizCauzione = GetDizCauzione(IdCauzione,0)

Set pdf=CreateJsObject("FPDF")
pdf.CreatePDF()
pdf.SetPath("fpdf/")
pdf.SetFont "Arial","",12
pdf.Open()

pdf.AddPage()
pdf.SetFillColor 220

StartCol  =  10
MaxRow    = 270
MaxCol    = 195
BottomRow = 280

PrimaVolta=true
Dim NewPage,Col,Row
NewPage=false
'CurRow=MaxRow+1
   Col  = StartCol
   wCol = MaxCol-Col 
   Row=10

   'logo se previsto 
   if DescLogo<>"" then 
      FileLogo=Server.MapPath(DescLogo)
      If ScriptObject.FileExists(FileLogo) = true Then
         pdf.Image DescLogo,10,8,60,17
         Row = pdf.GetY()
      end if 
   end if 

   pdf.SetFont "Arial","B",6 

   IdAccountCliente = GetDiz(DizCauzione,"N_IdAccountCliente")
   MySql = ""
   MySql = MySql & " select * from Cliente Where IdAccount=" & IdAccountCliente
   Rs.CursorLocation = 3 
   Rs.Open MySql, ConnMsde   
   if rs.eof=false then 
      cf = Rs("CodiceFiscale")
      pi = Rs("PartitaIva")
      Denominazione = Rs("Denominazione")
      Cognome       = Rs("Cognome")
      Nome          = Rs("Nome")
      if cf="" then 
         cf=pi
         pi=""
      end if 
   else
      cf = ""
      pi = ""
      Denominazione = ""
      Cognome       = ""
      Nome          = ""
   end if 
   rs.close 
   IdProdottoTemplate = GetDiz(DizCauzione,"N_IdProdottoTemplate")
   DescTemplate = LeggiCampo("select * from ProdottoTemplate where IdProdottoTemplate=" & IdProdottoTemplate,"DescProdottoTemplate")
   if DescTemplate="" then 
      DescTemplate="Cauzione Provvisoria"
   end if 
   
   'scrivo contraente 
   Row = Row + 5
   xx=writeBox(Col   ,Row,MaxCol,Row+7,"Tipologia di Richiesta ","DF",DescTemplate,"","","")
   
   'scrivo contraente 
   Row = Row + 13
   xx=writeBox(Col   ,Row,Col+120,Row+7,"Obbligato principale/Contraente ","DF",Denominazione,"","","")
   xx=writeBox(Col+122,Row,MaxCol,Row+7,"CF/PI ","DF",cf & " " & pi,"","","")
   
   Row = Row + 13
   testo = GetDiz(DizCauzione,"S_Indirizzo")
   xx=writeBox(Col   ,Row,MaxCol-20,Row+7,"Indirizzo ","DF",testo,"","","")
   testo = GetDiz(DizCauzione,"S_Civico")
   xx=writeBox(MaxCol-18,Row,MaxCol,Row+7,"civico "  ,"DF",testo,"","","")

   Row = Row + 13
   testo = GetDiz(DizCauzione,"S_Cap")
   xx=writeBox(Col   ,Row,Col +20 ,Row+7,"Cap ","DF",testo,"","","")

   testo = GetDiz(DizCauzione,"S_Comune")
   xx=writeBox(Col+22,Row,Col +120 ,Row+7,"Comune","DF",testo,"","","")

   testo = GetDiz(DizCauzione,"S_Provincia")
   xx=writeBox(Col+122,Row,MaxCol ,Row+7,"Provincia","DF",testo,"","","")
   
   'beneficario 
   Row = Row + 18
   testo = GetDiz(DizCauzione,"S_Beneficiario")
   xx=writeBox(Col   ,Row,Col+120,Row+7,"Beneficiario ","DF",testo,"","","")
   testo = GetDiz(DizCauzione,"S_BeneficiarioCF") & " " & GetDiz(DizCauzione,"S_BeneficiarioPI") 
   xx=writeBox(Col+122,Row,MaxCol,Row+7,"CF/PI ","DF",testo,"","","")
   
   Row = Row + 13
   testo = GetDiz(DizCauzione,"S_BeneficiarioIndirizzo")
   xx=writeBox(Col   ,Row,MaxCol,Row+7,"Indirizzo ","DF",testo,"","","")

   
   Row = Row + 13
   testo = GetDiz(DizCauzione,"S_BeneficiarioCap")
   xx=writeBox(Col   ,Row,Col +20 ,Row+7,"Cap ","DF",testo,"","","")
   
   testo = GetDiz(DizCauzione,"S_BeneficiarioSede")
   xx=writeBox(Col+22,Row,Col +120 ,Row+7,"Comune/Sede","DF",testo,"","","")

   testo = GetDiz(DizCauzione,"S_BeneficiarioProvincia")
   xx=writeBox(Col+122,Row,MaxCol ,Row+7,"Provincia","DF",testo,"","","")
   
   'garanzia
   Row = Row + 18   
   testo = GetDiz(DizCauzione,"S_OggettoAppalto")
   xx=writeBoxCell(Col   ,Row,MaxCol,Row+47,"Descrizione Garanzia ","DF",testo)
   
   'garanzia
   ImptGaranzia = cdbl(GetDiz(DizCauzione,"N_ImportoLotto"))
   Row = Row + 55  
   FinCol = Col +40   
   testo = "   " & InsertPoint(ImptGaranzia,2)
   xx=writeBox(Col   ,Row, FinCol ,Row+7,"Importo Garantito Euro ","DF",testo,"","","")
   
   testo = " " & TrasformaInLettere(ImptGaranzia)
   Col    = FinCol +2 
   FinCol = MaxCol 
   xx=writeBox(Col   ,Row, FinCol ,Row+7,"Importo in lettere ","DF",testo,"","","")   
   'garanzia
   Row = Row + 12   
   testo = " " & Stod(GetDiz(DizCauzione,"N_ValidoDal"))
   testo = testo & " - " & Stod(GetDiz(DizCauzione,"N_ValidoAl"))
   xx=writeBox(Col   ,Row,Col +60 ,Row+7,"Periodo di garanzia  ","DF",testo,"","","")   
   
   testo = "  " & GetDiz(DizCauzione,"N_GiorniValidita")
   xx=writeBox(Col+62   ,Row, Col+92 ,Row+7,"Pari a giorni ","DF",testo,"","","")
   
   'garanzia
   Row    = Row + 18   
   Col    = StartCol
   FinCol = MaxCol
   testo  = GetDiz(DizCauzione,"S_noteServizioFornitore")
   xx=writeBoxCell(Col ,Row , FinCol , Row+47 ,"Note ","DF",testo)

   
   pdf.Output (filename), false
   pdf.Close()
   
   'controllo se esiste su upload come Riepilogo 
   IdTipoDocumento = LeggiCampo("Select * from Documento Where IdDocumentoInterno='RIEPILOGO'","IdDocumento")
   if Cdbl(IdTipoDocumento)>0 then 
      MyQ = "" 
      MyQ = MyQ & " select IdUpload From Upload "   
      MyQ = MyQ & " Where IdTabella='CAUZ_PROV' and IdTabellaKeyInt=" & IdCauzione 
      MyQ = MyQ & " and IdTipoDocumento=" & IdTipoDocumento
      IdUpload = Cdbl("0" & LeggiCampo(MyQ,"IdUpload"))
	  if Cdbl(IdUpload)=0 then 
         MyQ = "" 
         MyQ = MyQ & " insert into Upload (IdTabella,IdTabellaKeyInt,IdTabellaKeyString,DataUpload"
         MyQ = MyQ & " ,TimeUpload,IdTipoDocumento,DescBreve,DescEstesa,NomeDocumento,PathDocumento,ValidoDal,ValidoAl) "
         MyQ = MyQ & " values ("
         MyQ = MyQ & " 'CAUZ_PROV'"
         MyQ = MyQ & ", " & IdCauzione
         MyQ = MyQ & ",'" & IdCauzione & "'"
         MyQ = MyQ & "," & DTos() & ",0,'" & IdTipoDocumento & "','Riepilogo','Riepilogo','" & flnm & "','" & NomePdfOut & "',0,20991231)" 
         ConnMsde.execute MyQ 
      end if 
  
   end if 
   
   
   
   if SendBrowser="S" then
      response.redirect replace(virtualpath & DirectoryUpload & NomePdfOut,"\","/")
   else 
      response.write "OK:" & NomePdfOut
   end if 
   response.end
   


%> 