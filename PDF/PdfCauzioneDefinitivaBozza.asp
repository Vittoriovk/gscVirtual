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
<!--#include virtual="/gscVirtual/modelli/FunctionCauzioneDef.asp"-->
<!--#include virtual="/gscVirtual/modelli/functionOpzioni.asp"-->
<!--#include file="fpdf.asp"-->
<%

Set Rs = Server.CreateObject("ADODB.Recordset")

IdCauzione  = "0" & Request("IdCauzione")
SendBrowser = Request("SendBrowser")

if IsNumeric(IdCauzione)=false then 
   IdCauzione=0
else
   IdCauzione=cdbl(IdCauzione)
end if 

'verifico se devo ricreare il pdf 

FlagRewrite=1

'verifico se il file esiste nel caso lo leggo e lo invio al browser
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")

nome="PdfCauzioni/CauzioneDef_" & IdCauzione & ".pdf"
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

Set DizCauzione = GetDizCauzioneDef(IdCauzione,0)

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
   
   'informazioni aggiuntive : sono cablate per tipo 
   MySql=GetQueryGeneTemplate("TECN",IdProdottoTemplate)   

   
   Dim ReadRs
   set ReadRs = ConnMsde.execute(MySql)
   if err.number=0 then 
      if not ReadRs.eof then 
         Rigo=ReadRs("Rigo")
         Row = Row + 18  
         iniCol = Col 
         Do while not ReadRs.eof 
           'cambio rigo
            if Rigo<>ReadRs("Rigo") then 
               Rigo=ReadRs("Rigo")
               Row = Row + 13
			   iniCol = Col
            end if 
            sizeCol=100
			testo  = getValoreOpzione("CAUZ_DEFI",IdCauzione,ReadRs("IdOpzione"),"ValoreOpzione")
            if ReadRs("Formato")="PERC" or  ReadRs("Formato")="NUMERO" then 
               sizeCol=50
			   testo = " " & insertPoint(testo,2)
            end if 
            if ReadRs("Formato")="TESTO" And cdbl(ReadRs("maxLen"))<51 then 
               sizeCol=50
            end if  
            
            xx=writeBox(IniCol,Row,IniCol + sizeCol ,Row+7,ReadRs("DescWeb"),"DF",testo,"","","")
			iniCol=iniCol+sizeCol+5
            ReadRs.MoveNext 
         loop
      end if 
   end if 


   
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
   if SendBrowser="S" then
      response.redirect replace(virtualpath & DirectoryUpload & NomePdfOut,"\","/")
   else 
      response.write "OK:" & NomePdfOut
   end if 
   response.end
   

function writeBox(x0,y0,x1,y1,header,fill,text1,text2,text3,text4)
'wx = larghezza finestra
'wy = altezza finestra 
dim wx,wy,criga 
   pdf.SetFont "Arial","",10
   wx = cdbl(x1)-cdbl(x0)
   wy = cdbl(y1)-cdbl(y0)
   
   pdf.Rect x0,y0 ,wx ,wy, fill
   if header<>"" then 
      pdf.SetFont "Arial","B",6
      pdf.Text x0,y0 - 1 , header 
   end if 
   pdf.SetFont "Arial","B",10
   cY = y0
   cX = x0 + 0.5
   if text1<>"" then 
      cY = CY + 5
  pdf.Text cX, cY , text1
   end if 
   if text2<>"" then 
      cY = CY + 3
  pdf.Text cX, cY , text2
   end if 

end function 

function writeBoxCell(x0,y0,x1,y1,header,fill,text1)
dim wx,wy,criga,fi,ws   
Dim wxMax 
    
   pdf.SetFont "Arial","",10
   wx = cdbl(x1)-cdbl(x0)
   wxMax = wx - 4
   wy = cdbl(y1)-cdbl(y0)
   pdf.Rect x0,y0 ,wx ,wy, fill
   if header<>"" then 
      pdf.SetFont "Arial","B",6
      pdf.Text x0,y0 - 1 , header
   end if 
   pdf.SetFont "Arial","B",10
   cY = y0
   cX = x0 + 0.5
   if text1<>"" then 
      'ciclo 
  ttw = ""
  lastSpace = 0
  do while text1<>"" 
     if len(text1)=1 then 
        ttw = ttw & Text1
    Text1 = ""
 else 
        ttw = ttw & Mid(Text1,1,1)
    Text1 = Mid(Text1,2)
 end if 
 if mid(ttw,len(ttw),1)=" " then 
    lastspace = len(ttw)
 end if 
 ws = pdf.GetStringWidth(ttw)
 if ws > wxMax or Text1="" then 
            cY = CY + 5
if lastSpace<>0 then 
           pdf.Text cX, cY , mid(ttw,1,lastspace)
               ttw = mid(ttw,lastSpace+1)

else 
           pdf.Text cX, cY , ttw
               ttw = ""
    end if
            lastspace = 0
 end if 
  loop 
   
   end if 
   'ws = pdf.GetStringWidth(text1)
   'cY = CY + 5
   'pdf.Text cX, cY , ws 

   'stringa = "ZZZZZZZZZ0ZZZZZZZZZ0ZZZZZZZZZ0"
   'stringa = stringa & stringa & "iiiiiiiii0ZZZZZZZZZ0ZZZZZ"
   'cY = CY + 5
   'pdf.Text cX, cY , stringa 
   
   'cY = CY + 5
   'pdf.Text cX, cY , pdf.GetStringWidth(stringa) 
   
   

end function 




%> 