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
<!--#include file="fpdf.asp"-->
<%

Set Rs = Server.CreateObject("ADODB.Recordset")

IdEstrattoConto="0" & Request("IdEstrattoConto")

if IsNumeric(IdEstrattoConto)=false then 
   IdEstrattoConto=0
else
   IdEstrattoConto=cdbl(IdEstrattoConto)
end if 

'verifico se devo ricreare il pdf 

FlagRewrite=1

'verifico se il file esiste nel caso lo leggo e lo invio al browser
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")

nome="PdfEstrattoConto/EC" & IdEstrattoConto & ".pdf"
filename=Server.MapPath(VirtualPath & nome )

' Modifiche per Android
NomePdfOut = nome
PathC=Server.MapPath(VirtualPath)

response.write filename
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

 
Set pdf=CreateJsObject("FPDF")
pdf.CreatePDF()
pdf.SetPath("fpdf/")
pdf.SetFont "Arial","",12
pdf.Open()

PageNumber=0
MaxRow=270
PrimaVolta=true
Dim NewPage
NewPage=false
'CurRow=MaxRow+1
Row=0

'recuper dati estratto conto
DataEstratto=""
DescEstratto=""
ImptDare    =0
ImptDare    =0
DataLettura =""

MySql = ""
MySql = MySql & " Select a.*, isnull(DescTipoCredito,'') as descTipoCredito"
MySql = MySql & ",C.Nominativo as DescCliente,C.IdTipoAccount "
MySql = MySql & " From EstrattoConto A left join TipoCredito b  "
MySql = MySql & " on A.IdTipoCredito = B.IdTipoCredito "
MySql = MySql & " inner join Account C  "
MySql = MySql & " on A.IdAccount = C.IdAccount "
MySql = MySql & " Where A.IdEstrattoConto=" & IdEstrattoConto

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

   IdTipoCredito     = "ZZZZ"
   DescTipoCredito   = "ZZZZZZZZZZZZZZZ"
   IdAccount         = 0
   IdTipoAccount     = "CLIE"
   DescCliente       = "Cliente Prova"
   DataEstratto      = "01/01/2021"
   DescEstratto      = "non funziona"
   ImptEstratto      = 9999.99
   IdStatoEstratto   = "creazione"
   DescStatoEstratto = "da aggiungere "
   DataInizio        = "01/01/1999"
   DataFine          = "01/01/2002"

if not rs.eof then 
   IdTipoCredito     = Rs("IdTipoCredito")
   DescTipoCredito   = rs("DescTipoCredito")
   IdAccount         = rs("IdAccount")
   IdTipoAccount     = Rs("IdTipoAccount")
   DescCliente       = Rs("DescCliente")
   DataEstratto      = StoD(Rs("DataEstratto"))
   DescEstratto      = Rs("DescEstratto")
   ImptEstratto      = Rs("ImptEstratto")
   IdStatoEstratto   = Rs("IdStatoEstratto")
   DescStatoEstratto = "da aggiungere "
   DataInizio        = StoD(Rs("DataInizio"))
   DataFine          = StoD(Rs("DataFine"))
end if 

rs.close 

descLogo   =""
dNominativo=DescCliente

    MySql = ""
	MySql = MySql & " Select * "
	MySql = MySql & " From AccountMovEco "
	MySql = MySql & " where  IdEstrattoConto = " & IdEstrattoConto
	MySql = MySql & " Order By DataMovEco "

	response.write MySql 
	
	Rs.CursorLocation = 3 
	Rs.Open MySql, ConnMsde
	t_Dare =0
	t_Avere=0
	t_Paga =0
	t_Prov =0

	BottomRow=282
	NewWriga=0
	do while not rs.eof 

		xx=WriteHeader()
		If NewPage=true Then
			Wriga=Pdf.GetY()+4
			NewWRiga=0
			NewPage=false
		Else
			Wriga=Pdf.GetY()
		End If

		If cdbl(NewWRiga)>cdbl(Wriga) Then
			Wriga=NewWRiga
		End if
		
		pdf.SetFont "Arial","",8
		pdf.SetXY  10, Wriga
		pdf.Write   4, Stod(RS("DataMovEco"))
		
		pdf.SetXY  27, Wriga
		'pdf.Write   0, RS("DescCarrello")
		pdf.multiCell 105,4,Rs("DescMovEco")
		NewWriga=Pdf.GetY()

		ImptDare   = 0
		ImptAvere  = 0
		ImptPagato = 0
		if Rs("Segnosistema")=1 then 
		   ImptAvere = cdbl(Rs("ImptMovEco"))
		else
		   ImptDare  = cdbl(Rs("ImptMovEco"))
		end if 
		if Rs("IdTipoCredito") = "BORS" then 
		   ImptPagato = cdbl(Rs("ImptMovEco")) 
		end if 

		pdf.SetXY 131, Wriga
		if ImptDare<>0 then 
			pdf.Cell 18,4,formatnumber(ImptDare),0,0,"R"
		end if 

		pdf.SetXY 148, Wriga
		if ImptAvere<>0 then 
			pdf.Cell 18,4,formatnumber(ImptAvere),0,0,"R"
		end if 
			
		pdf.SetXY 165, Wriga
		if ImptPagato<>0 then 
			pdf.Cell 18,4,formatnumber(ImptPagato),0,0,"R"
		end if 

	
		t_Dare = t_Dare  + ImptDare
		t_Avere= t_Avere + ImptAvere
		t_Paga = t_paga  + ImptPagato
		
		rs.movenext 
		
		if rs.eof then 
			pdf.SetFont "Arial","B",8

			Wriga=Pdf.GetY()+4
			If cdbl(NewWRiga)>cdbl(Wriga) Then
				Wriga=NewWRiga
			End if
			'CurRow=Pdf.GetY()
			
			pdf.SetXY  90, Wriga
			pdf.Write   4, "Totali estratto"
			
			pdf.SetXY 131, Wriga
			pdf.Cell 18,4,formatnumber(t_Dare),0,0,"R"

			pdf.SetXY 148, Wriga
			pdf.Cell 18,4,formatnumber(t_Avere),0,0,"R"
				
			pdf.SetXY 165, Wriga
			pdf.Cell 18,4,formatnumber(t_Paga),0,0,"R"

		end if 
		
	loop
	rs.close


Set Rs=Nothing

pdf.SetFont "Arial","",8

' NomePdfOut = "prova.pdf"
pdf.Output (filename), false
response.redirect replace(virtualpath & NomePdfOut,"\","/")
response.end

pdf.Close()
pdf.Output(filename)
pdf.Output()

Function WriteHeader()

	CurRow=Pdf.GetY()
	if CurRow>MaxRow or PrimaVolta=true then 
		PrimaVolta=false
		NewPage=true

		CurRow=1
		PageNumber=PageNumber+1
		pdf.AddPage()
		if DescLogo<>"" then 
			FileLogo=Server.MapPath(DescLogo)
			If ScriptObject.FileExists(FileLogo) = true Then
				pdf.Image DescLogo,10,8,60,17
			end if 
		end if 
		
		Row = 25
		pdf.Rect 10,Row+  2,190,BottomRow-Row-2
		'linee orizzontali
		pdf.Line 10,Row+ 15,200,Row+ 15
		
		'linee verticali
		pdf.Line  27,Row +15, 27,BottomRow
		pdf.Line 132,Row +15,132,BottomRow 
		pdf.Line 149,Row +15,149,BottomRow 
		pdf.Line 166,Row +15,166,BottomRow 
		pdf.Line 183,Row +15,183,BottomRow 

		pdf.SetFont "Arial","B",8
		pdf.SetXY  120, 8  
		pdf.Cell 18,4,"Spett. ",0,0
		
		pdf.SetXY  130, 8  
		pdf.Cell 18,4,dNominativo,0,0
		
		Row = Row + 2
		pdf.SetFont "Arial","B",8
		pdf.SetXY  10, Row 
		pdf.Cell 18,4,"Estratto n." & IdEstrattoConto & " - " & DescEstratto,0,0

		pdf.SetXY  10, Row + 4
		pdf.Cell 18,4, "Del : " & DataEstratto ,0,0
		
		pdf.SetXY  40, Row + 4
		pdf.Cell 18,4, "Importo Estratto : " & formatnumber(ImptEstratto) ,0,0
		
		pdf.SetXY 180, Row + 4
		pdf.Cell 18,4,"Pag." & PageNumber ,0,0
		
		Row = Row + 14
		pdf.SetFont "Arial","B",8
		pdf.SetXY  10, Row 
		pdf.Cell 18,4,"Movim. Del",0,0

		pdf.SetXY  27, Row 
		pdf.Cell 18,4,"Descrizione",0,0		
		
		pdf.SetXY 124, Row 
		pdf.Cell 18,4,"Dare",0,0,"R"

		pdf.SetXY 142, Row 
		pdf.Cell 18,4,"Avere",0,0,"R"

		pdf.SetXY 164, Row 
		pdf.Cell 18,4,"Pagato",0,0,"R"
		
		Row = Row + 4
		pdf.SetFont "Arial","",8
		
		'scrivo righe
	
	end if 
end Function 	
%> 