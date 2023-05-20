<%@language=vbscript%>

<!--#include virtual="/gscVirtual/common/function.asp"-->
<!--#include virtual="/gscVirtual/common/functionNew.asp"-->
<!--#include virtual="/gscVirtual/common/connDb.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAccessoDb.asp"-->
<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<!--#include file="fpdf.asp"-->
<%

Set Rs = Server.CreateObject("ADODB.Recordset")

Savefile   = request("SaveFile")
IdCauzione = "0" & Request("IdCauzioneSt")

if IsNumeric(IdCauzione)=false then 
   IdCauzione=0
else
   IdCauzione=cdbl(IdCauzione)
end if 
IdAccount="0" & Session("IdAccountLogin")

if IsNumeric(IdAccount)=false then 
   IdAccount=0
else
   IdAccount=cdbl(IdAccount)
end if 

'lettura dei dati per la cauzione 


MySql = ""
MySql = MySql & " Select A.*,C.Denominazione as RagioneSociale " 
MySql = MySql & ", IsNull(A.OggettoAppalto,'') As DescBreveBando " 
MySql = MySql & ", IsNull(A.LuogoEsecuzione,'') As LuogoLavori " 
MySql = MySql & ", IsNull(A.ImportoLotto,0) As ImportoAsta " 
MySql = MySql & ", IsNull(C.PartitaIva,'') As PartitaIva " 
MySql = MySql & ", IsNull(C.CodiceFiscale,'') As CodiceFiscale " 
MySql = MySql & ", IsNull(A.Indirizzo,'') As Indirizzo " 
MySql = MySql & ", IsNull(A.CAP,'') As CAP " 
MySql = MySql & ", IsNull(A.Comune,'') As Comune " 
MySql = MySql & ", IsNull(A.Provincia,'') As Provincia " 
MySql = MySql & ", IsNull(A.TipologiaAppalto,'') As DescTipoBando " 
MySql = MySql & ", IsNull(A.Beneficiario,'') As DescEnte " 
MySql = MySql & ", IsNull(A.BeneficiarioSede,'') As SedeEnte " 
MySql = MySql & " From Cauzione A  "
MySql = MySql & "                 Left Join Cliente C "
MySql = MySql & " on   A.IdAccountCliente = C.IdAccount "
MySql = MySql & " Where A.IdCauzione=" & IdCauzione

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde
if not rs.eof then 

   v_ati=rs("DescATI")
	
	if v_ati<>"" then 
	   v_ati = "(ATI) "
	end if

	v_DataRichiestaProv  =Stod(rs("DataApertura"))
	v_ImptRichiestaProv  =Rs("ImptGarantito")
	v_CostoProv          =Rs("ImptTotaleCauzione")
	v_DataRichiestaDef   =0
	v_ImptRichiestaDef   =0
	v_CostoDef           =0
	v_IdTipoCauzione     =""
	v_IdTipoSpedizione   =""
	v_CostoTipoCauzione  =0
	v_CostoTipoSpedizione=0

	v_RagioneSociale = Rs("RagioneSociale")
	v_PartitaIva     = Rs("PartitaIva")
	v_CodiceFiscale  = Rs("CodiceFiscale")
	v_Indirizzo      = rs("Indirizzo")
	v_CAP            = rs("cap")
	v_Comune         = rs("Comune")
	v_Provincia      = rs("Provincia")
	
	v_ente           = lcase(rs("DescEnte"))
	v_SedeEnte       = lcase(rs("SedeEnte"))
		
    v_GaraAppalto   = rs("DescTipoBando")
	v_DataOfferta   = Rs("DataPubblicazione")
	v_DescGara      = rs("DescBreveBando")
	v_Luogo         = rs("LuogoLavori")
	v_costo         = rs("ImportoAsta")
	v_perc          = rs("PercGarantita")
	v_fide          = 0
	v_numero        = ""
	v_DataGara      = Stod(rs("DataApertura"))
    v_DataInizio    = Stod(rs("ValidoDal"))
    v_DataFine      = Stod(rs("ValidoAl"))
	v_NoteGara      = ""
	v_ImportoTasse  = 0
	
else
	v_DataRichiestaProv=Dtos()
	v_ImptRichiestaProv=999999/100
	v_CostoProv=9999/100
	v_DataRichiestaDef="20082010"
	v_ImptRichiestaDef=88888/100
	v_CostoDef=77777/100
	v_IdTipoCauzione=1
	v_IdTipoSpedizione=1
	v_CostoTipoCauzione=3333/100
	v_CostoTipoSpedizione=2222/100
	
	v_RagioneSociale="xxxxxxxxxxxxxxxxxxxxx"
	v_PartitaIva="qqqqqqqq"
	v_CodiceFiscale="qwqwqwqwqwqw"
	v_Indirizzo="xxcxcxcxcxcxcx xxcxcxcxc xcxcx xxcxcxcxc xcxcx xxcxcxc xcxcxcx xxcxcxcx cxcxcx xxcxcx cxcxcxcx "
	v_CAP="00000"
	v_Comune="ROMA ROMA ROMA ROMA ROMA ROMA ROMA ROMA ROMA ROMA ROMA ROMA ROMA ROMA ROMA ROMA "
	v_Provincia="RM"
	
	v_ente="Ente applatante ******** Ente applatante ******** Ente applatante ******** Ente applatante ******** Ente applatante ******** Ente applatante ******** "
	v_SedeEnte="Sede ente, Sede Ente, Sede Ente, Sede Ente "
   v_GaraAppalto="zzzzzzzzz"
	v_DataGara="x1/cv/aszx"
	v_DataOfferta="aaaaaaaaa"
	v_DescGara="descrizione gara - descrizione gara - descrizione gara - descrizione gara " 
	v_NoteGara="Annotazion della gara come esempio"
	v_Luogo="lululul"
	v_costo = "999999"
	v_perc = "3"
	v_fide = "555"
	
	v_numero="1"

   v_DataInizio="20/11/2008"
   v_DataFine  ="25/05/2009"
	v_ImportoTasse="5"
end if 
rs.close
v_importoTotale=0
v_importoTotale=v_importoTotale + cdbl(TestNumeroPos(v_CostoProv))

v_ImportoAut=0
v_ImportoAlt=0


Set pdf=CreateJsObject("FPDF")
pdf.CreatePDF()
pdf.SetPath("fpdf/")
pdf.SetFont "Arial","",12
pdf.Open()
pdf.AddPage()

Row=5
pdf.SetFont "Arial","B",12 
pdf.SetXY 30, Row
pdf.Cell 10,10,"APPALTI"

Row=10
pdf.SetFont "Arial","B",8 
pdf.SetXY 120, Row
pdf.Cell 10,10,"Sede legale:"

Row=Row+3
pdf.SetFont "Arial","",8 
pdf.SetXY 120, Row
pdf.Cell 10,10,""

Row=Row+3
pdf.SetFont "Arial","B",8 
pdf.SetXY 120, Row
pdf.Cell 10,10,"Sede amministrativa:"

Row=Row+3
pdf.SetFont "Arial","",8 
pdf.SetXY 120, Row
pdf.Cell 10,10,""

Row=Row+3
pdf.SetFont "Arial","",8 
pdf.SetXY 120, Row
pdf.Cell 10,10,""

Row=Row+3
pdf.SetFont "Arial","B",8 
pdf.SetXY 120, Row
pdf.Cell 10,10,""

Row=Row+3
pdf.SetFont "Arial","",8 
pdf.SetXY 120, Row
pdf.Cell 10,10,""

Row=Row+3
pdf.SetFont "Arial","",8 
pdf.SetXY 120, Row
pdf.Cell 10,10,"Numero verde  - e-mail: "
  
Row=Row+8
pdf.SetFont "Arial","B",10
pdf.SetXY  80, Row 
pdf.Cell 10,10,"POLIZZA FIDEJUSSORIA"

Row=Row+5
pdf.SetFont "Arial","B",8
pdf.SetXY  30, Row 
pdf.Cell 10,10,"ai sensi dell'art.30,comma 1,della legge n.109/94 e delle successive modifiche di cui all'art.75 del Dlgs.163/2006"

Row=Row+4
pdf.SetFont "Arial","",8
pdf.SetXY 10, Row
pdf.Cell 10,10,"La presente scheda tecnica costituisce parte integrante dello Schema Tipo 1.1 di cui al D.M. 12 marzo 2004, n. 123 e Dlgs n.163 del 12/04/2006 e"
Row=Row+3
pdf.SetXY 10, Row
pdf.Cell 10,10,"riporta i dati e le informazioni necessarie all'attivazione della garanzia fidejussoria di cui al citato Schema Tipo: la sua sottoscrizione costituisce atto"

Row=Row+3
pdf.SetXY 10, Row
pdf.Cell 10,10,"formale di accettazione incondizionata di tutte le condizioni previste nelo Schema Tipo e di quanto disposto dall'art. 75 del Dlgs 163/2006 "

Row=Row+6
pdf.SetFont "Arial","",10

pdf.Rect 9,Row+1,190,9
pdf.Line 60,Row+1,60,Row+10 
pdf.SetXY 10, Row-1
pdf.Cell 10,10,"SCHEMA TIPO 1.1"

pdf.SetXY 65, Row+1
pdf.Cell 10,10,"GARANZIA FIDEJUSSORIA PER LA CAUZIONE PROVVISORIA"

pdf.SetXY 10, Row+3 
pdf.Cell 10,10,"SCHEDA TECNICA 1.1"

Row=Row+9
pdf.Rect  9,Row+2,190,10
pdf.Line 60,Row+2, 60,Row+12 
pdf.Line  9,Row+7,199,Row+ 7
pdf.SetXY 10, Row
pdf.Cell 10,10,"Garanzia fidejussoria n.                  Rilasciata da "

pdf.SetFont "Arial","B",10
pdf.SetXY 15, Row+5 
pdf.Cell 10,10,"A" & right("000000" & v_numero,7)

pdf.SetFont "Arial","B",10
pdf.SetXY 65, Row+5 
pdf.Cell 10,10," <NECT BROKER "

Row=Row+7
pdf.SetFont "Arial","B",8 
pdf.Rect 9,Row+6,190,10
pdf.Line  9,Row+10,199,Row+ 10
pdf.Line 160,Row+6, 160,Row+16
pdf.SetXY 10, Row+3
pdf.Cell 10,10,"Contraente (obbligato principale)"
pdf.SetFont "Arial","",10 
pdf.SetXY 10, Row+8
pdf.Cell 10,10,v_ati & v_RagioneSociale


pdf.SetFont "Arial","B",8 
pdf.SetXY 160, Row+3
pdf.Cell 10,10,"C.F./P.IVA"
pdf.SetFont "Arial","",10 
pdf.SetXY 160, Row+8
pdf.Cell 10,10,v_PartitaIva



pdf.SetFont "Arial","B",8 

Row=Row+10
pdf.Rect 9,Row+6,190,18
pdf.Line 9,Row+10,199,Row+ 10

pdf.Line  80,Row+6,  80,Row+24
pdf.Line 160,Row+6, 160,Row+24
pdf.Line 175,Row+6, 175,Row+24

pdf.SetXY 10, Row+3
pdf.Cell 10,10,"Via/Piazza,n.civico"
pdf.SetXY 80, Row+3
pdf.Cell 10,10,"Località"
pdf.SetXY 160, Row+3
pdf.Cell 10,10,"Cap"
pdf.SetXY 175, Row+3
pdf.Cell 10,10,"Prov."


pdf.SetFont "Arial","",10 
pdf.SetXY  10, row+11
pdf.MultiCell 65,3,v_Indirizzo

pdf.SetXY  80, row+11
pdf.MultiCell 65,3,v_comune

pdf.SetXY 160, row+8
pdf.Cell 10,10,v_cap

pdf.SetXY 175, row+8
pdf.Cell 10,10,v_Provincia


Row=Row+17
pdf.SetFont "Arial","B",8 
pdf.Rect   9,Row+8,190,15
pdf.Line   9,Row+12,199,Row+ 12
pdf.Line 140,Row+8, 140,Row+23
pdf.SetXY  10, Row+5
pdf.Cell 10,10,"Stazione appaltante (Beneficiario)"
pdf.SetFont "Arial","",10 
pdf.SetXY 140, Row+5
pdf.Cell 10,10,"Sede"

pdf.SetXY  10, Row+13
pdf.MultiCell 130,3,v_ente

pdf.SetXY 140, Row+13
pdf.MultiCell 55,3,v_SedeEnte

Row=Row+20
pdf.SetFont "Arial","B",8 
pdf.Rect 9,Row+4,190,9
pdf.Line   9,Row+8,199,Row+8
pdf.Line 130,Row+4, 130,Row+13
pdf.Line 160,Row+4, 160,Row+13
pdf.SetXY  10, Row+1
pdf.Cell 10,10,"Gara D'appalto"
pdf.SetXY 130, Row+1
pdf.Cell 10,10,"Gara del giorno"
pdf.SetXY 160, Row+1
pdf.Cell 10,10,"Data presentazione offerta"

row=row+6
pdf.SetFont "Arial","",10 
pdf.SetXY 10, row
pdf.Cell 10,10,v_GaraAppalto

pdf.SetXY 135, row
pdf.Cell 10,10,v_DataGara

pdf.SetXY 170, row
pdf.Cell 10,10,v_DataInizio

Row=Row+3 
pdf.SetFont "Arial","B",8 
pdf.Rect 9,Row+5,190,38
pdf.Line 9,Row+10,199,Row+10
pdf.SetXY 90, Row+3
pdf.Cell 10,10,"Descrizione Opera"
pdf.SetFont "Arial","",10 
pdf.SetXY 10, row+12
v_DescGara=lcase(v_DescGara)
if v_NoteGara<>"" then 
   v_DescGara=v_DescGara & chr(10) & chr(10) & "ANNOTAZIONI:" & v_NoteGara
end if 
pdf.MultiCell 175,3,v_DescGara


Row=row+25 
pdf.SetFont "Arial","B",8 
pdf.Rect 9,Row+18,190,6
pdf.SetXY 10, Row+16
pdf.Cell 10,10,"Luogo di Esecuzione:"

pdf.SetFont "Arial","",10 
pdf.SetXY 40, row+16
pdf.Cell 10,10,v_Luogo

Row=row+8 
pdf.SetFont "Arial","B",8 
pdf.Rect 9,Row+18,190,12
pdf.Line 9,Row+24,120,Row+24
pdf.Line 120,Row+18, 120,Row+30

pdf.SetXY 10, Row+16
pdf.Cell 10,10,"Costo complessivo previsto dell'opera"
pdf.SetFont "Arial","",10 
pdf.SetXY 20, row+22
pdf.Cell 10,10,formatnumber(v_costo) & " €"

pdf.SetXY 150, row+22
pdf.Cell 10,10,formatnumber(v_fide) & " €"


pdf.SetFont "Arial","B",8 
pdf.SetXY 120, Row+16
pdf.Cell 10,10,"Somma garantita           % del costo complessivo previsto"
pdf.SetFont "Arial","",10 
pdf.SetXY 145, row+16
pdf.Cell 10,10,v_perc


Row=row+30
pdf.SetFont "Arial","B",8 
pdf.Rect 9,Row+ 2,190,11
pdf.Line 9,Row+ 7,199,Row+7
pdf.Line 99,Row+2, 99,Row+13
pdf.SetXY 10, Row
pdf.Cell 10,10,"Data inizio polizza fidejussoria - v.art.2 Schema Tipo 1.1"
pdf.SetXY 99, Row
pdf.Cell 10,10,"Data cessazione polizza fidejussoria - v.art.2 Schema Tipo 1.1"

pdf.SetFont "Arial","B",10
pdf.SetXY 40, Row+5
pdf.Cell 10,10,v_DataInizio

pdf.SetXY 140, Row+5
pdf.Cell 10,10,v_DataFine


Row=row+13
pdf.SetFont "Arial","B",8 
pdf.Rect 9,Row+ 2,190,11
pdf.Line 9,Row+ 7 ,199,Row+7
pdf.Line  33,Row+2 , 33,Row+13
pdf.Line  66,Row+2 , 66,Row+13
pdf.Line  99,Row+2 , 99,Row+13
pdf.Line 132,Row+2, 132,Row+13
pdf.Line 165,Row+2, 165,Row+13
pdf.SetXY 10, Row
pdf.Cell 10,10,"PREMIO"

pdf.SetXY 33, Row+5
pdf.Cell 10,10,"       Premio Netto"

pdf.SetXY 66, Row+5
pdf.Cell 10,10,"          Accessori"

pdf.SetXY 99, Row+5
pdf.Cell 10,10,"        Autentica"

pdf.SetXY 132, Row+5
pdf.Cell 10,10,"             Tasse"

pdf.SetXY 165, Row+5
pdf.Cell 10,10,"            Totale"


pdf.SetFont "Arial","",10 
pdf.SetXY  40, Row
pdf.Cell 0,10,formatnumber(v_CostoProv) & " €"

pdf.SetXY  75, Row
pdf.Cell 0,10,formatnumber(v_ImportoAlt) & " €"

pdf.SetXY 105, Row
pdf.Cell 0,10,formatnumber(v_ImportoAut) & " €"

pdf.SetXY 140, Row
pdf.Cell 0,10,formatnumber(v_ImportoTasse) & " €"

pdf.SetXY 170, Row
pdf.Cell 0,10,formatnumber(v_importoTotale) & " €"

pdf.Close()
if Trim(Savefile) <> "" then
    pathFile=Server.MapPath("\gscvirtual") & "\ServiziUpload\" & Savefile
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(pathFile) Then
	   fso.DeleteFile pathFile
    end if 
	pdf.Output (Server.MapPath("\gscvirtual") & "\ServiziUpload\" & Savefile), true
	response.write "OK"
else
	pdf.Output()
end if

%> 
