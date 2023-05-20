<%
Const DEFAULT_TEXT_CRIPTAZIONE = "AppaltiCauzioni"											
Const StatoPolizza_Incompleta=1
Const StatoPolizza_InLavorazione=2
Const StatoPolizza_Approvata=3
Const StatoPolizza_InAttesa=4
Const StatoPolizza_Respinta=5

Function checkSQLsemplice(textValue)
    textValue=ltrim(rtrim(TextValue))
	If Len(textValue) > 0 AND IsNull(textValue) = False Then
	   textValue = Replace(textValue, "'", "''")
       textValue = Replace(textValue, """", """""")
       textValue = Replace(textValue, "\", "")
    End If
    checkSQLsemplice = textValue
End Function


Function CheckTimePageLoad()
RetVal=false
if Session("TimePageLoad")=Request("TimePageLoad") then 
   RetVal=true  
end if 
CheckTimePageLoad = RetVal 
End function 

Function GetIdentity()
Dim MyQ,Id,GetRs 

	On Error Resume Next
	Set GetRs = Server.CreateObject("ADODB.Recordset")
	err.clear
   Id=0
	MyQ=""
	MyQ=MyQ & " SELECT SCOPE_IDENTITY() As Id "
	GetRs.CursorLocation = 3
	GetRs.Open MyQ, ConnMsde
	Id=GetRs("Id")
	GetRs.close
	if err.number>0 then 
	   Id=0
	end if
	err.clear
	GetIdentity=Id
	Set GetRs = nothing
	
end function


function SaveFileFromUrl(Url, FileName)
    dim objXMLHTTP, objADOStream, objFSO

    ' Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")	

    objXMLHTTP.open "GET", Url, false
    objXMLHTTP.send()

    If objXMLHTTP.Status = 200 Then 
        Set objADOStream = CreateObject("ADODB.Stream")
        objADOStream.Open
        objADOStream.Type = 1 'adTypeBinary

        objADOStream.Write objXMLHTTP.ResponseBody
        objADOStream.Position = 0 'Set the stream position to the start

        Set objFSO = Createobject("Scripting.FileSystemObject")
        If objFSO.Fileexists(FileName) Then objFSO.DeleteFile FileName
        Set objFSO = Nothing

        objADOStream.SaveToFile FileName
        objADOStream.Close
        Set objADOStream = Nothing

        SaveFileFromUrl = objXMLHTTP.getResponseHeader("Content-Type")
    else
        SaveFileFromUrl = "KO"
    End if

    Set objXMLHTTP = Nothing
end function

Function CallHttpPdf(ServiceToCall,PathC,PostData)
	Dim ret, HttpDown
	ret = true
	
	set HttpDown = Server.createObject("Microsoft.XMLHTTP")

	StartSito = ucase(Request.ServerVariables("SERVER_NAME") & VirtualPath)
	if Mid(StartSito,1,4)<>"HTTP" then 
	   StartSito = "HTTP://" & StartSito
	end if 
	
	ServiceToCall = StartSito & ServiceToCall
	HttpDown.open "POST", ServiceToCall, false 
	HttpDown.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	' HttpDown.setRequestHeader "HTTP_REFERER", "http://localhost"

	HttpDown.Send(PostData)

	status = CInt(HttpDown.status)
	if status <> 200 then
		MsgErrore = status
        ret = false
	else    
		MsgErrore=HttpDown.responseText
	end if 
	set HttpDown = nothing

	CallHttpPdf = ret
End Function

Function VerificaCampo(Testo,ValoreOld,ValoreNew)
on error resume next 
	if ValoreOld=ValoreNew then 
	   response.write Testo
	else
	   response.write "<font color=red><label title='valore iniziale=" & ValoreOld & "'>" & Testo & "</label></font>"
	end if
	err.clear

end function

Function GetDescTipoCalcolo(ValoreKey,RetCampo)
	GetDescTipoCalcolo=LeggiCampo("Select * from ElencoKey Where IdKey='ModCalcolo' and ValoreKey='" & ValoreKey & "'",RetCampo)
End Function


Function LnkCaricaImg(LogoDiv,size)
Dim lsize
   lSize=""
	if size<>"" then 
	   lSize=" height='" & size & "' "
	end if 
	response.write "<img " & lsize & "  src='" & LogoDiv & "' border='0' align='absmiddle'>"
	
End Function


Function generatePassword(passwordLength)
'Declare variables
    Dim isUpper 
	Dim isLower
	Dim isDigit 
	Dim isSymbol
	Dim sDefaultChars
	Dim iCounter
	Dim sMyPassword
	Dim iPickedChar
	Dim iDefaultCharactersLength
	Dim iPasswordLength
	Dim tmpC 
	sDefaultChars="abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789-.,!$"
	iPasswordLength=passwordLength
	iDefaultCharactersLength = Len(sDefaultChars) 
	Randomize
    isUpper  = false
	isLower  = false
	isDigit  = false
	isSymbol = false
	for iCounter = 1 To iPasswordLength
		iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1)
		tmpC = Mid(sDefaultChars,iPickedChar,1)
        if tmpC>="A" and tmpC<"Z" then 
		   isUpper  = true 
		end if 
        if tmpC>="a" and tmpC<"z" then 
		   isLower  = true 
		end if 
        if tmpC>="0" and tmpC<"9" then 
		   isDigit  = true 
		end if 
        if tmpC="-" or tmpC="." or tmpC="," or tmpC="!" or tmpC="&" then 
		   isSymbol = true 
		end if 		
		sMyPassword = sMyPassword & Mid(sDefaultChars,iPickedChar,1)
	Next 
    if isUpper  = false then
       sMyPassword = sMyPassword & "K"
    end if 
	if isLower  = false then 
       sMyPassword = sMyPassword & "r"
    end if 
	if isDigit  = false then 
       sMyPassword = sMyPassword & "5"
    end if 
	if isSymbol = false then
       sMyPassword = sMyPassword & "."
    end if 
	generatePassword = sMyPassword
End Function

Function DisplayDataPicker(NomeD)
Dim X 
	x="<input type=button style=""width:20px"" value="".."" onclick=""displayDatePicker('" & NomeD & "', false, 'dmy', '/');"">"
   DisplayDataPicker=X
end function 

Function GetDivisioneAccount(UserId)
Dim ReadRs,Sql,IdDivisione 
	on error resume next

	IdDivisione=0
	Session("Fun_Descrizione")=""

	Sql = ""
	Sql = Sql & "select B.IdDivisione,B.DescBreveDivisione "
	
	if Session("CategoriaLogin")="ACCOUNT" then
		Sql = Sql & " from Account A"
	else
	   Sql = Sql & " from Utente  A "
	end if
	Sql = Sql & " ,divisione B"
	Sql = Sql & " Where A.IdDivisione=B.IdDivisione "
	Sql = Sql & " And   A.UserId='" & Apici(UserId) & "'"
	
	'response.write sql
	set ReadRs = ConnMsde.execute(Sql)
	if err.number=0 then 
		if not ReadRs.eof then
			IdDivisione = ReadRs("IdDivisione")
			Session("Fun_Descrizione")=ReadRs("DescBreveDivisione")
		end if
	end if 
	ReadRs.close
	Set ReadRs=nothing
	err.clear
	
	GetDivisioneAccount=IdDivisione
	
End function 

Function GetTrattFiscaleServizio(IdS,DataR)
Dim MyQ,GetRs 

	On Error Resume Next
	Set GetRs = Server.CreateObject("ADODB.Recordset")

	GetTrattFiscaleServizio=false
	
	Session("Temp_IdTrattFiscale")=""
	Session("Temp_Aliquota")=""
	Session("Temp_DescEsenzione1")=""
	Session("Temp_DescEsenzione2")=""

	MyQ=""
	MyQ=MyQ & " select b.* from TrattFiscaleServizio a, AliquoteTrattFiscale B "
	MyQ=MyQ & " Where A.IdServizio='" & apici(IdS) & "'"
	MyQ=MyQ & " and   A.ValidoDal<=" & DataR
	MyQ=MyQ & " and   A.ValidoAl >=" & DataR
	MyQ=MyQ & " and   A.IdTrattFiscale=B.IdTrattFiscale "
	MyQ=MyQ & " and   B.ValidoDal<=" & DataR
	MyQ=MyQ & " and   B.ValidoAl >=" & DataR

	'response.write MyQ
	
	GetRs.CursorLocation = 3
	GetRs.Open MyQ, ConnMsde
	
	If GetRs.eof=false then 
		Session("Temp_IdTrattFiscale")=GetRs("IdTrattFiscale")
		Session("Temp_Aliquota")=GetRs("Aliquota")
		Session("Temp_DescEsenzione1")=GetRs("DescEsenzione1")
		Session("Temp_DescEsenzione2")=GetRs("DescEsenzione2")
		GetTrattFiscaleServizio=true
	end if 
	GetRs.close

	err.clear
	

	
	Set GetRs = nothing
end function

Function GetTableIdentity(TableName)
Dim MyQ,Id,GetRs 

	On Error Resume Next
	Set GetRs = Server.CreateObject("ADODB.Recordset")
	err.clear
   Id=0
	MyQ=""
	
	MyQ=MyQ & " SELECT isnull(IDENT_CURRENT('" & TableName & "'),0) As Id "
	GetRs.CursorLocation = 3
	GetRs.Open MyQ, ConnMsde
	
	
	if err.number>0 then
	   Id=0
	else
		Id=GetRs("Id")
		GetRs.close
		if err.number>0 then 
		   Id=0
		end if
	end if 
	err.clear
	GetTableIdentity=Id
	Set GetRs = nothing

end function


Function Audit(Testo)
Dim AuditQ
	on error resume next 
	
	AuditQ=""
	AuditQ=AuditQ & " Insert into Audit (Info,UserId) "
	AuditQ=AuditQ & " values" 
	AuditQ=AuditQ & "('" & Apici(Testo) & "'"
	AuditQ=AuditQ & ",'" & Apici(Session("UtenteLogin_Account") & "-" & Session("UtenteLogin_Nominativo")) & "'"
	AuditQ=AuditQ & ")"
	
	ConnMsde.execute AuditQ
	err.clear
	
end function 

Function TestNumeroPos(NN)
Dim RetV
   RetV=0
	xx="0" & NN
	if IsNumeric(xx)=true then 
		RetV=Cdbl(xx)
	end if 
	TestNumeroPos=RetV
end function 

Function TestNumeroNeg(NN)
Dim RetV,ValAbs 
	RetV=0
	xx="0" & NN
	if IsNumeric(xx)=true then 
		RetV=Cdbl(xx)
	else
		If IsNumeric(NN)=true Then
			ValAbs=abs(NN)
			xx="0" & ValAbs
			if IsNumeric(xx)=true then 
				RetV=Cdbl(NN)
			End If
		End If 
	end if 
	TestNumeroNeg=RetV	
end function 

Function EsisteFile(File,ErrC)
Dim fs,fn
  on error resume next 
  ErrC=1
  fn=server.mappath("\") & virtualpath & "\" & File   
  
  set fs=Server.CreateObject("Scripting.FileSystemObject")
  if fs.FileExists(fn)=false or err.number<>0 then
      ErrC=0
  end if
  set fs=nothing
  err.clear
    
End Function


Function DataToSqlServer(DataIn)

	if Instr(DataIn,"/")>0 then 
	   DataToSqlServer="convert(smalldatetime,'" & DataIn &"',103)" 
	end if 
	
End Function 

Function DataToSqlServerNull(DataIn)
Dim RetV 

	RetV="null"

	if Instr(DataIn,"/")>0 then 

	   RetV="convert(smalldatetime,'" & DataIn &"',103)" 
	end if 
	DataToSqlServerNull=RetV
End Function 


Function RequestHtml(Obj)
	RequestHtml=server.htmlEncode(Request(Obj))
End Function

Function InsertPoint(Num,Dec)
   InsertPoint=FormatNumber (Num,Dec, true) 
End function

Function LeggiCampo(Q,C)
Dim ReadRs 
	on error resume next
    Set ReadRs = Server.CreateObject("ADODB.Recordset")
    ReadRs.Open Q, ConnMsde

    If ReadRs.EOF then	
	   LeggiCampo=""
	else
	   LeggiCampo=ReadRs(C)
	end if 
	Set ReadRs=nothing
	err.clear
	
End function

Function LeggiStatoServizio(id,campo)
Dim q,r
   q="select * from StatoServizio Where IdStatoServizio='" & id & "'"
   if Campo="" then 
      r=LeggiCampo(q,"DescStatoServizio")
   else
      r=LeggiCampo(q,campo)
   end if 
   LeggiStatoServizio = r
End function

Function LeggiCampoTabella(T,id)
Dim ReadRs,Q 
	on error resume next
	Q="select * from " & T & " Where Id" & T & "=" & id
    Set ReadRs = Server.CreateObject("ADODB.Recordset")
    ReadRs.Open Q, ConnMsde

    If ReadRs.EOF then	
	   LeggiCampoTabella=""
	else
	   LeggiCampoTabella=ReadRs("Desc" & T)
	end if 
	Set ReadRs=nothing
	err.clear
	
End function

Function LeggiCampoTabellaText(T,id)
Dim ReadRs,Q 
	on error resume next
	Q="select * from " & T & " Where Id" & T & "='" & apici(id) & "'"
    Set ReadRs = Server.CreateObject("ADODB.Recordset")
    ReadRs.Open Q, ConnMsde

    If ReadRs.EOF then	
	   LeggiCampoTabellaText=id
	else
	   LeggiCampoTabellaText=ReadRs("Desc" & T)
	end if 
	Set ReadRs=nothing
	err.clear
	
End function

Function ImgInsert(AltInfo)
    If AltInfo="" then 
	   AltInfo="Registra nuovo record"
	end if 
	ImgInsert  ="Salva" '"<img src=""images/registra.gif"" alt='" & AltInfo & "' border='0'>"
end function

Function ImgGestisci(AltInfo)
	ImgGestisci="<img src=""images/rinnova.gif""  alt='" & AltInfo & "' border=""0"">"
end function

Function ImgAttenzioneVai(AltInfo,Href)

	X=""
	X=X & "<img src='" & VirtualPath & "/images/Attenzione.gif'  alt='" & AltInfo & "' border='0'>" & AltInfo 
	X=X & "<a href='" & Href & "'><img src='" & VirtualPath & "/images/vai.gif'  border='0'></a>"
	ImgAttenzioneVai=X
end function

Function ImgControlla(AltInfo)
	ImgControlla="-Controlla"
end function

Function ImgGenerica(AltInfo,ImgName)
	if AltInfo="" then 
	   AltInfo="Registra"
	end if 
	ImgGenerica="<img src=""" & VirtualPath & "/images/" & ImgName & """ border=0  align=absmiddle alt=""" & AltInfo & """ title=""" & AltInfo & """>"
end function

Function ImgGenericaPath(AltInfo,ImgName,Sz)
	if AltInfo="" then 
	   AltInfo="Registra"
	end if 
	ImgGenericaPath="<img " & sz & " src=""" & VirtualPath & "/" & ImgName & """ border=0  align=absmiddle alt=""" & AltInfo & """ title=""" & AltInfo & """>"
end function

Function ImgRegistra(AltInfo)
	if AltInfo="" then 
	   AltInfo="Registra"
	end if 
	ImgRegistra="<img src=""" & VirtualPath & "/images/registra.gif"" border=0  align=absmiddle alt=""" & AltInfo & """ title=""" & AltInfo & """>"
end function

Function ImgModifica(AltInfo)
	if AltInfo="" then 
	   AltInfo="Modifica"
	end if 
	ImgModifica="<img src=""" & VirtualPath & "/images/modifica.gif"" border=0  align=absmiddle alt=""" & AltInfo & """ title=""" & AltInfo & """>"
end function
Function ImgLista(AltInfo)
	if AltInfo="" then 
	   AltInfo="Modifica"
	end if 
	ImgLista="<img src=""" & VirtualPath & "/images/lista.gif"" border=0  align=absmiddle alt=""" & AltInfo & """ title=""" & AltInfo & """>"
end function
Function ImgCestino(AltInfo)
	if AltInfo="" then 
	   AltInfo="Cancella"
	end if 
	ImgCestino="<img src=""" & VirtualPath & "/images/Cestino.gif"" border=0  align=absmiddle alt=""" & AltInfo & """ title=""" & AltInfo & """>"
end function


Function GetTitolo(Gest,AddInfo)
	GetTitolo=Gest 
	if Trim(AddInfo)<>"" then 
	   GetTitolo=GetTitolo & ">><i><b><u><font size=3> " & AddInfo & "</font></u></b></I><<"
    end if 	
End Function


Function Apici(X)
Dim retV 
    on error resume next 
    retV=X 
	'provo a pulire
	retV=Pulisci(X)
	if Err.Number = 0 then 
	   retV = replace(retV,"'","''")
	else 
	   retV = replace(X,"'","''")
    end if 
	Apici=retV
End Function

Function StoNum(X)
	StoNum=replace(X,",",".")
End Function


Function NumForDb(X)
dim v
    v=TestNumeroNeg(X)
	NumForDb=replace(V,",",".")
End Function

Function TimeToS()
'restituire il time informato HHMMSS
	TimeToS=""
	TimeToS=TimeToS & Right("0" & hour(time()),2)
	TimeToS=TimeToS & Right("0" & minute(time()),2)
	TimeToS=TimeToS & Right("0" & second(time()),2)
End Function

Function StoTime(TimeSeriale)
'restituire la data nel formato HH:MM:SS
    TimeSeriale = Right ("00" & TimeSeriale,6)
    StoTime = ""
	StoTime = StoTime & mid(TimeSeriale,1,2) & ":"
	StoTime = StoTime & mid(TimeSeriale,3,2) & ":"	
	StoTime = StoTime & mid(TimeSeriale,5,2) 
End Function

Function DtoS()
'restituire la data nel formato AAAAMMGG
	DtoS = ""
	DtoS = Dtos & Right("20" & Year(Date()),4)
	DtoS = Dtos & Right("0"  & Month(Date()),2)
	DtoS = Dtos & Right("0"  & Day(Date()),2)
End Function

Function DataStringa(DT)
'restituire la data nel formato AAAAMMGG
	DataStringa = ""
	DataStringa = DataStringa & mid(dt,7,4)
	DataStringa = DataStringa & mid(dt,4,2)
	DataStringa = DataStringa & mid(dt,1,2)
End Function

Function StoD(DataSeriale)
'restituire la data nel formato GG/MM/AAAA
   if len(DataSeriale)=8 then 
      StoD = ""
      StoD = StoD & mid(DataSeriale,7,2) & "/"
      StoD = StoD & mid(DataSeriale,5,2) & "/"	
      StoD = StoD & mid(DataSeriale,1,4) 
   else
      StoD = ""
   end if 
End Function

Function StoIso(DataSeriale)
'restituire la data nel formato AAAA-MM-GG
    StoIso = ""
	StoIso = StoIso & mid(DataSeriale,1,4) & "-"
	StoIso = StoIso & mid(DataSeriale,5,2) & "-"
	StoIso = StoIso & mid(DataSeriale,7,2) 	
	 
End Function

Function NumToData(DataSeriale)
'restituire la data nel formato GG/MM/AAAA se è un numero di 8 cifre 
Dim nn 
    nn=DataSeriale
	if len(nn)=8 then 
		NumToData = Stod(nn)
	else
        NumToData = ""	
	end if 
	
End Function


Function OptionSiNo(Nome,Valore,Parm)
Dim X 
   X=""
   X=X & "<td  align='center' " & Parm & " valign='middle' class='boldTableCalendario'>"

   SelS=""
   SelN=""

   if ucase(Valore)="S" then 
      SelS=" selected "
   else
	  SelN=" selected "
   end if 
   X=X & "<select name=SiNo_" & Nome & " class='new_inputText'>"
   X=X & "<option value ='S' " & SelS & " >Si</option>"
   X=X & "<option value ='N' " & SelN & " >No</option>"
   X=X & "</select>"
   X=X & "</td>"
   OptionSiNo=X
End Function

Function OptionDataNoTd(Nome,Anno,tmese,Giorno,Parm)
Dim xx
	xx=OptionData(Nome,Anno,tmese,Giorno,Parm)
	ptr=instr(xx,"<select")
	if ptr>0 then 
		xx=mid(xx,ptr)
	end if 
	ptr=instr(xx,"</TD>")
	if ptr>0 then 
		xx=mid(xx,1,ptr-1)
	end if 
	
	OptionDataNoTd=xx
End Function 

Function OptionData(Nome,Anno,tmese,Giorno,Parm)
Dim X,J,mm(13),sel


    for j=1 to 12 
        mm(j)=""
    next
	
    mm(tmese)= " Selected "	

	X=""
    X=X & "<td  align='center' " & Parm & " valign='middle' class='boldTableCalendario'>" & chr(13)
	X=X & "<select name='GG_" & Nome & "'  size='1' class='new_inputText'>" 
	for j=1 to 31 
		if j=cint(Giorno) then 
			Sel= " Selected "	
		else
			Sel = " "
		end if 
		X=X & "<option value=" & right("0" & J,2) & Sel & ">" & right("0" & J,2) & "</option>"
	next 
	X=X & "</select> "  & chr(13)
	X=X & "<select name='MM_" & Nome & "' size='1' class='new_inputText'>" 
    X=X & "<option value='01'" & mm(1) & ">Gennaio</option>" 
    X=X & "<option value='02'" & mm(2) & ">Febbraio</option>" 
    X=X & "<option value='03'" & mm(3) & ">Marzo</option> "
    X=X & "<option value='04'" & mm(4) & ">Aprile</option> "
    X=X & "<option value='05'" & mm(5) & ">Maggio</option> "
    X=X & "<option value='06'" & mm(6) & ">Giugno</option> "
    X=X & "<option value='07'" & mm(7) & ">Luglio</option> "
    X=X & "<option value='08'" & mm(8) & ">Agosto</option> "
    X=X & "<option value='09'" & mm(9) & ">Settembre</option>" 
    X=X & "<option value='10'" & mm(10) & ">Ottobre</option> "
    X=X & "<option value='11'" & mm(11) & ">Novembre</option> "
    X=X & "<option value='12'" & mm(12) & ">Dicembre</option> "
    X=X & "</select> " & chr(13)
	X=X & "<select name='AA_" & Nome & "'  size='1' class='new_inputText'>" 
	
	AnnoInizio=Anno
	If cdbl(Anno)<0 then
	   Anno=Cdbl(Anno)*(-1)
	   AnnoInizio=Anno
	   if AnnoInizio=2029 then 
	      AnnoInizio=year(date())
	   end if 
	end if 
	If AnnoInizio>2006 then 
	   AnnoInizio=2006
	end if 
	for j=cint(AnnoInizio)-1 to 2029 
		if j=cint(Anno) then 
			Sel= " Selected "	
		else
			Sel = " "
		end if 
		X=X & "<option value=" & j & " " & Sel & ">" & j & "</option>" 
	next 
    X=X & "</select> " & chr(13)

	OptionData=X & "</TD>" & chr(13)
End Function 


Function OptionDataAnno(Nome,Anno,tmese,Giorno,Vuota)
Dim X,J,mm(13),sel


    for j=1 to 12 
        mm(j)=""
    next
	
    mm(tmese)= " Selected "	

	X=""
	X=X & "<select name='GG_" & Nome & "'  size='1' class='new_inputText'>" 
	if vuota=1 then 
	   X=X & "<option value=-1>  </option>"
	end if 
	for j=1 to 31 
		if j=cint(Giorno) then 
			Sel= " Selected "	
		else
			Sel = " "
		end if 
		X=X & "<option value=" & right("0" & J,2) & Sel & ">" & right("0" & J,2) & "</option>"
	next 
	X=X & "</select> "  & chr(13)
	X=X & "<select name='MM_" & Nome & "' size='1' class='new_inputText'>" 
	if vuota=1 then
	   X=X & "<option value='-1'>  </option>" 
	end if 
    X=X & "<option value='01'" & mm(1) & ">Gennaio</option>" 
    X=X & "<option value='02'" & mm(2) & ">Febbraio</option>" 
    X=X & "<option value='03'" & mm(3) & ">Marzo</option> "
    X=X & "<option value='04'" & mm(4) & ">Aprile</option> "
    X=X & "<option value='05'" & mm(5) & ">Maggio</option> "
    X=X & "<option value='06'" & mm(6) & ">Giugno</option> "
    X=X & "<option value='07'" & mm(7) & ">Luglio</option> "
    X=X & "<option value='08'" & mm(8) & ">Agosto</option> "
    X=X & "<option value='09'" & mm(9) & ">Settembre</option>" 
    X=X & "<option value='10'" & mm(10) & ">Ottobre</option> "
    X=X & "<option value='11'" & mm(11) & ">Novembre</option> "
    X=X & "<option value='12'" & mm(12) & ">Dicembre</option> "
    X=X & "</select> " & chr(13)
	X=X & "<input name='AA_" & Nome & "' size='4' type='text' class='new_inputText' id='AA_" & Nome & "' value='" & Anno & "'>" 
	

	OptionDataAnno=X & chr(13)
End Function 

Function OptionRangeNumLis(Nome,DaNum,ANum,Valore,ListaNumeri,ZeroBlank,PrimoZero)
Dim J,V,T 
'crea l'option DalNum AlNum mettendo i testi prelevati 
	on error resume next
	v=split(ListaNumeri,";")
	
	X=""
    X=X & "<select name='" & Nome & "'  size='1' class='new_inputText'>" 
	t=0
	if PrimoZero=1 then 
	   X=X & "<option value=""" & "0" & """" & " Selected " & "> </option>"
	end if
	for j=cdbl(DaNum) to Cdbl(Anum) 
		if j=cint(Valore) then 
			Sel= " Selected "	
		else
			Sel = " "
		end if 
		if ListaNumeri<>"" then 
		   OutJ=v(T)
		   if OutJ="" then 
		      OutJ=J
		   end if 
		else
		   OutJ=j
		   if j=0 and ZeroBlank=1 then 
		      OutJ=""
		   end if 
		end if 
		T=T+1
		X=X & "<option value=""" & J & """" & Sel & ">" & OutJ & "</option>"
	next 
	OptionRangeNumLis=X & "</select> "  & chr(13)

	err.clear
	
End Function 


Function OptionListaValoriConId(Nome,Id,ListaValori,ValoreBase)
Dim J,V,T 


	on error resume next
	v=split(ListaValori,";")
	
	X=""
   X=X & "<select Id='" & Id & "'" & " name='" & Nome & "'  size='1' class='form-control'>" 
	t=0

	for j=0 to ubound(v)-1 step 2
	   Testo=v(j)
		Valore=V(j+1)
		if Cstr(Valore)=Cstr(ValoreBase) then 
			Sel= " Selected "	
		else
			Sel = " "
		end if 
		T=T+1
		X=X & "<option value=""" & Valore & """" & Sel & ">" & Testo & "</option>"
	next 
	OptionListaValoriConId=X & "</select> "  & chr(13)

	err.clear
	
End Function 


Function OptionListaValoriConIdClasse(Nome,Id,ListaValori,ValoreBase,Classe)
Dim J,V,T 


	on error resume next
	v=split(ListaValori,";")
	
	
	X=""
	X=X & "<select Id='" & Id & "'" & " name='" & Nome & "' " & Classe & " >" 
	t=0

	for j=0 to ubound(v)-1 step 2
	   Testo=v(j)
		Valore=V(j+1)
		if Valore=ValoreBase then 
			Sel= " Selected "	
		else
			Sel = " "
		end if 
		T=T+1
		X=X & "<option value=""" & Valore & """" & Sel & ">" & Testo & "</option>"
	next 
	OptionListaValoriConIdClasse=X & "</select> "  & chr(13)

	err.clear
	
End Function 


Function OptionRangeNumChange(Nome,DaNum,ANum,Valore,Change)
Dim J

	X=""
    X=X & "<select name='" & Nome & "' id='" & Nome & "' onchange='" & Change & "' size='1' class='new_inputText'>" 
	for j=cdbl(DaNum) to Cdbl(Anum) 
		if j=cint(Valore) then 
			Sel= " Selected "	
		else
			Sel = " "
		end if 
		X=X & "<option value=""" & J & """" & Sel & ">" & J & "</option>"
	next 
	OptionRangeNumChange=X & "</select> "  & chr(13)

	
End Function 

Function OptionRangeNumDecChange(Nome,DaNum,ANum,Stp,Valore,Change)
Dim J
	X=""
    X=X & "<select name='" & Nome & "' id='" & Nome & "' onchange='" & Change & "' size='1' class='new_inputText'>" 
	for j=cdbl(DaNum) to Cdbl(Anum) step Stp
		if j = cdbl(Valore) then 
			Sel= " Selected "	
		else
			Sel = " "
		end if 
		X=X & "<option value=""" & J & """" & Sel & ">" & J & "</option>"
	next 
	OptionRangeNumDecChange=X & "</select> "  & chr(13)	
End Function 

Function decripta(strText)
temp_text = Trim(strText) 							
len_const_cript = Len(DEFAULT_TEXT_CRIPTAZIONE)										
array_alf = Array(" ","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","A","B","C","D","E","F","G","H","I","J","L","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","0","1","2","3","4","5","6","7","8","9",".","_","-")
number_criptazione = 0
For ik = 1 To CDbl(len_const_cript)													
	char_const_crip = Mid(DEFAULT_TEXT_CRIPTAZIONE, ik, 1)							
	For i = LBound(array_alf) To UBound(array_alf)									
		If CStr(char_const_crip) = CStr(array_alf(i)) Then							
			char_cript_one = i														
		End If
	Next
	number_criptazione = CDbl(number_criptazione) + CDbl(char_cript_one)			
Next
If Right(temp_text, 1) = "|" Then													
	temp_text = Left(temp_text, Len(temp_text) - 1)									
End If
If Left(temp_text, 1) = "|" Then													
	temp_text = Right(temp_text, Len(temp_text) - 1)								
End If
split_text = Split(temp_text, "|")													
number_pos = 2																		
temp_text_decrip = ""
Do While Not number_pos > UBound(split_text)										
	char_ascii = CDbl(split_text(number_pos)) / CDbl(number_criptazione)			
	char_ascii = CDbl(char_ascii) / CDbl(split_text(number_pos - 1))				
	char_decript = Chr(CDbl(char_ascii))											
	temp_text_decrip = temp_text_decrip & char_decript								
	number_pos = number_pos + 3														
Loop
temp_text_decrip = StrReverse(temp_text_decrip)										
decripta = temp_text_decrip															
End Function																			

Function cripta(strText)																
temp_text = Trim(strText) 							
temp_text = StrReverse(temp_text)													
len_text = Len(temp_text) : len_const_cript = Len(DEFAULT_TEXT_CRIPTAZIONE)			
array_alf = Array(" ","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","A","B","C","D","E","F","G","H","I","J","L","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","0","1","2","3","4","5","6","7","8","9",".","_","-")
number_criptazione = 0
For ik = 1 To CDbl(len_const_cript)													
	char_const_crip = Mid(DEFAULT_TEXT_CRIPTAZIONE, ik, 1)							
	For iLp = LBound(array_alf) To UBound(array_alf)									
		If CStr(char_const_crip) = CStr(array_alf(iLp)) Then							
			char_cript_one = iLp														
		End If
	Next
	number_criptazione = CDbl(number_criptazione) + CDbl(char_cript_one)			
Next
text_criptazione = ""
For ikj = 1 To CDbl(len_text)														
	char_text = Mid(temp_text, ikj, 1)												
	ascii_text = ASC(char_text) : len_ascii_text = Len(ascii_text)					
	char_text_crip = CDbl(ascii_text) * CDbl(number_criptazione)					
	char_text_crip = CDbl(char_text_crip) * CDbl(len_ascii_text)					
	len_char_text_crip = Len(char_text_crip)										
	text_criptazione = text_criptazione & len_char_text_crip & "|"					
	text_criptazione = text_criptazione & len_ascii_text & "|" 						
	text_criptazione = text_criptazione & char_text_crip & "|"						
Next
text_criptazione = Trim(text_criptazione)											
If Right(text_criptazione, 1) = "|" Then											
	text_criptazione = Left(text_criptazione, Len(text_criptazione) - 1)			
End If
If Left(text_criptazione, 1) = "|" Then												
	text_criptazione = Right(text_criptazione, Len(text_criptazione) - 1)			
End If
cripta = text_criptazione															
End Function																			


Function ListaRangeChangeSize (FromNum,ToNum,StepNum,Name,CodValue,FlagFormat,FlagVuoto,Change,Larghezza)
'FromNum = da quale numero 
'ToNum   = a quale numoero
'StepNum = incremento 
'Name = nome del campo da assegnare
'CodValue = valore del codice
'FlagFormat = colonna della descrizione per eventuale ofrmattazione 
'Flagvuoto=1 indica che devo aggiungere la riga vuota
'Change=funzione da eseguire in caso di change
'Campo=se presente crea un campo hidden 
'Larghezza imposta la size della lista
    if StepNum="" then 
	   StepNum=1
	end if
	if IsNumeric(StepNum)=false then 
	   StepNum=1
	end if 

    if FromNum="" then 
	   FromNum=0
	end if
	if IsNumeric(FromNum)=false then 
	   FromNum=0
	end if 
    if ToNum="" then 
	   ToNum=0
	end if
	if IsNumeric(ToNum)=false then 
	   ToNum=0
	end if 
	

	Stile=""
	 
	If len(larghezza)>0 then 
		if IsNumeric(Larghezza) then 
		   Stile= "style='width: " & Larghezza & "px'"
		end if 
	end if 
	 
	 
	if Change="" then 
	   Response.write "<SELECT " & Stile & " class=select name='" & Name & "' size=1  id='" & Name & "' >"   
	else
	   Response.write "<SELECT " & Stile & " onchange=""" & Change & """ class=select name='" & Name & "' size=1  id='" & Name & "' >"   
	end if 

	if FlagVuoto=1 then 
		If Trim(CodValue)="" then	
			Response.write "<option value=-1 selected > </option>"
		else
			Response.write "<option value=-1 > </option>"
		end if
	end if 
    
    TT=cdbl(FromNum)
    do while TT <= Cdbl(ToNum) 
	    Selected=""
		if TT=CodValue then 
		   Selected = " Selected "
		end if 
		
		if FlagFormat="S" then 
		   DD=formatNumber(TT)
		else
           DD=TT		
		end if 
		
		Response.write "<option value='" & tt & "' " & Selected &  " >" & server.htmlencode(DD) & "</option>"
        TT=TT+cdbl(StepNum)
	loop

    Response.write "</SELECT>"
    
    'MySet.Close
	'Set MySet = nothing
    
end function

Function ListaRangeChangeSizeClasse (FromNum,ToNum,StepNum,Name,CodValue,FlagFormat,FlagVuoto,Change,Larghezza,Classe)
'FromNum = da quale numero 
'ToNum   = a quale numoero
'StepNum = incremento 
'Name = nome del campo da assegnare
'CodValue = valore del codice
'FlagFormat = colonna della descrizione per eventuale ofrmattazione 
'Flagvuoto=1 indica che devo aggiungere la riga vuota
'Change=funzione da eseguire in caso di change
'Campo=se presente crea un campo hidden 
'Larghezza imposta la size della lista
'
    if StepNum="" then 
	   StepNum=1
	end if
	if IsNumeric(StepNum)=false then 
	   StepNum=1
	end if 

    if FromNum="" then 
	   FromNum=0
	end if
	if IsNumeric(FromNum)=false then 
	   FromNum=0
	end if 
    if ToNum="" then 
	   ToNum=0
	end if
	if IsNumeric(ToNum)=false then 
	   ToNum=0
	end if 
	

	Stile=""
	 
	If len(larghezza)>0 then 
		if IsNumeric(Larghezza) then 
		   Stile= "style='width: " & Larghezza & "px'"
		end if 
	end if 
	 
	 
	if Change="" then 
	   Response.write "<SELECT " & Stile & " " & Classe & " name='" & Name & "' size=1  id='" & Name & "' >"   
	else
	   Response.write "<SELECT " & Stile & " " & Classe & " onchange=""" & Change & """ name='" & Name & "' size=1  id='" & Name & "' >"   
	end if 

	if FlagVuoto=1 then 
		If Trim(CodValue)="" then	
			Response.write "<option value=-1 selected > </option>"
		else
			Response.write "<option value=-1 > </option>"
		end if
	end if 
    
    TT=cdbl(FromNum)
    do while TT <= Cdbl(ToNum) 
	    Selected=""
		if TT=CodValue then 
		   Selected = " Selected "
		end if 
		
		if FlagFormat="S" then 
		   DD=formatNumber(TT)
		else
           DD=TT		
		end if 
		
		Response.write "<option value='" & tt & "' " & Selected &  " >" & server.htmlencode(DD) & "</option>"
        TT=TT+cdbl(StepNum)
	loop

    Response.write "</SELECT>"
    
    'MySet.Close
	'Set MySet = nothing
    
end function

Function ErroreDb(Descrizione) 
Dim x
	if instr(ucase(Descrizione),"DUPLICAT")>0 then 
	   x = "Occorrenza presente in archivio"
	else
       x = Descrizione
	end if 
    ErroreDb=x	
end function

Function CaricaLog(x)
Dim MyQ
   on error resume next
   MyQ=MyQ & " Insert Into TraceLog (DataLog,TipoUtente,IdUtente,TipoErrore,Procedura,Descerrore,AddInfo) "
   MyQ=MyQ & " values (getdate(),'Utente Interno',0,'Errore','','" & apici(x) & "','')" 
   ConnMsde.execute MyQ
   err.clear
end function

Function AddTag(StIn,TagIn,Valore)
Dim X
	X=StIn
	'response.write x & "<br>"
	X=X & "<" & TagIn & ">"
	'response.write x & "<br>"
	X=X &  Valore 
	X=X & "</" & TagIn & ">"
	
	AddTag=X

End Function

Function EstraiTag(StIn,TagIn)
Dim X,Inizio
	X=""
	Ptr1=instr(StIn,"<"& TagIn & ">")
	if Ptr1>0 then
       Inizio=Ptr1+len(TagIn)+2	
	   Ptr2=instr(StIn,"</"& TagIn & ">")
	   if Ptr2>0 then 
	      X=Mid(StIn,Inizio,Ptr2-Inizio)
	   end if 
	end if
	
	EstraiTag=X

End Function

Function UpdateTag(Add,Xml,TagIn,StIn)
' Add   - Se true, se il Tag non esiste lo aggiunge
' Xml   - File xml da aggiornare
' TagIn - Tag da modificare
' StIn  - Nuovo valore
Dim X,Inizio,input,output
 Ptr0=instr(Xml,"<"& TagIn & ">")
 if (Ptr0 = 0 and Add) then
	xml = AddTag(xml,TagIn,StIn)
	UpdateTag = xml
 else
	 X=""
	 Ptr1=instr(Xml,"<"& TagIn & ">")
	 if Ptr1>0 then
		   Inizio=Ptr1+len(TagIn)+2 
		Ptr2=instr(Xml,"</"& TagIn & ">")
		if Ptr2>0 then 
		   X=Mid(Xml,Inizio,Ptr2-Inizio)
		end if 
	 end if
	 input  = "<"& TagIn & ">" & X & "</"& TagIn & ">"
	 output = "<"& TagIn & ">" & StIn & "</"& TagIn & ">"
	 UpdateTag = replace(Xml,input,output)
 end if
End Function

Function CalcolaDiff(D1,D2,MM,GG)
Dim DT
	Dt=D1
	DtIn=DateSerial(mid(dt,1,4),Mid(dt,5,2),mid(dt,7,2))
	
	Dt=D2
	DtFi=DateSerial(mid(dt,1,4),Mid(dt,5,2),mid(dt,7,2))
	
	MM=DateDiff("m",DtIn,DtFi) 
	Dt=DateAdd("m",MM,DtIn)
	GG=DateDiff("g",Dt,DtIn) 

End function

Function CalcolaGiorniDiff(D1,D2,GG)
Dim DT
	Dt=D1
	DtIn=DateSerial(mid(dt,1,4),Mid(dt,5,2),mid(dt,7,2))
	
	Dt=D2
	DtFi=DateSerial(mid(dt,1,4),Mid(dt,5,2),mid(dt,7,2))
	
	GG=DateDiff("d",DtIn,DtFi) 

End function

Function ColoreRiga(Tipo,CC)
   ColoreRiga=""
	BgColor="white" 
	if Tipo="APP" then  
	   BgColor="#99ff99"
	elseif Tipo="ATT" then 
	   BgColor="#ffff99"
	elseif Tipo="ANN" then 
	   BgColor="red"
	elseif Tipo="LAV" then 
	   BgColor="coral"
	end if 
	CC=BgColor
	
End Function 

Function SetSpace()
   SetSpace="&nbsp;"
end function 

Function CreaZipFile(PathZip,PathFile,QQ,NomeCampoFile,QQFile)
'PathZip= il path assoluto compreso del nome del file zip da creare
'PathFile= il path assoluto della cartella dove si trovano i file da zippare compreso di "\" finale (esempio: C:\pippo\)
'QQ= Query da lanciare per recuperare i nomi dei file da zippare
'NomeCampoFile= Il nome del campo della tabella che contiene il nome del file da zippare
'QQFile = In alternativa alla query può essere passata una stringa di file da zippare separati da ";"
	Dim objZip
	Err.clear
	ErroreZip=false
	
	If Len(Trim(QQ))<3 Then
		If Len(Trim(QQFile))>3 Then
			
			Set objZip = Server.CreateObject("XStandard.Zip.1")
		
			KeyCode=split(QQFile,";")
			For J=lbound(Keycode) to Ubound(KeyCode)
				If ErroreZip=false Then
					KK=Trim(Keycode(J))
					objZip.Pack PathFile & KK,PathZip,,,9

					If objZip.ErrorCode>0 Then
						ErroreZip=true
						MsgErrore=objZip.ErrorDescription
					End If
				End If
			Next
			Set objZip = Nothing
		Else
			ErroreZip=true
			MsgErrore="Nei parametri manca sia la query sia la stringa con i nomi dei file da zippare."
		End If
	Else
		Set RsQ = Server.CreateObject("ADODB.Recordset")
		
		RsQ.CursorLocation = 3 
		RsQ.Open QQ, ConnMsde
	'response.write Pathzip & "<br>" & PathFile & "<br>" & QQ & "<br>" & NomeCampoFile
		if err.number<>0 then
			MsgErrore = Err.Description 
			ErroreZip=true
		ElseIf RsQ.EOF then
			MsgErrore= "Nessun dettaglio in archivio"
			ErroreZip=true
		Else
		
			Set objZip = Server.CreateObject("XStandard.Zip.1")
			
			Do While Not RsQ.EOF and ErroreZip=false
				objZip.Pack PathFile & RsQ(NomeCampoFile),PathZip,,,9

				If objZip.ErrorCode>0 Then
					ErroreZip=true
					MsgErrore=objZip.ErrorDescription
				End If
				RsQ.MoveNext
			Loop
			RsQ.Close
			Set objZip = Nothing
		End If
		Set RsQ = Nothing
	End If
	If ErroreZip=true Then
		CreaZipFile=MsgErrore
	Else
		CreaZipFile=""
	End If
	
	Err.Clear
End Function


Function ArrontondaEccesso(Num)
	'se ha una virgola vuol dire che è decimale 
	seDecimale = Instr(  Num,"," )
	'se c'è la virgola....
	if seDecimale <> 0 then
		divisore = ","
		'divido gli elementi separati dalla virgola e ne ottengo un array
		elementi = Split( Num,divisore)
		'elemento dell'array a sinistra della virgola (intero) 
		intero = elementi(0)
		'elemento dell'array a destra della virgola (decimale)   
		decimale= elementi(1)
		If decimale>0 Then
			risultato = intero + 1
		Else
			risultato = Num
		End If
	else
		'se non era un decimale non c'è bisogno di arrotondare niente
		risultato = Num
	end if
	ArrontondaEccesso=risultato
	err.clear
End Function

Function EliminaFile(PathFile,QQ,NomeCampoFile,QQFile)
'PathFile= il path assoluto della cartella dove si trova il file o i files da eliminare compreso di "\" finale (esempio: C:\pippo\)
'QQ= Query da lanciare per recuperare i nomi dei file da eliminare
'NomeCampoFile= Il nome del campo della tabella che contiene il nome del file da eliminare
'QQFile = In alternativa alla query può essere passata una stringa di file da eliminare separati da ";"

	Set Fs = Server.CreateObject ("Scripting.FileSystemObject")
	Esito=""

	If Len(Trim(QQ))<3 Then
		If Len(Trim(QQFile))>3 Then
			'Elimino i file della lista che mi è stata passata
			KeyCode=split(QQFile,";")
			For J=lbound(Keycode) to Ubound(KeyCode)
				KK=Trim(Keycode(J))
				If Fs.FileExists(PathFile & KK) Then
					Fs.DeleteFile PathFile & KK,true
				Else
					Esito="File non trovato. Path completo del File: " & PathFile & KK
					TraceDb "Errore","Function EliminaFile",Esito,QQFile
				End If
			Next
		Else
			Esito="Nei parametri manca sia la query sia la stringa con i nomi dei file da eliminare."
			TraceDb "Errore","Function EliminaFile. ",Esito,""
		End If
	Else
		'Recupero i nomi dei file da eliminare e li elimino
		
		Set RsQ = Server.CreateObject("ADODB.Recordset")
		
		RsQ.CursorLocation = 3 
		RsQ.Open QQ, ConnMsde

		if err.number<>0 then
			Esito=Err.Description
		ElseIf RsQ.EOF then
			Esito="Nessun dettaglio in archivio."
			RsQ.Close
		Else
			Do While Not RsQ.EOF 
				NomeFileDel=""
				NomeFileDel=RsQ(NomeCampoFile)
				
				If Fs.FileExists(PathFile & NomeFileDel) Then
					Fs.DeleteFile PathFile & NomeFileDel,true
				Else
					Esito="File non trovato. Path completo del File: " & PathFile & NomeFileDel
					TraceDb "Errore","Function EliminaFile. Pagina: " & NomePagina,Esito,QQ
				End If
				RsQ.MoveNext
			Loop
			RsQ.Close
		End If
		Set RsQ = Nothing
	End If

	Set Fs = Nothing

	EliminaFile=Esito

	Err.Clear
End Function

function UpCase(StrUpCase)
	dim frase
	
	frase = Split(StrUpCase, " ")

	for i = 0 to ubound(frase)
		If Trim(frase(i))<>"" Then
			frase(i) = UCase ( Mid(frase(i), 1, 1) ) & LCase ( Mid(frase(i), 2) )
		End If
	next

	UpCase = Join(frase, " ")

End Function

Function Pulisci(t)
Dim Rv, c 
	Rv=t 
	'utf 8 
	Rv=Replace(Rv,"Ã " ,"a'"   , 1, -1, 0)
	Rv=Replace(Rv,"Ã¨" ,"e'"   , 1, -1, 0)
	Rv=Replace(Rv,"Ã©" ,"e'"   , 1, -1, 0)
	Rv=Replace(Rv,"Ã¬" ,"i'"   , 1, -1, 0)
	Rv=Replace(Rv,"Ã²" ,"o'"   , 1, -1, 0)
	Rv=Replace(Rv,"Ã¹" ,"u'"   , 1, -1, 0)

	'ansi 
	Rv=Replace(Rv,"à" ,"a'"   , 1, -1, 0)
	Rv=Replace(Rv,"è" ,"e'"   , 1, -1, 0)
	Rv=Replace(Rv,"é" ,"e'"   , 1, -1, 0)
	Rv=Replace(Rv,"ì" ,"i'"   , 1, -1, 0)
	Rv=Replace(Rv,"ò" ,"o'"   , 1, -1, 0)
	Rv=Replace(Rv,"ù" ,"u'"   , 1, -1, 0)

	
	Rv=Replace(Rv,"°"  ," "    , 1, -1, 0)
	Rv=Replace(Rv,"Ï"  ,"I"    , 1, -1, 0)
	Rv=Replace(Rv,"¿"  ," "    , 1, -1, 0)
	Rv=Replace(Rv,"½"  ," "    , 1, -1, 0)
	Rv=Replace(Rv,"€"  ,"euro" , 1, -1, 0)
	Rv=Replace(Rv,"’"  ,"'"    , 1, -1, 0)
	Rv=Replace(Rv,"“"  ,""""   , 1, -1, 0)
	Rv=Replace(Rv,"”"  ,""""   , 1, -1, 0)
	Rv=Replace(Rv,"•"  ,"-"    , 1, -1, 0)
	Rv=Replace(Rv,"–"  ,"-"    , 1, -1, 0)
	
	'unicode
	Rv=Replace(Rv,"\u2019"  ,"'"  , 1, -1, 0)
	Rv=Replace(Rv,"\u00d9"  ,"u'" , 1, -1, 0)
	Rv=Replace(Rv,"\u00a0"  ," "  , 1, -1, 0)
	Rv=Replace(Rv,"\u201c"  ,"'"  , 1, -1, 0)
	Rv=Replace(Rv,"\u201d"  ,"'"  , 1, -1, 0)
	Rv=Replace(Rv,"\u00b0"  ,""   , 1, -1, 0)
	Rv=Replace(Rv,"\u00c0"  ,"a'" , 1, -1, 0)
	Rv=Replace(Rv,"\u2013"  ,"-"  , 1, -1, 0)
	
	
	Pulisci=Rv
End Function


Function PulisciNomeFile(t)
Dim Rv
	Rv=t 
	Rv=Replace(Rv,"~","_")
	Rv=Replace(Rv,"#","_")
	Rv=Replace(Rv,"%","_")
	Rv=Replace(Rv,"&","_")
	Rv=Replace(Rv,"*","_")
	Rv=Replace(Rv,"{","_")
	Rv=Replace(Rv,"}","_")
	Rv=Replace(Rv,"\","_")
	Rv=Replace(Rv,":","_")
	Rv=Replace(Rv,"<","_")
	Rv=Replace(Rv,">","_")
	Rv=Replace(Rv,"?","_")
	Rv=Replace(Rv,"/","_")
	Rv=Replace(Rv,"|","_")
	Rv=Replace(Rv,"“","_")
	Rv=Replace(Rv,"”","_")
	Rv=Replace(Rv,"""","_")
	Rv=Replace(Rv,"+","_")
	Rv=Replace(Rv,"–","-")
	
	Rv=Replace(Rv,"à","a'")
	Rv=Replace(Rv,"À","A'")
	Rv=Replace(Rv,"è","e'")
	Rv=Replace(Rv,"é","e'")
	Rv=Replace(Rv,"ì","i'")
	Rv=Replace(Rv,"ò","o'")
	Rv=Replace(Rv,"ù","u'")
	Rv=Replace(Rv,"&","")
	Rv=Replace(Rv,"°","")
	Rv=Replace(Rv,"Ï","I")
	Rv=Replace(Rv,"¿","")
	Rv=Replace(Rv,"½","")
	Rv=Replace(Rv,"€","euro")
	Rv=Replace(Rv,"’","")
	Rv=Replace(Rv,"“","")
	Rv=Replace(Rv,"”","")
	Rv=Replace(Rv,"•","")
	
	PulisciNomeFile=Rv
End Function

Function ArrXEccessoAlMezzoDecimale(NumIngresso)
'Arrontonda un numero decimale per eccesso al mezzo decimale più vicino
	
	Dim TmpNum,NumUscita,StrDec
	
	NumUscita = 0
	TmpNum    = 0
	StrDec    = ""
	
	TmpNum = Round(NumIngresso,2)
	
	If TmpNum<>NumIngresso Then
		If Instr(NumIngresso,",")>0 Then
		'response.write "<br> QUIIII:"& Mid(NumIngresso,1,Instr(NumIngresso,",")-1) & " <br>"
			StrDec="0" & Mid(NumIngresso,Instr(NumIngresso,","),20)
			
			If Cdbl(StrDec)<0.5  Then
				NumUscita=Cdbl(Mid(NumIngresso,1,Instr(NumIngresso,",")-1))+0.5
			ElseIf Cdbl(StrDec)>0.5 Then
				NumUscita=Cdbl(Mid(NumIngresso,1,Instr(NumIngresso,",")-1))+1
			End If
		End If
	Else
		NumUscita=NumIngresso
	End If
	
	ArrXEccessoAlMezzoDecimale=Cdbl(NumUscita)
End Function



Function CreateGUID2()
  Randomize Timer
  Dim tmpCounter,tmpGUID
  Const strValid = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  For tmpCounter = 1 To 15
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
	if (tmpCounter mod 5) = 0 and tmpCounter<15 then
		tmpGUID = tmpGUID & "-"
	end if
  Next
  CreateGUID2 = tmpGUID
End Function

Function GetEccezioni(Tipo,Ambito,Id)
	Dim Xml, sql
	sql = ""
	sql = sql & " Select * from Eccezioni "
	sql = sql & " where TipoEccezione = '" & ucase(Tipo) & "' "
	sql = sql & " and AmbitoEccezione = '" & ucase(Ambito) & "' "
	sql = sql & " and IdRiferimento   = '" & ucase(Id) & "' "
	Xml = LeggiCampo(sql,"DatiXmlOut")
	GetEccezioni = Xml
End Function
%>
