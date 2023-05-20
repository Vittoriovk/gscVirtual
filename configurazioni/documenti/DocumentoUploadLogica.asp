<%
IdTabella=""
idTabellaDesc=""
IdUpload=0
IdTabellaKeyInt=0
IdTabellaKeyString=""
FlagFileUpload  ="S"
FlagDescEstesa  ="S"
FlagDataScadenza="S"
if FirstLoad then 
   IdTabella          = Session("swap_IdTabella")
   IdTabellaDesc      = Session("swap_IdTabellaDesc")
   IdTabellaKeyInt    = cdbl("0" & Session("swap_IdTabellaKeyInt"))
   IdUpload           = cdbl("0" & Session("swap_IdUpload"))
   IdTabellaKeyString = Session("swap_IdTabellaKeyString")
   OperAmmesse        = Session("swap_OperAmmesse")
   FlagFileUpload     = Session("swap_FlagFileUpload")
   FlagDescEstesa     = Session("swap_FlagDescEstesa")
   FlagDataScadenza   = Session("swap_FlagDataScadenza")   
   if idTabella="" then 
      IdTabella = getValueOfDic(Pagedic,"IdTabella")
   end if 
   if Cdbl(IdTabellaKeyInt)=0 then 
      IdTabellaKeyInt = cdbl("0" & getValueOfDic(Pagedic,"IdTabellaKeyInt"))
   end if
   if Cdbl(IdUpload)=0 then 
      IdUpload = cdbl("0" & getValueOfDic(Pagedic,"IdUpload"))
   end if
   if IdTabellaKeyString="" then 
      IdTabellaKeyString = getValueOfDic(Pagedic,"IdTabellaKeyString")
   end if    
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
   if FlagFileUpload="" then
      FlagFileUpload     = getValueOfDic(Pagedic,"FlagFileUpload")
   end if 
   if FlagDescEstesa="" then 
      FlagDescEstesa     = getValueOfDic(Pagedic,"FlagDescEstesa")
   end if 
   if FlagDataScadenza="" then 
      FlagDataScadenza   = getValueOfDic(Pagedic,"FlagDataScadenza")
   end if 
else
   IdTabella          = getValueOfDic(Pagedic,"IdTabella")
   IdTabellaDesc      = getValueOfDic(Pagedic,"IdTabellaDesc")
   IdTabellaKeyInt    = cdbl("0" & getValueOfDic(Pagedic,"IdTabellaKeyInt"))
   IdUpload           = cdbl("0" & getValueOfDic(Pagedic,"IdUpload"))
   IdTabellaKeyString = getValueOfDic(Pagedic,"IdTabellaKeyString")
   OperAmmesse        = getValueOfDic(Pagedic,"OperAmmesse")
   PaginaReturn       = getValueOfDic(Pagedic,"PaginaReturn")
   FlagFileUpload     = getValueOfDic(Pagedic,"FlagFileUpload")
   FlagDescEstesa     = getValueOfDic(Pagedic,"FlagDescEstesa")
   FlagDataScadenza   = getValueOfDic(Pagedic,"FlagDataScadenza")
end if 
if FlagFileUpload="" then 
   FlagFileUpload  ="S"
end if 
if FlagDescEstesa  ="" then 
   FlagDescEstesa  ="S"
end if 
if FlagDataScadenza="" then 
   FlagDataScadenza="S"
end if 
'response.write "QUII" & IdTabella
 ' response.end 
if IdTabella="" then 
   response.redirect PaginaReturn
   response.end
end if 
ShowValidoDal=false
ShowValidoAl =false
if IdTabellaDesc="" then
   If ucase(IdTabella)="PRODOTTO" then 
      IdTabellaDesc = "Documento per prodotto"
	  tmpString = "Select * from Prodotto Where idProdotto=" & IdTabellaKeyInt
	  tmpString = LeggiCampo(tmpString,"DescProdotto")
	  if tmpString<>"" then 
	     IdTabellaDesc = IdTabellaDesc & " : " & tmpString
	  end if 
   end if
   If ucase(IdTabella)="PRODOTTO_COND" then 
      IdTabellaDesc = "Condizioni di Contratto per il prodotto"
	  tmpString = "Select * from Prodotto Where idProdotto=" & IdTabellaKeyInt
	  tmpString = LeggiCampo(tmpString,"DescProdotto")
	  if tmpString<>"" then 
	     IdTabellaDesc = IdTabellaDesc & " : " & tmpString
	  end if 
   end if     
end if 
If ucase(IdTabella)="PRODOTTO_COND" then 
   ShowValidoDal=true
end if 

DescElenco = IdTabellaDesc

OperAmmesse="IUD"
on error resume next
If Oper="UPDATE" then
   if Cdbl(IdUpload)=0 then 
      Oper="INS"
   else
      Oper="UPD"
   end if 
end if 

FileCambiato=false
if Oper="INS" then 
    Session("TimeStamp")=TimePage
	KK=0
	MyQ = "" 
	MyQ = MyQ & " insert into Upload (IdTabella,IdTabellaKeyInt,IdTabellaKeyString,DataUpload"
	MyQ = MyQ & " ,TimeUpload,IdTipoDocumento,DescBreve,DescEstesa,NomeDocumento,PathDocumento,ValidoDal,ValidoAl) "
	MyQ = MyQ & " values ("
	MyQ = MyQ & " '" & Apici(IdTabella) & "'"
	MyQ = MyQ & ", " & IdTabellaKeyInt
	MyQ = MyQ & ",'" & Apici(IdTabellaKeyString) & "'"
	MyQ = MyQ & ",0,0,'','','','','',0,20991231)" 
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else
	    FileCambiato=true
	    IdUpload=GetTableIdentity("Upload")
		Oper="UPD"
	End If
end if
If FileCambiato=false and o.FileNameOf("FileIn0")<>"" then 
   FileCambiato=true
end if 

if Oper="UPD" and FileCambiato then 
   NomeFilFull = o.FileNameOf("FileIn0")
   sFileSplit = split(NomeFilFull, "\")
   sFile = sFileSplit(Ubound(sFileSplit))

   sFileWrite = "CX" & IdUpload & "_" & Year(Now()) & Month(Now()) & Day(Now()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now()) &  "_" & sFile
  
   o.FileInputName = "FileIn0"
   o.FileFullPath = PathBaseUpload  & sFileWrite
   o.save
   if o.Error <> ""  then
       MsgErrore= "Caricamento Fallito: " & o.Error & o.FileFullPath
   elseif err.number<>0 then
       MsgErrore= "Caricamento Errore : " & Err.Description
   else
       qUpd = ""
	   qUpd = qUpd & " Update Upload Set "
	   qUpd = qUpd & " NomeDocumento='" & apici(sFile) & "'"
	   qUpd = qUpd & ",PathDocumento='" & apici(sFileWrite) & "'"
	   qUpd = qUpd & " where IdUpload=" & IdUpload
	   
       ConnMsde.execute qUpd 
   end if 
end if 


if Oper="UPD" then 
    Session("TimeStamp")=TimePage
	KK=o.ValueOf("ItemToRemove")
    ValidoDal = "0" & DataStringa(o.ValueOf("ValidoDal0"))
    if isnumeric(ValidoDal)=false then 
       ValidoDal=0
    else
       ValidoDal=Cdbl(ValidoDal)
    end if 
    ValidoAl  = "0" & DataStringa(o.ValueOf("ValidoAl0"))
    if isnumeric(ValidoAl)=false then 
       ValidoAl=20991231
    else
       ValidoAl=Cdbl(ValidoAl)
    end if 
	if ValidoAl=0 then 
	   ValidoAl=20991231
	end if 
	   
	MyQ = "" 
	MyQ = MyQ & " update Upload set "
	MyQ = MyQ & " DataUpload = " & Dtos()
	MyQ = MyQ & ",TimeUpload = " & TimeToS()
	MyQ = MyQ & ",IdTipoDocumento = " & o.ValueOf("IdTipoDocumento0")
	MyQ = MyQ & ",DescBreve='"        & apici(o.ValueOf("DescBreve0")) & "'"
	MyQ = MyQ & ",DescEstesa='"       & apici(o.ValueOf("descEstesa0")) & "'"
	MyQ = MyQ & ",ValidoDal=" & ValidoDal
	MyQ = MyQ & ",ValidoAl="  & ValidoAl
	MyQ = MyQ & " where IdUpload = " & Idupload 
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else
	    FlagUpdLista=true
		response.redirect VirtualPath & PaginaReturn
		response.end
	End If
	DescIn=""
End if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdTabella"          ,IdTabella)
  xx=setValueOfDic(Pagedic,"IdTabellaDesc"      ,IdTabellaDesc)
  xx=setValueOfDic(Pagedic,"IdTabellaKeyInt"    ,IdTabellaKeyInt)
  xx=setValueOfDic(Pagedic,"IdUpload"           ,IdUpload)
  xx=setValueOfDic(Pagedic,"IdTabellaKeyString" ,IdTabellaKeyString)
  xx=setValueOfDic(Pagedic,"OperAmmesse"        ,OperAmmesse)
  xx=setValueOfDic(Pagedic,"PaginaReturn"       ,PaginaReturn)
  xx=setValueOfDic(Pagedic,"FlagFileUpload"     ,FlagFileUpload)
  xx=setValueOfDic(Pagedic,"FlagDescEstesa"     ,FlagDescEstesa)
  xx=setValueOfDic(Pagedic,"FlagDataScadenza"   ,FlagDataScadenza)

  xx=setCurrent(NomePagina,livelloPagina) 

%>