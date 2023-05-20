<%
NameLoaded=NameLoaded & "DataInizio,DTO"
NameLoaded=NameLoaded & ";PercImposta,FLQ"

IdTrattamentoFiscale=0
if FirstLoad then 
   IdTrattamentoFiscale   = "0" & Session("swap_IdTrattamentoFiscale")
   if Cdbl(IdTrattamentoFiscale)=0 then 
      IdTrattamentoFiscale = cdbl("0" & getValueOfDic(Pagedic,"IdTrattamentoFiscale"))
   end if 
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
else
   IdTrattamentoFiscale = "0" & getValueOfDic(Pagedic,"IdTrattamentoFiscale")
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
end if 
if cdbl(IdTrattamentoFiscale)=0 then 
   response.redirect "\"
end if 
IdTrattamentoFiscale    = cdbl(IdTrattamentoFiscale)
DescDettaglio    = LeggiCampo("Select * from TrattamentoFiscale where IdTrattamentoFiscale=" & IdTrattamentoFiscale,"DescTrattamentoFiscale") 

on error resume next 
FlagUpdRiferimento=false 

if Oper="INS" then 
    Session("TimeStamp")=TimePage
	KK="0"
	DataInizio  = Request("DataInizio"  & KK)
	DataInizio  = DataStringa(DataInizio)
	IdRegione   = Request("IdRegione"   & KK)
	IdRegione   = VuotoNoLista(IdRegione)
	IdProvincia = Request("IdProvincia" & KK)
	IdProvincia = VuotoNoLista(IdProvincia)
	PercImposta = TestNumeroPos(Request("PercImposta" & KK))
	if Cdbl(IdTrattamentoFiscale)>0 then 
		MyQ = "" 
		MyQ = MyQ & " Insert into TrattamentoFiscaleStorico ("
		MyQ = MyQ & " IdTrattamentoFiscale,IdRegione,IdProvincia,DataInizio,PercImposta"
		MyQ = MyQ & ") values ("			
		MyQ = MyQ & "  " & IdTrattamentoFiscale 
		MyQ = MyQ & ",'" & apici(IdRegione) & "'"
		MyQ = MyQ & ",'" & apici(IdProvincia) & "'"
		MyQ = MyQ & ", " & StoNum(DataInizio)
		MyQ = MyQ & ", " & StoNum(PercImposta)
		MyQ = MyQ & ")"

		ConnMsde.execute MyQ 
		If Err.Number <> 0 Then 
			MsgErrore = ErroreDb(Err.description)
		else
			FlagUpdRiferimento=true
			DescIn=""
		End If
	END if 
End if 
if Oper="UPD" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	DataInizio  = Request("DataInizio"  & KK)
	DataInizio  = DataStringa(DataInizio)
	IdRegione   = Request("IdRegione"   & KK)
	IdRegione   = VuotoNoLista(IdRegione)
	IdProvincia = Request("IdProvincia" & KK)
	IdProvincia = VuotoNoLista(IdProvincia)
	PercImposta = TestNumeroPos(Request("PercImposta" & KK))

	MyQ = "" 
	MyQ = MyQ & " update TrattamentoFiscaleStorico set"
	MyQ = MyQ & " IdRegione = '"    & apici(IdRegione)   & "'"
	MyQ = MyQ & ",IdProvincia = '"  & apici(IdProvincia) & "'"
	MyQ = MyQ & ",DataInizio =  "   & StoNum(DataInizio)
	MyQ = MyQ & ",PercImposta =  "  & StoNum(PercImposta)
	MyQ = MyQ & " where IdTrattamentoFiscaleStorico = " & KK
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else
	    FlagUpdRiferimento=true
	End If	
End if 
	
if Oper="DEL" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	MyQ = "" 
	MyQ = MyQ & " delete from TrattamentoFiscaleStorico "
	MyQ = MyQ & " where IdTrattamentoFiscaleStorico = " & KK
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else
	    FlagUpdRiferimento=true
	End If
	DescIn=""
End if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdTrattamentoFiscale"     ,IdTrattamentoFiscale)
  xx=setValueOfDic(Pagedic,"PaginaReturn"             ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

%>