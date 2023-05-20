<%
IdTabella=""
idTabellaDesc=""
IdTabellaKeyInt=0
IdTabellaKeyString=""
if FirstLoad then 
   IdTabella          = Session("swap_IdTabella")
   IdTabellaDesc      = Session("swap_IdTabellaDesc")
   IdTabellaKeyInt    = cdbl("0" & Session("swap_IdTabellaKeyInt"))
   IdTabellaKeyString = Session("swap_IdTabellaKeyString")
   OperAmmesse        = Session("swap_OperAmmesse")
   if idTabella="" then 
      IdTabella = getValueOfDic(Pagedic,"IdTabella")
   end if 
   if Cdbl(IdTabellaKeyInt)=0 then 
      IdTabellaKeyInt = cdbl("0" & getValueOfDic(Pagedic,"IdTabellaKeyInt"))
   end if
   if IdTabellaKeyString="" then 
      IdTabellaKeyString = getValueOfDic(Pagedic,"IdTabellaKeyString")
   end if    
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
else
   IdTabella          = getValueOfDic(Pagedic,"IdTabella")
   IdTabellaDesc      = getValueOfDic(Pagedic,"IdTabellaDesc")
   IdTabellaKeyInt    = cdbl("0" & getValueOfDic(Pagedic,"IdTabellaKeyInt"))
   IdTabellaKeyString = getValueOfDic(Pagedic,"IdTabellaKeyString")
   OperAmmesse        = getValueOfDic(Pagedic,"OperAmmesse")
   PaginaReturn       = getValueOfDic(Pagedic,"PaginaReturn")
end if 

if IdTabella="" then 
   response.redirect PaginaReturn
   response.end
end if 
ShowValidDate=false
if IdTabellaDesc="" then
   If ucase(IdTabella)="PRODOTTO" then 
      IdTabellaDesc = "Documenti per prodotto"
	  tmpString = "Select * from Prodotto Where idProdotto=" & IdTabellaKeyInt
	  'response.write tmpString 
	  'response.end
	  tmpString = LeggiCampo(tmpString,"DescProdotto")
	  if tmpString<>"" then 
	     IdTabellaDesc = IdTabellaDesc & " : " & tmpString
	  end if 
   end if
   If ucase(IdTabella)="PRODOTTO_COND" then 
      IdTabellaDesc = "Condizioni di Contratto per il prodotto"
	  tmpString = "Select * from Prodotto Where idProdotto=" & IdTabellaKeyInt
	  'response.write tmpString 
	  'response.end
	  tmpString = LeggiCampo(tmpString,"DescProdotto")
	  if tmpString<>"" then 
	     IdTabellaDesc = IdTabellaDesc & " : " & tmpString
	  end if 
   end if   
end if 
If ucase(IdTabella)="PRODOTTO_COND" then 
   ShowValidDate=true
end if 
DescElenco = IdTabellaDesc

OperAmmesse="IUD"
on error resume next 
if Oper="CALL_INS" or Oper="CALL_UPD" then 
   xx=RemoveSwap()
   itemId = Request("ItemToRemove") 
  
   'response.end 
   Session("swap_IdTabella")          = IdTabella
   Session("swap_IdTabellaKeyInt")    = IdTabellaKeyInt
   Session("swap_IdTabellaKeyString") = IdTabellaKeyString
   Session("swap_IdUpload")           = Cdbl("0" & itemId)
   Session("swap_PaginaReturn")  = "configurazioni/documenti/Documentielenco.asp"
   response.redirect virtualPath & "configurazioni/documenti/DocumentoUpload.asp"
   response.end 
end if

if Oper="DEL" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	MyQ = "" 
	MyQ = MyQ & " delete from Upload "
	MyQ = MyQ & " where IdUpload = " & KK
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else
	    FlagUpdLista=true
	End If
	DescIn=""
End if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdTabella"          ,IdTabella)
  xx=setValueOfDic(Pagedic,"IdTabellaDesc"      ,IdTabellaDesc)
  xx=setValueOfDic(Pagedic,"IdTabellaKeyInt"    ,IdTabellaKeyInt)
  xx=setValueOfDic(Pagedic,"IdTabellaKeyString" ,IdTabellaKeyString)
  xx=setValueOfDic(Pagedic,"OperAmmesse"        ,OperAmmesse)
  xx=setValueOfDic(Pagedic,"PaginaReturn"       ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

%>