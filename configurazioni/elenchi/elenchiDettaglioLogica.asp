<%
NameLoaded=NameLoaded & "DescElenco,TE"
MessageFromCaller = ""

IdElenco=0
if FirstLoad then 
   IdElenco   = "0" & Session("swap_IdElenco")
   if Cdbl(IdElenco)=0 then 
      IdElenco = cdbl("0" & getValueOfDic(Pagedic,"IdElenco"))
   end if 
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
else
   IdElenco      = "0" & getValueOfDic(Pagedic,"IdElenco")
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
end if 
if cdbl(IdElenco)=0 then 
   response.redirect "\"
end if 

flagModElenco = LeggiCampo("Select * from elenco where IdElenco=" & IdElenco,"FlagModificabile")
DescElenco    = LeggiCampo("Select * from elenco where IdElenco=" & IdElenco,"DescElenco") 

on error resume next 
if Oper="INS" then 
    Session("TimeStamp")=TimePage
	ConnMsde.execute "Update ElencoValore set FlagMod=0 where IdElenco=" & IdElenco
	datiElenco=request("DescElenco0")
	arDati=split(datiElenco,",")
	Seq=0
	for j=lbound(arDati) to uBound(arDati)
	    if ArDati(j)<>"" then 
			Seq=Seq+1
			Id="0" & LeggiCampo("select * from ElencoValore Where Idelenco=" & IdElenco & " and Sequenza = "  & Seq,"IdValoreElenco")
			id=Cdbl(Id)
			if Id=0 then 
			   qUpd = ""
			   qUpd = qUpd & " Insert into ElencoValore(IdElenco,ValoreElenco,Sequenza,FlagMod)"
			   qUpd = qUpd & " values(" & IdElenco & ",'" & apici(ArDati(j)) & "'," & Seq & ",1)"
			else
			   qUpd = ""
			   qUpd = qUpd & " update ElencoValore set"
			   qUpd = qUpd & " FlagMod=1,ValoreElenco='" & apici(ArDati(j)) & "'"
			   qUpd = qUpd & " where IdElenco=" & Idelenco
			   qUpd = qUpd & " and Sequenza=" & Seq
			end if 
			ConnMsde.execute qUpd
			If Err.Number <> 0 Then 
				MsgErrore = ErroreDb(Err.description)
			else
				MessageFromCaller="Aggiornamento eseguito"
			End If
		end if 
	next 
	'rimuovo non movimentati
	ConnMsde.execute "delete From ElencoValore Where Idelenco=" & IdElenco & " and FlagMod = 0"

End if
   if MessageFromCaller<>"" then 
      Session("Swap_MessageFromCaller")=MessageFromCaller
	  response.redirect virtualpath & PaginaReturn
	  response.end 
   end if 
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"Idelenco"     ,IdElenco)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  
  xx=setCurrent(NomePagina,livelloPagina) 
%>