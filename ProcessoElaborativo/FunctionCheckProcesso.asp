<%
function checkProcessoElaborativo(IdProdotto,IdAccountFornitore,IdMovimento)
dim q,prElab,procedura  
On error resume next 
   'controllo se il prodotto prevede un processo 
   q = ""
   q = q & " select * from AccountProdotto"
   q = q & " Where IdAccount=" & IdAccountFornitore 
   q = q & " and IdProdotto=" & IdProdotto
   prElab = LeggiCampo(q,"IdProcessoElaborativo")
   if prElab<>"" then 
      urlServizio = getDomain() & VirtualPath & "ProcessoElaborativo/esegui_" & prElab & ".asp"
      opData = ""
      opData = "IdMovimento=" & IdMovimento 
      opResp=""
	  opType=""
	  xx=CallOtherPage(urlServizio,"POST",opData,opType,"",opResp)
   end if 
   

end function 
function checkProcessoElaborativoServizio(IdProdotto,IdAccountFornitore,IdServizioRichiesto)
dim q,prElab,procedura  
On error resume next 
   'controllo se il prodotto prevede un processo 
   q = ""
   q = q & " select * from AccountProdotto"
   q = q & " Where IdAccount=" & IdAccountFornitore 
   q = q & " and IdProdotto=" & IdProdotto
   prElab = LeggiCampo(q,"IdProcessoElaborativo")
   xx=writeTraceAttivita("checkProcessoElaborativoServizio " & q & " = " & prElab ,"IdServizioRichiesto",IdServizioRichiesto)  
   if prElab<>"" then 
      urlServizio = getDomain() & VirtualPath & "ProcessoElaborativo/eseguiServizio_" & prElab & ".asp"
      opData = ""
      opData = "IdServizioRichiesto=" & IdServizioRichiesto 
      opResp=""
	  opType=""
	  xx=CallOtherPage(urlServizio,"POST",opData,opType,"",opResp)
   end if 
   if err.number<>0 then 
      xx=writeTraceAttivita("checkProcessoElaborativoServizio " & err.description ,"IdServizioRichiesto",IdServizioRichiesto)     
   end if 
end function 

%>