<%
   'si aspetta in ingresso IdAttivita ed IdNumAttivita
   if CheckTimePageLoad()=false then 
      Oper=""
   end if 
   if mid(oper,1,7)="SDLDOCU" then
      actionSDL="SDLDOCUADD"
      if mid(Oper,1,len(actionSDL))=actionSDL then 
	     'response.write Oper & Request("ItemToRemove")
	     IdDocumento = Cdbl("0" & Request("ItemToRemove"))
		 if Cdbl(IdDocumento)>0 then
            xx=SetAttivitaDocumentoAdd(IdAttivita,IdNumAttivita,IdDocumento,"",1,1)
		 end if  
	  end if 

      actionSDL="SDLDOCUOK_"
      if mid(Oper,1,len(actionSDL))=actionSDL then 
	     IdAttivitaDocumento = Cdbl("0" & Mid(Oper,len(actionSDL)+1))
		 if Cdbl(IdAttivitaDocumento)>0 then
            xx=SetAttivitaDocumentoValido(IdAttivitaDocumento,"")
		 end if  
	  end if 
      actionSDL="SDLDOCUKO_"
      if mid(Oper,1,len(actionSDL))=actionSDL then 
	     IdAttivitaDocumento = Cdbl("0" & Mid(Oper,len(actionSDL)+1))
		 if Cdbl(IdAttivitaDocumento)>0 then
		    mt=Request("ItemToModify")
            xx=SetAttivitaDocumentoNonValido(IdAttivitaDocumento,mt)
		 end if  
	  end if 	  
      actionSDL="SDLDOCUDELE_" 
      if mid(Oper,1,len(actionSDL))=actionSDL then 
	     IdAttivitaDocumento = Cdbl("0" & Mid(Oper,len(actionSDL)+1))
		 if Cdbl(IdAttivitaDocumento)>0 then
            xx=SetAttivitaDocumentoDelete(IdAttivitaDocumento)
		 end if  
	  end if
      actionSDL="SDLDOCUREQS_" 
      if mid(Oper,1,len(actionSDL))=actionSDL then 
	     IdAttivitaDocumento = Cdbl("0" & Mid(Oper,len(actionSDL)+1))
		 if Cdbl(IdAttivitaDocumento)>0 then
            xx=SetAttivitaDocumentoRichiesto(IdAttivitaDocumento)
		 end if  
	  end if  	  
      actionSDL="SDLDOCUREQN_" 
      if mid(Oper,1,len(actionSDL))=actionSDL then 
	     IdAttivitaDocumento = Cdbl("0" & Mid(Oper,len(actionSDL)+1))
		 if Cdbl(IdAttivitaDocumento)>0 then
            xx=SetAttivitaDocumentoNonRichiesto(IdAttivitaDocumento)
		 end if  
	  end if	  
   end if 

%>