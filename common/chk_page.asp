<%

if default_check_profile="" then 
   'recupero i profili abilitati per la pagina : * per tutti
   qq="select top 1 * from PaginaTipoAccount where NomePagina='" & NomePagina & "'"
   xx=LeggiCampo(qq,"ListaTipoAccount")

   'response.write qq & "::" & xx & err.description
   'response.end 
   'se il profilo non è contenuto e non vale per tutti
   if instr(ucase(xx),ucase(Session("LoginTipoUtente")))=0 and xx<>"*" and Ambiente="DEV" then 
      Response.write NomePagina
      response.end 
   End If
   if instr(ucase(xx),ucase(Session("LoginTipoUtente")))=0 and xx<>"*" then 
      Response.Redirect VirtualPath & "/SessioneScaduta.asp"
   End If
else 
   'controllo se il profilo è contenuto nel default_check_profile
   if instr(ucase(default_check_profile),ucase(Session("LoginTipoUtente")))=0 and default_check_profile<>"*" then 
      Response.Redirect VirtualPath & "/SessioneScaduta.asp"
   End If
end if 


%>