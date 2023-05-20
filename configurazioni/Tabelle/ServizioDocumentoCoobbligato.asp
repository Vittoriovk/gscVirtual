<%
session("Swap_PaginaReturn") = "SupervisorConfigurazioni.asp"
Session("swap_IdAnagServizio") = "CAUZ_PROV"
Session("swap_IdTipoUtenza")   = "COOB"
response.redirect "ServizioDocumento.asp"
response.end 
%>