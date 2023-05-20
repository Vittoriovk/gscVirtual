<%
session("Swap_PaginaReturn") = "SupervisorConfigurazioni.asp"
Session("swap_IdAnagServizio") = "CAUZ_PROV"
Session("swap_IdTipoUtenza")   = "ATI"
response.redirect "ServizioDocumento.asp"
response.end 
%>