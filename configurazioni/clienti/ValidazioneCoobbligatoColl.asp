<!--#include virtual="/gscVirtual/include/utility.asp"-->
<%
xx=RemoveSwap()
Session("swap_PaginaReturn") = "link/CollAffidamento.asp"
Session("swap_TipoRife")     = "COOB"
response.redirect "ValidazioneBackO.asp"
response.end 
%>