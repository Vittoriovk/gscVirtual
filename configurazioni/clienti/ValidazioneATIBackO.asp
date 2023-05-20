<!--#include virtual="/gscVirtual/include/utility.asp"-->
<%
xx=RemoveSwap()
Session("swap_PaginaReturn") = "link/BackOAffidamento.asp"
Session("swap_TipoRife")     = "ATI"
response.redirect "ValidazioneBackO.asp"
response.end 
%>