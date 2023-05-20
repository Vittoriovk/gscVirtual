<!--#include virtual="/gscVirtual/include/utility.asp"-->
<%
xx=RemoveSwap()
DestroyCurrent()
Session("swap_PaginaReturn") = "link/BackOAffidamento.asp"
Session("swap_TipoRichiesta")="STORICO"
Session("swap_TipoRife")     = "COOB"
response.redirect "ValidazioneBackO.asp"
%>