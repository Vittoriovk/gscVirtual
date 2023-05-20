<!--#include virtual="/gscVirtual/include/utility.asp"-->
<%
xx=RemoveSwap()
DestroyCurrent()
Session("swap_PaginaReturn") = "link/BackOAffidamento.asp"
Session("swap_TipoRichiesta")="STORICO"
Session("swap_TipoRife")     = "ATI"
response.redirect "ValidazioneBackO.asp"
%>