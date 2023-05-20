<!--#include virtual="/gscVirtual/include/utility.asp"-->
<%
'salvo eventuali parametri 
TipoRicercaExt   = Session("swap_TipoRicercaExt")
testo_ricercaExt = Session("swap_testo_ricercaExt")
xx=RemoveSwap()
'ripristino 
Session("swap_TipoRicercaExt")   = TipoRicercaExt
Session("swap_testo_ricercaExt") = testo_ricercaExt

Session("swap_PaginaReturn") = "link/BackOAffidamento.asp"
Session("swap_TipoRife")     = "COOB"
response.redirect "ValidazioneBackO.asp"
response.end 
%>