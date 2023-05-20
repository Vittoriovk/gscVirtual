<%
  NomePagina="SwapDispEconCliente.asp"
  titolo="Swap Disponibilita' economica cliente"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
   if Session("LoginTipoUtente")=ucase("Clie") then
	  Session("swap_PaginaReturn")   = "/link/ClientePagamento.asp"
      Session("swap_IdAccount")      = Session("LoginIdAccount") 
      response.redirect RitornaA("configurazioni/pagamenti/DisponibilitaEconomica.asp")
   else 
      response.redirect Virtualhost
   end if 
   response.end 
%>