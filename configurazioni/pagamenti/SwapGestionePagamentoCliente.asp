<%
  NomePagina="SwapGestionePagamentoCliente.asp"
  titolo="Swap Gestione Borsellino cliente"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
   if Session("LoginTipoUtente")=ucase("Clie") then
	  Session("swap_PaginaReturn")   = "/link/ClientePagamento.asp"
      Session("swap_IdCliente")      = Session("LoginIdCliente") 
      response.redirect RitornaA("configurazioni/pagamenti/ListaPagamentoAccount.asp")
   else 
      response.redirect Virtualhost
   end if 
   response.end 
%>