<%
  NomePagina="SwapGestionePagamentoCliente.asp"
  titolo="Swap Gestione Borsellino cliente"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
   if Session("LoginTipoUtente")=ucase("Clie") then
	  Session("swap_PaginaReturn")   = "/link/ClientePagamento.asp"
      Session("swap_IdCliente")      = Session("LoginIdCliente") 
	  Session("swap_IdTipoCredito")  = "BORS" 
      response.redirect RitornaA("configurazioni/pagamenti/ListaPagamentoAccountStorico.asp")
   else 
      response.redirect Virtualhost
   end if 
   response.end 
%>