<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%

Session("swap_CurrentFun")     = request("CurrentFun")
Session("swap_OperAmmesse")    = request("OperAmmesse")
Session("swap_PaginaReturn")   = request("PaginaReturn")
Session("swap_PageToCall")     = request("PageToCall")
Session("swap_opzioneSidebar") = request("opzioneSidebar")
   
response.redirect RitornaA("utility/CercaCliente.asp")
response.end 

%>