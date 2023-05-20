<%
   if len(idDaPulire)>0 then 
      RiferimentoA=";#!;;1;puli;Pulisci;;pulisciCampo('" & idDaPulire & "');N"
%>
      <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
<% end if 

   if len(idDaSalvare)>0 then 
      RiferimentoA=";#!;;1;save;Registra;;salvaCampo('" & idDaSalvare & "');N"
%>
      <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
<% end if 
%>

