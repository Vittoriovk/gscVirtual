<%
qSel = ""
qSel = qsel & " select top 1 * "
qSel = qsel & " from Upload "
qSel = qsel & " Where IdTabella = '" & IdTipoCauzDoc & "'"
qSel = qsel & " and IdTabellaKeyInt = " & IdCodeCauzdoc

'response.write qSel
trDocCauz="0" & LeggiCampo(qSel,"IdTabellaKeyInt")
if cdbl(trDocCauz)>0 then 
   RiferimentoA=";#;;2;pdf;Documenti;;popolaDocCauzione('" & IdTipoCauzDoc & "','" & IdCodeCauzdoc & "');N" 
%>
   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
<%
end if  
%>