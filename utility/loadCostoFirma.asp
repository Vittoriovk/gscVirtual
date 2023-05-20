<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
IdCompagnia   = cdbl("0" & Request("IdCompagnia"))
IdFornitore   = cdbl("0" & Request("IdFornitore"))
IdTipoFirma   = Request("IdTipoFirma")
IdRefCampo    = Request("IdRefCampo")
  
%>
<!--#include virtual="/gscVirtual/utility/loadCostoFirmaElabora.asp"-->
