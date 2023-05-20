<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
flagDebug=true
Oper = ""
IdCompagnia   = cdbl("0" & Request("IdCompagnia"))
IdFornitore   = cdbl("0" & Request("IdFornitore"))

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
   flagDebug = false 
end if 

costoDiritti           = 0
ImptIntermediazione    = 0
ImptIntermediazioneMin = 0
ImptIntermediazioneMax = 0
if Cdbl(IdCompagnia)>0 and Cdbl(IdFornitore)>0 then
   IdAccForn  = Cdbl(LeggiCampo("select * from Fornitore Where IdFornitore=" & IdFornitore,"IdAccount")) 
   qProd      = "select * from Prodotto where IdCompagnia=" & IdCompagnia & " and IdAnagServizio='CAUZ_PROV'"
   idProdotto = Cdbl(LeggiCampo(qProd,"IdProdotto"))
   qFascia    = qFascia
   qFascia    = qFascia & " select * from DirittiEmissione"
   qFascia    = qFascia & " Where IdAccountFornitore in (" & idAccForn & ",0)"
   qFascia    = qFascia & " and IdProdotto in ("  & idProdotto & ",0)"
   qFascia    = qFascia & " and IdCompagnia in (" & IdCompagnia  & ",0)"
   qFascia    = qFascia & " order by IdAccountFornitore Desc,IdCompagnia Desc, IdProdotto Desc"

   if flagDebug=true then 
      response.write qFascia & "<br>"
   end if    

   Set Rs = Server.CreateObject("ADODB.Recordset")
   'response.write MyContQ
   Rs.CursorLocation = 3
   Rs.Open qFascia, ConnMsde 
   if Rs.eof=false then 
      ImptIntermediazione    = Rs("ImportoIntermediazione")
      ImptIntermediazioneMin = Rs("ImportoIntermediazioneMin")
      ImptIntermediazioneMax = Rs("ImportoIntermediazioneMax")
  
      costoDiritti = Rs("Importo")
   end if 
   rs.close 
end if 

response.write costoDiritti & "|" & ImptIntermediazione  & "|" & ImptIntermediazioneMin  & "|" & ImptIntermediazioneMax

%>