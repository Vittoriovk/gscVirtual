<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%

cf      = trim(request("cf"))
id      = Cdbl("0" & trim(request("idAccountCliente")))
retVal  = ""

Set Rs = Server.CreateObject("ADODB.Recordset")
if cf<>"" and cdbl(id)>0 then 
   Q = ""
   q = q & " select top 1 * from Formazione "
   q = q & " Where IdAccountCliente=" & Id
   q = q & " and codiceFiscale<>'' "
   q = q & " and codiceFiscale='" & apici(cf) & "'"
   Rs.CursorLocation = 3 
   'response.write q
   Rs.Open q, ConnMsde   
   if Rs.eof = false then 
      Cognome  = Rs("Cognome")
	  Nome     = Rs("Nome")
	  userPiattaforma = rs("userPiattaforma")
	  passPiattaforma = rs("passPiattaforma")
      retVal= "cogn=" & cognome & "|nome=" & nome & "|user=" & userPiattaforma & "|pass=" & passPiattaforma      
   end if 
   Rs.close  
end if 

response.write retVal
response.end 

%>
