<%
LCID = 1040 
Session.LCID = LCID 

if Session.SessionID <> Session("SessionId") then 
   Response.Redirect VirtualPath & "/SessioneScaduta.asp"
End If


%>