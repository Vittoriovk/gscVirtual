<%
Ambiente="DEV"
Dim ConnMsde
Set ConnMsde = Server.CreateObject("ADODB.Connection")
ConnMsde.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=1234;Initial Catalog=Servizi;Data Source=DESKTOP-BEQG7RI\SQLEXPRESS;Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
'ConnMsde.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=MSSql193650;Password=64m78e4545;Initial Catalog=MSSql193650;Data Source=62.149.153.38;Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"

function ConnBeginTran()
	on error resume next
	ConnMsde.BeginTrans  
	err.clear

end function 

function ConnCommitTran()
	on error resume next
	ConnMsde.CommitTrans  
	err.clear

end function 

function ConnRollbackTran()
	on error resume next
	ConnMsde.RollbackTrans  
	err.clear

end function 

function writeTrace(descTrace)
on error resume next 
Dim qi
   qi = "Insert into Trace (Description) values ('" & apici(descTrace) & "')"
   'response.write qi 
   connMsde.execute qi
err.clear 
end function 
function writeTraceAttivita(descTrace,IdAttivita,IdNumAttivita)
on error resume next 
Dim qi
   qi = "Insert into Trace (Description,IdAttivita,IdNumAttivita) values ('" & apici(descTrace) & "','" & apici(IdAttivita) & "'," & NumForDb(IdNumAttivita) & ")"
   'response.write qi 
   connMsde.execute qi
err.clear 
end function 
%>