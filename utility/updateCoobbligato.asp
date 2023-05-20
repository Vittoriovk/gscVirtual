<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
flagDebug=true
IdCauzione = cdbl("0" & Request("IdCauzione"))
IdDett     = cdbl("0" & Request("IdCauzioneCoobbligato"))
action     = ucase(trim(Request("action")))
RagSoc     = apici(Request("RagSoc"))
CF         = apici(Request("CF"))
PI         = apici(Request("PI"))
Indirizzo  = apici(Request("Indirizzo"))
Comune     = apici(Request("Comune"))
Cap        = apici(Request("Cap"))
Provincia  = apici(Request("Provincia"))

if Cdbl(IdCauzione)>0 then 
   MyQ = ""
   if Cdbl(IdDett)=0 then 
      MyQ = MyQ & " insert into CauzioneCoobbligato"
	  MyQ = MyQ & "(IdCauzione,PI,CF,RagSoc,Indirizzo,Cap,Comune,Provincia) values "
      MyQ = MyQ & "(" & IdCauzione & ",'" & PI & "','" & CF & "','" & RagSoc & "','" & Indirizzo & "'"
	  MyQ = MyQ & ",'" & Cap & "','" & Comune & "','" & Provincia & "')"  
   elseif action="DELETE" then 
      MyQ = MyQ & " delete from CauzioneCoobbligato"
	  MyQ = MyQ & " Where IdCauzioneCoobbligato = " & IdDett 
	  MyQ = MyQ & " and   IdCauzione=" & IdCauzione   
   else 
      MyQ = MyQ & " update CauzioneCoobbligato set "
	  MyQ = MyQ & " PI ='" & PI & "'"  
	  MyQ = MyQ & ",CF ='" & CF & "'"   
	  MyQ = MyQ & ",RagSoc = '" & Ragsoc & "'"   
      MyQ = MyQ & ",Indirizzo = '" & Indirizzo & "'"   
      MyQ = MyQ & ",Cap = '" & Cap & "'"   
      MyQ = MyQ & ",Comune = '" & Comune & "'"   
      MyQ = MyQ & ",Provincia = '" & Provincia & "'"   
	  MyQ = MyQ & " Where IdCauzioneCoobbligato = " & IdDett 
	  MyQ = MyQ & " and   IdCauzione=" & IdCauzione
   end if 
   if flagDebug=true then 
      response.write MyQ
   end if 
   ConnMsde.execute MyQ 
end if 


%>