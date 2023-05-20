<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
IdCauzione = cdbl("0" & Request("IdCauzione"))
IdDett     = cdbl("0" & Request("IdCauzioneATI"))
action     = ucase(trim(Request("action")))

if Cdbl(IdCauzione)>0 then 
   MyQ = ""
   if Action="INSERT" then 
      MyQ = MyQ & " insert into CauzioneATI"
	  MyQ = MyQ & "(IdCauzione,PI,RagSoc,IdAccountAti) "
	  MyQ = MyQ & " select " & IdCauzione & " as IdCauzione,PI,RagSoc,IdAccountAti"
	  MyQ = MyQ & " from AccountAti Where IdAccountAti=" & IdDett
   else 
      MyQ = MyQ & " delete from CauzioneATI"
	  MyQ = MyQ & " Where IdAccountAti = " & IdDett 
	  MyQ = MyQ & " and   IdCauzione=" & IdCauzione   
   end if 
   ConnMsde.execute MyQ 
end if 


%>