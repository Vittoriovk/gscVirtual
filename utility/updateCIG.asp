<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
flagDebug=true
IdCauzione = cdbl("0" & Request("IdCauzione"))
IdDett     = cdbl("0" & Request("IdCauzioneCIG"))
action     = ucase(trim(Request("action")))
CIG        = apici(Request("CIG"))
DescCIG    = apici(Request("DescCIG"))

if Cdbl(IdCauzione)>0 then 
   MyQ = ""
   if Cdbl(IdDett)=0 then 
      MyQ = MyQ & " insert into CauzioneCIG"
	  MyQ = MyQ & "(IdCauzione,DescCIG,CIG) values "
      MyQ = MyQ & "(" & IdCauzione & ",'" & DescCIG & "','" & CIG & "')"  
   elseif action="DELETE" then 
      MyQ = MyQ & " delete from CauzioneCIG"
	  MyQ = MyQ & " Where IdCauzioneCIG = " & IdDett 
	  MyQ = MyQ & " and   IdCauzione=" & IdCauzione   
   else 
      MyQ = MyQ & " update CauzioneCIG set "
	  MyQ = MyQ & " DescCIG ='" & DescCIG & "'"  
	  MyQ = MyQ & ",CIG = '"    & CIG & "'"   
	  MyQ = MyQ & " Where IdCauzioneCIG = " & IdDett 
	  MyQ = MyQ & " and   IdCauzione=" & IdCauzione
   end if 
   if flagDebug=true then 
      response.write MyQ
   end if 
   ConnMsde.execute MyQ 
end if 


%>