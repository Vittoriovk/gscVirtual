<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/common/functionDataList.asp"-->
<%

action  = ucase(trim(request("action"))) 
id      = trim(request("id"))
attr    = trim(Request("attr"))

if action<>"" then 
   if action=ucase("Stato") then 
      retVal = createDataList("STATO",id,"")
   end if 
   if action=ucase("COMUNE_IT") then 
      retVal = createDataList("COMUNE_IT",id,"")
   end if 
   if action=ucase("PROVINCIA_IT") then 
      retVal = createDataList("PROVINCIA_IT",id,"")
   end if    
   if action=ucase("COMUNE_BYSIGLAPROV_IT") then 
      retVal = createDataList("COMUNE_BYSIGLAPROV_IT",id,attr)
   end if 
   if action=ucase("COMUNE_BYPROVINCIA_IT") then 
      retVal = createDataList("COMUNE_BYPROVINCIA_IT",id,attr)
   end if 
   
   
end if 
response.write retVal
response.end 

%>
