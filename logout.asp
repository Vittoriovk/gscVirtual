<%
if Session("LoginExtePage") = Session("LoginHomePage") then 
   retP = ""
else
   retP = Session("LoginExtePage")
end if 
if retP = "" then 
   retP = "/gscVirtual"
end if 
Session.Abandon
response.redirect retP
response.end 

%>
