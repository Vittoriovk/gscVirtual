<%if ao_row<>"" then %>
<div class="row <%=ao_row%>" >
<%end if %>

<%if ao_lbd<>"" then %>
   <div class="<%=ao_lbs%>">
      <p class="font-weight-bold"><%=ao_lbd%></p>
   </div>
<%end if %>



<% 
if ao_div<>"" then 
%>
   <div class = "<%=ao_div%>">
<%
end if
 
parm_AddObject = ao_Obj & "|id=" & ao_Nid & "|name=" & ao_Nid & ao_Val & ao_Att & ao_Plh & ao_Cla & ao_Lev & ao_Eve & ao_Par & ao_Tex & ao_Tit & ao_ico 
'response.write parm_AddObject 
response.write AddObject(parm_AddObject) 

if ao_div<>"" then %>
   </div>
<% end if %>

<%if ao_3ld<>"" then %>
   <div class="<%=ao_3ls%>">
      <p class="font-weight-bold"><%=ao_3ld%></p>
   </div>
<%end if %>

<%if ao_row<>"" then %>
</div> 
<%end if %>
