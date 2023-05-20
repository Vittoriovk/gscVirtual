<%if ao_row<>"" then %>
<div class="row <%=ao_row%>">
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
end if%> 

<%if SoloLettura then%>
    <input type="text" class="form-control" readonly id="<%=ao_nid%>" name="<%=ao_nid%>" value="<%=ao_val%>">

<%else%>
    <input type="text" class="form-control" id="<%=ao_nid%>" name="<%=ao_nid%>" value="<%=ao_val%>">
<%
end if 

if ao_div<>"" then 
%>
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
