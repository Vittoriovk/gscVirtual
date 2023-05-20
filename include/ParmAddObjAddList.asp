<%if ao_row<>"" then %>
<div class="row <%=ao_row%> " >
<%end if %>

<%if ao_lbd<>"" then %>
   <div class="<%=ao_lbs%>">
      <p class="font-weight-bold"><%=ao_lbd%></p>
   </div>
<%end if %>

<%
if SoloLettura=true then
   ao_Att = "0" 
   ao_ccc = "Where " & ao_ids & " = " & ao_val & " "
   if instr(ucase(ao_tex),"WHERE")=0 then 
      pos = instr(ucase(ao_tex),"ORDER BY")
	  if pos=0 then 
	     ao_tex = ao_tex & ao_ccc
	  else 
	     ao_tex = mid(ao_tex,1,pos-1) & " " & ao_ccc & mid(ao_tex,pos)
	  end if 
   end if 
   'response.write ao_tex
end if 
%>

<% 
if ao_div<>"" then 
%>
   <div class = "<%=ao_div%>">
<%
end if 
'              ListaDbChangeCompleta(Query ,Name  ,CodValue,ColCod,ColText,FlagVuoto,Change,Campo,Larghezza,DescVuoto,DescNoData,Classe)
response.write ListaDbChangeCompleta(ao_Tex,ao_nid,ao_val  ,ao_Ids,ao_Des ,ao_Att   ,ao_Eve,""   ,""       ,ao_Plh   ,ao_NoD    ,ao_cla)
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
