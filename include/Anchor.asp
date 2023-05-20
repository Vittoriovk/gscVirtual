<%
on error resume next 
'RiferimentoA="divCol;Href;classeA;size;oper;tooltip;text;onclick;showDiv"
ArD=split(RiferimentoA,";")
dcol=ArD(0)
dref=ArD(1)
clas=ArD(2)
size=ArD(3)
if size="1" then 
   size="fa-lg "
end if 
if size="2" then 
   size="fa-2x "
end if 
opbt=ArD(4)
totp=ArD(5)
text=ArD(6)
oncl=ArD(7)
if ubound(ArD)>7 then 
   sdiv=ArD(8)
else 
   sdiv=""
end if 
err.clear 
imag="fa-plus-square"
if opbt="plus" then 
   imag="fa-plus-square"
end if 
if opbt="prev" then 
   imag="fa-arrow-left"
end if 
if opbt="upda" then 
   imag="fa-refresh"
end if 
if opbt="dele" then 
   imag="fa-trash"
end if
if opbt="save" then 
   imag="fa-save"
end if
if opbt="dett" then 
   imag="fa-list-ol"
end if
if opbt="dett-t" then 
   imag="fa-text-width"
end if
if opbt="dett-g" then 
   imag="fa-glide-g"
end if
if opbt="tecn" then 
   imag="fa-cogs"
end if
if opbt="prof" then 
   imag="fa-user"
end if
if opbt="clie" then 
   imag="fa-users"
end if
if opbt="pdf" then 
   imag="fa-file-pdf-o"
end if
if opbt="docu" then 
   imag="fa-file-pdf-o"
end if
if opbt="prod" then 
   imag="fa-product-hunt"
end if
if opbt="forn" then 
   imag="fa-user-circle-o"
end if
if opbt="comp" then 
   imag="fa-vcard-o"
end if
if opbt="perc" then 
   imag="fa-percent"
end if
if opbt="from" then 
   imag="fa-caret-square-o-left"
end if
if opbt="money" then 
   imag="fa-money"
end if
if opbt="pict" then 
   imag="fa-picture-o"
end if
if opbt="hand" then 
   imag="fa-handshake-o"
end if
if opbt="info" then 
   imag="fa-info"
end if
if opbt="mail" then 
   imag="fa-envelope-o"
end if
if opbt="lucc" then 
   imag="fa-unlock-alt"
end if
if opbt="puli" then 
   imag="fa-eraser"
end if
if opbt="penn" then 
   imag="fa-pencil-square-o"
end if
if opbt="manu" then 
   imag="fa-wrench"
end if
if opbt="ok" then 
   imag="fa-thumbs-o-up"
end if
if opbt="ko" then 
   imag="fa-thumbs-o-down"
end if
if opbt="uplo" then 
   imag="fa-upload "
end if
if opbt="matr" then 
   imag="fa-table "
end if
if opbt="cert" then 
   imag="fa-certificate "
end if
if opbt="sele" then 
   imag="fa-hand-pointer-o "
end if
if opbt="crea" then 
   imag="fa-magic"
end if
if opbt="minu" then 
   imag="fa-minus-square"
end if 
if opbt="card" then 
   imag="fa-id-card"
end if 
if opbt="requ" then 
   imag="fa-check"
end if 
if opbt="remo" then 
   imag="fa-times"
end if 
if opbt="erre" then 
   imag="fa-registered"
end if 
if opbt="logi" then 
   imag="fa-sign-in"
end if 
if opbt="copy" then 
   imag="fa-copyright"
end if 
if opbt="effe" then 
   imag="fa-foursquare"
end if 


if dcol<>"" then 
   dcol="class=""" & dcol & """"
end if 
if clas<>"" then 
   clas="class=""" & clas & """"
end if 

if dRef<>"" then 
   dRef= "href='" & dRef & "'"
end if 

%> 
<%if sdiv<>"N" then%> 
<div <%=dcol%>>
<%end if %>

   <a <%=dref%> <%=clas%>
	<%
	if totp<>"" then 
		'response.write " data-toggle='tooltip' data-placement='bottom' title='" & totp & "'"
		response.write " title=""" & totp & """"
	end if 
	if oncl<>"" then 
		response.write " onclick=""" & oncl & ";"""
	end if 
	response.write ">"
	%>
	<i class="fa <%=size%> <%=imag%>"></i><%=text%></a>  

<%if sdiv<>"N" then%> 
</div>
<%end if %>
	


