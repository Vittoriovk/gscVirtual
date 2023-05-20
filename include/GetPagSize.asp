<%
PageSize=0

if UsaPaginazione=true then 
	if Request("CallingPage")=NomePagina then 
	   PageSize=cdbl("0" & Request("RowPagina"))
	else
       PageSize=0
	end if 

	if Cdbl(PageSize)=0 then 
	   PageSize=10
	end if 
    err.clear
	   
end if 

Cpag = cdbl("0" & Request.form("Pagina"))
if Cpag = 0 then
	Cpag = cdbl("0" & Request.querystring("Pagina"))
end if
If CPag=0 then 
   CPag=1 
End If 

Oper = Request("Oper")
if Oper = "" then
   Oper = Request.querystring("Oper")
end if
Oper = ucase(Oper)
'SERVE  A GESTIRE UN EVENTUALE REFRESH DELLA PAGINA 
TimeStamp = Dtos() & TimeTos()
TimePage = Request("TimePage")

If (Oper="INS" or OPER="UPD" or OPER=ucase("RemoveItem")) and Session("TimeStamp")<>"" then  
	If Session("TimeStamp") = TimePage Then
		Oper=" "
	End If
end if 
FirstLoad=(Request("CallingPage")<>NomePagina)
%>