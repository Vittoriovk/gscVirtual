<tbody>

<%
'caricamento tabella 
err.clear
Set Rs = Server.CreateObject("ADODB.Recordset")

MySql = "" 
MySql = MySql & " Select* From ElencoValore "
MySql = MySql & " Where IdElenco = " & IdElenco
MySql = MySql & " order By Sequenza"

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

DescElenco = ""
If Err.number<>0 then	
	response.write "<tr><td align='center'><b>" & Err.description & "</b></td></tr>"
else
    Do While Not rs.EOF
		if DescElenco<>"" then 
		   DescElenco=DescElenco & ","
		end if 
		DescElenco=DescElenco & Rs("ValoreElenco")
		rs.MoveNext
	Loop
	rs.close
End if 


DescLoaded="0"
NumCols = 2
%>
	<!--#include virtual="/gscVirtual/include/showError.asp"-->
<tr>
	<td>
        <input value="<%=DescElenco%>" type="text" name="DescElenco0" id="DescElenco0" class="form-control"  >
   </td>
</tr> 
<%if flagModElenco=1 then %>
<tr>
	<td align="center">
		<%
		funToCall="SaveWithOper('INS')"
		RiferimentoA="center;#;;2;save;Registra; Registra;" & funToCall & ";S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
	</td>
</tr>
<%end if %>

</tbody>