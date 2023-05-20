
<%
If Err.number<>0 then	
	MsgNoData = Err.description
elseIf Rs.EOF then	
	MsgNoData = ""
End if
if MsgNoData<>"" or MsgErrore<>"" then %>
	<div class="table-responsive"><table class="table"><tbody>
	<tr>
		<td colspan='<%=NumCols%>' align='center'>
		<div class="bg-danger text-white"><%=server.htmlencode(MsgErrore) & " " & server.htmlencode(MsgNoData) %></div>
		</td>
	</tr>	
	</tbody></table></div>
<% end if %>



