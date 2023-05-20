
<%


if MessageFromCaller<>"" then %>
	<div class="table-responsive"><table class="table"><tbody>
	<tr>
		<td colspan='<%=NumCols%>' align='center'>
		<div class="bg-success text-white"><%=server.htmlencode(MessageFromCaller) %></div>
		</td>
	</tr>	
	</tbody></table></div>
<% end if 

If Err.number<>0 then	
	MsgNoData = Err.description
elseIf Rs.EOF then	
	MsgNoData = "Nessun dettaglio in archivio"
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



