<% if MsgErrore<>"" then  %>
<tr scope="col">
	<td colspan='<%=NumCols%>' align='center'>
	<div class="bg-danger text-white"><%=server.htmlencode(MsgErrore)%></div>
	</td>
</tr>
<% end if %>


