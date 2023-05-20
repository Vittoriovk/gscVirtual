<script language="JavaScript">
function CambiaPaginazione()
{
	var tt;
	tt = ValoreDi("PageSize");
	xx=ImpostaValoreDi("Oper","P");
	xx=ImpostaValoreDi("RowPagina",tt);	
	xx=ImpostaValoreDi("Pagina",1);
	document.Fdati.submit();
}
</script>
<% 
if IsNumeric(Pagesize)=false then 
	PageSize=0
end if 
If cdbl(PageSize)>0  then %>
    <nav class="paginazione">
        <ul class="pagination">
			<%
			pagine_prima_dopo = 3
			pagina_attuale = cPag
			RefH="<a class='page-link' href='javascript:Paginazione("
			If CDbl(pagina_attuale - 1) > 0 Then
				Response.Write "<li class='paginate_button page-item previous'>" & refH & "0)'>Inizio</a></li>"
				Response.Write "<li class='paginate_button page-item '>" & refH & (CPag-1) & ")' aria-label='Precedente'><span aria-hidden='true'>&laquo;</span></a></li>"
			End If
			For i = (pagina_attuale - pagine_prima_dopo) To pagina_attuale + pagine_prima_dopo
				If (i <= (pageTotali )) AND (i > 0) Then
					If CDbl(i) = CDbl(pagina_attuale) Then
					%>
						<li class="paginate_button page-item active"><a class='page-link'><%=I%></a></li>
					<%
					Else
					%>
						<li class="paginate_button page-item"><% response.write refH & i & ")'>" & I %></a></li>
					<%
					End If
				End If
			Next
			If CDbl(pagina_attuale) < CDbl(pageTotali) Then
				Response.Write "<li class='paginate_button page-item next'>" & refH & (CPag+1) & ")' aria-label='Successivo'><span aria-hidden='true'>&raquo;</span></a></li>"
				Response.Write "<li class='paginate_button page-item'>" & refH & (pageTotali) & ")'>Fine</a></li>"
			End If
			%>
		</ul>
		<%if false then %>
        <nav class="navbar-right">
            <div class="form-group form-inline">
				<label for="PageSize">Righe per pagina:</Label>
				<%
				ElencoOption="5;5;10;10;25;25;50;50"
				response.write  OptionListaValori("PageSize",ElencoOption,PageSize)
				%>
            </div>
        </nav>
		<%end if %>
	</nav>
<% Else %>
   <input type="Hidden" Name="PageSize" Id="PageSize" value="<%=PageSize%>">
<% End If %>