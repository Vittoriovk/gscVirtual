
<%
   Set Rs = Server.CreateObject("ADODB.Recordset")
   
   MySql = "" 
   MySql = MySql & " select top 10 * from EventoAccount A"
   MySql = MySql & " inner join Evento B on A.idEvento = B.IdEvento " 
   MySql = MySql & " left join  Processo C on b.IdProcesso = c.IdProcesso"
   MySql = MySql & " left join  TipoEvento D on b.IdTipoEvento = D.IdTipoEvento"
   MySql = MySql & " Where IdAccount = " & Session("LoginIdAccount") 
   MySql = MySql & " order by A.IdEvento Desc "   

   Rs.CursorLocation = 3 
   Rs.Open MySql, ConnMsde

   MsgNoData  = ""
   
  
%>

<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<!--#include virtual="/gscVirtual/include/CheckRsNoData.asp"-->
	  <%if MsgNoData="" then 
	       if rs.EOF=false then 
	  %>
		<div class="table-responsive"><table class="table"><tbody>
		<thead>
		<tr>
			<th scope="col" width="12%">Data Evento</th>
			<th scope="col">Tipo Evento</th>
			<th scope="col">Cliente</th>
		    <th scope="col">Descrizione</th>
			<th scope="col">Azioni</th>
		</tr>
		</thead>

		<%  end if 
		
			Do While Not rs.EOF
			   idEvento = Rs("IdEvento")
               dtEv = StoD(RS("DataEvento")) & " " & mid(StoTime(Rs("TimeEvento")),1,5)
			   DescTipoEvento = trim(Rs("DescTipoEvento"))
			   if DescTipoEvento<>"" then 
			      DescTipoEvento = " : " & DescTipoEvento
			   end if 
			%>
			<tr scope="col">
				<td>
					<input class="form-control" type="text" readonly value="<%=dtEv%>">
				</td>
                 <td>
                   <input class="form-control" type="text" readonly value="<%=Rs("DescProcesso") & DescTipoEvento %>">
                 </td>
                 <td>
				   <%
				   descCliente  = "N.D."
				   notecliente  = ""
				   linkRef      = ""
				   IdAffRichi   = 0
				   if ucase(rs("IdProcesso"))="AFFI" and Ucase(RS("IdTabella"))=ucase("AffidamentoRichiestaComp") then 
				      'response.write rs("key") & " " & RS("IdTabella")
					  qKey = ""
				      qKey = qKey & " select * from "
					  qKey = qKey & " AffidamentoRichiestaComp A,AffidamentoRichiesta B,Cliente C "
					  qKey = qKey & " Where " & Rs("IdKey")
					  qKey = qKey & " and A.IdAffidamentoRichiesta=B.IdAffidamentoRichiesta"
					  qKey = qKey & " and B.IdAccountCliente = C.IdAccount"
					  'response.write qKey 
					  descCliente = LeggiCampo(qKey,"Denominazione")
					  notecliente = LeggiCampo(qKey,"NoteAffidamentoCliente")
					  if notecliente<>"" then 
					     notecliente = " - " & notecliente
					  end if 
					  linkRef = "/gscVirtual/utility/swapEvento.asp?IdEvento=" & idEvento
				   end if 
				   if ucase(rs("IdProcesso"))="COOB" and Ucase(RS("IdTabella"))=ucase("AccountCoobbligato") then 
				      'response.write rs("key") & " " & RS("IdTabella")
					  qKey = ""
				      qKey = qKey & " select * from "
					  qKey = qKey & " AccountCoobbligato A,Cliente C "
					  qKey = qKey & " Where " & Rs("IdKey")
					  qKey = qKey & " and A.IdAccount = C.IdAccount"
					  'response.write qKey 
					  descCliente = LeggiCampo(qKey,"Denominazione")
					  linkRef = "/gscVirtual/utility/swapEvento.asp?IdEvento=" & idEvento
				   end if 
				   
				   
				   %>
                   <input class="form-control" type="text" readonly value="<%=DescCliente%>">
                 </td>				
                 <td>
                   <input class="form-control" type="text" readonly value="<%=Rs("DescEvento") & notecliente %>">
                 </td>
				 <td>
				   <%if linkRef<>"" then %>
				   <a href="<%=linkRef%>"><i class="fa fa-2x fa-sign-in">&nbsp;</i></a>
				   <%end if %>
				 </td>

			</tr>
		<%	
		rs.MoveNext
	Loop
end if 
rs.close

%>
</tbody></table></div> <!-- table responsive fluid -->

