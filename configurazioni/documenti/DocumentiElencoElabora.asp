<%
'caricamento tabella 
'input IdTabellaUpload,IdTabellaKeyInt,IdTabellaKeyString
'flagOperSede = U,D 
Set Rs = Server.CreateObject("ADODB.Recordset")
if Condizione<>"" then 
   Condizione = " And " & Condizione
end if 
MySql = "" 
MySql = MySql & " Select a.*,isnull(b.DescDocumento,'') as DescDocumento from Upload A "
MySql = MySql & " left join Documento B on a.IdTipoDocumento = b.IdDocumento "
MySql = MySql & " Where IdTabella = '" & Apici(IdTabella) & "'" 
MySql = MySql & " and   IdTabellaKeyInt = " & IdTabellaKeyInt
MySql = MySql & " and   IdTabellaKeyString = '" & Apici(IdTabellaKeyString) & "'" 
MySql = MySql & Condizione
MySql = MySql & " order By a.ValidoDal desc,a.NomeDocumento"

'response.write MySql 

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

DescLoaded=""
NumCols = 4
if ShowValidDate then
   NumCols = NumCols + 2
end if 
NumRec  = 0  
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
%>
	<!--#include virtual="/gscVirtual/include/CheckRsNoData.asp"-->
<div class="table-responsive"><table class="table"><tbody>
<thead>
	<tr>
		<th scope="col">Documento
		<%if instr(OperAmmesse,"U")>0 then %>
<a href="#"  title="Inserisci" onclick="AttivaFunzione('CALL_INS','0');">
	<i class="fa fa-2x fa-plus-square"></i></a>  
        <%end if%>
		
		</th>
		<th scope="col">Tipo</th>
		<th scope="col">Descrizione</th>
		<%if ShowValidDate then%>
		   <th scope="col">Valido dal</th>
		   <th scope="col">Valido al</th>
		<%end if %>
		<th scope="col">Azioni</th>
	</tr>
</thead>

<%
if MsgNoData="" then 
	Do While Not rs.EOF 
		Primo=Primo+1
		NumRec=NumRec+1
		Id=Rs("IdUpload")
		ValidoDal=Stod(Rs("ValidoDal"))
		ValidoAl=Stod(Rs("ValidoAl"))
		%> 
	<tr scope="col"> 
		<td><%
		      response.write Rs("NomeDocumento") 
			  docLink = replace(VirtualPath & DirectoryUpload & "/" & Rs("PathDocumento"),"//","/")
			  %>
			  <a href='<%=docLink%>' data-toggle="tooltip" data-placement="top" title="Scarica" download>
			  <i class="fa fa-2x fa-file-pdf-o"></i>
			  </a>  
		</td>
		<td><%=Rs("DescDocumento")%></td>
		<td><%=Rs("DescEstesa")%></td>
		<%if ShowValidDate then%>
		   <td><%=ValidoDal%></td>
		   <td><%=ValidoAl%></td>
		<%end if %>
		
		<td>
			<%if instr(OperAmmesse,"U")>0 then
			  RiferimentoA="col-2;#;;2;upda;Modifica;;AttivaFunzione('CALL_UPD'," & Id & ",true,'','','');N"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
			<%end if %>
			<%if instr(OperAmmesse,"D")>0 then 
			  RiferimentoA="col-2;#;;2;dele;Cancella;;AttivaFunzione('DEL'," & Id & ",true,'','','');N"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
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