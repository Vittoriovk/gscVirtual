<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
OperAmmesse   =Request("OperAmmesse")
NomeStruttura =Request("NomeStruttura")
campiStruttura=Request("campiStruttura")
IdAccount     =cdbl("0" & Request("IdAccount"))

dim campiStrutturaArrNome(100)
dim campiStrutturaArrValo(100)
defaultStruttura=""
idxMaxArr=0
KeyCode=split(campiStruttura,";")
i=0
For J=lbound(Keycode) to Ubound(KeyCode)
  if trim(Keycode(j))<>"" then 
     i=i+1
     dettCampi=split(Keycode(j),",")
     campiStrutturaArrNome(i)=dettCampi(0)
     campiStrutturaArrValo(i)=dettCampi(2)
	 defaultStruttura=defaultStruttura+dettCampi(2)+"~~"
  end if 
Next 
idxMaxArr=i
%>
<input type="hidden" name="<%=NomeStruttura%>_Row_0"     id="<%=NomeStruttura%>_Row_0"    value="<%=defaultStruttura%>">
<table class="table">
<tbody>

	<thead>
		<tr>
			<th scope="col">Tipo Contatto
			  <%if instr(OperAmmesse,"C")>0 then %>
			  <a class="button-color-click"  onclick="attivaFormContatto('<%=NomeStruttura%>',0)"   role="button" >
			  <i class="fa fa-2x fa-plus-square"></i></a>
			  <%end if %>
			</th>	
			<th scope="col">Descrizione</th>
			<th scope="col">Note</th>
			<th scope="col">Principale</th>
			<th scope="col">Azioni</th>
		</tr>
	</thead>


<%

'caricamento tabella per account 
'input IdAccount
'flagOperContatto = U,D 
Set Rs = Server.CreateObject("ADODB.Recordset")

MySql = "" 
MySql = MySql & " Select *,FlagPrincipale as ValFlagPrincipale from AccountContatto A, TipoContatto B "
MySql = MySql & " Where a.IdAccount = " & IdAccount
MySql = MySql & " and   a.IdTipoContatto   = b.IdTipoContatto"
MySql = MySql & " order By Ordine"

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

DescLoaded=""
MsgNoData  = ""

%>
	<!--#include virtual="/gscVirtual/include/CheckRsNoData.asp"-->
<%
Primo = 0
if MsgNoData="" then 
    
	Do While Not rs.EOF 
		Primo=Primo+1
		valoStruttura=""
		For i=1 to idxMaxArr
		    campoDb=campiStrutturaArrNome(i)
			if ucase(campoDb)="AZIONI" then 
			   valoStruttura = valoStruttura & "OLD~~"
			else 
			   valoStruttura = valoStruttura & rs(campoDb) & "~~"
			end if 
		Next 
		%> 
		<input type="hidden" name="<%=NomeStruttura%>_Row_<%=Primo%>" id="<%=NomeStruttura%>_Row_<%=Primo%>" value="<%=valoStruttura%>">
		<tr>
		    <td scope="col"><%=Rs("DescTipoContatto")%></td>
			<td scope="col"><%=Rs("DescContatto")%></td>
			<td scope="col"><%=Rs("NoteContatto")%></td>
			<td scope="col"><%=Rs("ValFlagPrincipale")%></td>

		    <td scope="col">
			  <div class="button-color-click">
			  <%if instr(OperAmmesse,"C")>0 then %>
			  <a onclick="attivaFormContatto('<%=NomeStruttura%>',<%=primo%>)" role="button" <i class="fa fa-2x fa-refresh"></i></a>
			  <%end if %>
			  </div>
			</td>
		</tr>	
	
		<%
		rs.MoveNext
	Loop
end if 
rs.close
%>
</tbody>
</table>

<input type="hidden" name="<%=NomeStruttura%>_MaxRow" id="<%=NomeStruttura%>_MaxRow" value="<%=Primo%>">