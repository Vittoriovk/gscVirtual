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
			<th scope="col">Tipo Sede
			  <%if instr(OperAmmesse,"C")>0 then %>
			  <a class="button-color-click" onclick="attivaFormSede('<%=NomeStruttura%>',0)"   role="button" >
			  <i class="fa fa-2x fa-plus-square"></i></a>
			  <%end if %>
			</th>	
			<th scope="col">Stato</th>
			<th scope="col">Indirizzo</th>
			<th scope="col">Cap</th>
			<th scope="col">Comune</th>
			<th scope="col">Provincia</th>			
			<th scope="col">Azioni</th>
		</tr>
	</thead>
<%

on error resume next 

MySql = "" 
MySql = MySql & " Select a.*,b.*,Provincia as ProvinciaIT ,C.DescStato from AccountSede A, TipoSede B, Stato C "
MySql = MySql & " Where a.IdAccount = " & IdAccount
MySql = MySql & " and   a.IdTipoSede   = b.IdTipoSede"
MySql = MySql & " and   A.IdStato = c.IdStato"
MySql = MySql & " order By Ordine"
'response.write MySql 

Set rs = Server.CreateObject("ADODB.Recordset")
Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

DescLoaded=""
NumCols = 6
NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
Primo=0

%>
	<!--#include virtual="/gscVirtual/include/CheckRsNoData.asp"-->
<%
if MsgNoData="" then 
    Primo = 0
	Do While Not rs.EOF 
		Primo=Primo+1
		valoStruttura=""
		For i=1 to idxMaxArr
		    'response.write "i=" & i & " " &  campiStrutturaArrNome(i) & "<br>"
		    campoDb=campiStrutturaArrNome(i)
			if ucase(campoDb)="AZIONI" then 
			   valoStruttura = valoStruttura & "OLD~~"
			else 
			   valoStruttura = valoStruttura & rs(campoDb) & "~~"
			end if 
		Next 
		'response.write primo & "<br>" & valoStruttura & "<br>"
		
		%> 
		<input type="hidden" name="<%=NomeStruttura%>_Row_<%=Primo%>" id="<%=NomeStruttura%>_Row_<%=Primo%>" value="<%=valoStruttura%>">
		
		<tr>
		    <td scope="col"><%=Rs("DescTipoSede")%></td>
			<td scope="col"><%=Rs("Indirizzo")%></td>
			<td scope="col"><%=Rs("DescStato")%></td>
			<td scope="col"><%=Rs("Cap")%></td>
			<td scope="col"><%=Rs("Comune")%></td>
			<td scope="col"><%=Rs("Provincia")%></td>

		    <td scope="col">
			  <div class="button-color-click">
			  <a onclick="attivaFormSede('<%=NomeStruttura%>',<%=primo%>)" role="button" <i class="fa fa-2x fa-refresh"></i></a>
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
