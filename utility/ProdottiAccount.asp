<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<script>
function AssProdAccount(id,sendData)
{
   var op="D";
   if($("#checkProd" + id).prop("checked") == true)
      op="I";
   var prod=$("#CodiceProdotto" + id).val();
   var dataIn="op=" + op + "&sendData=" + sendData + "&codiceProdotto=" + prod;  
   var vp=$("#localVirtualPath").val();
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "utility/updateData.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
        if (msg=="")
			esito = true;
		else {
			esito = false;
		}
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}
</script>
	<table class="table"><tbody>
	<thead>
		<tr>
			<th scope="col">Prodotto</th>
			<th scope="col">Codice</th>
			<th scope="col">Seleziona</th>
		</tr>
	</thead> 

<%
flagDebug=false
Oper = ""
IdAccount = 0

sendDatao = request("sendData")
sendData  = request("sendData")
if flagDebug=true then 
   response.write sendData & "<br>"
end if 
sendData = DecryptWithKey(sendData,Session("CryptKey"))
if flagDebug=true then 
   response.write sendData & "<br>"
end if 
arD=split(sendData,"|")
for J=lbound(arD) to ubound(arD)
   campo=arD(j)
   ptr=instr(campo,"=")
   if flagDebug=true then 
      response.write "campo " & Campo & "<br>"
	  response.write "ptr   " & ptr & "<br>"
   end if 
   
   if ptr>0 then 
      k=trim(mid(campo,1,ptr-1))
	  v=trim(mid(campo,ptr+1))
	  k=ucase(trim(k))
      if flagDebug=true then 
         response.write "k " & k & "<br>"
	     response.write "v " & v & "<br>"
      end if 	  
	  if k=ucase("IdAccount") then 
	     IdAccount = V
	  end if 
   end if 
next 
IdAccount = Cdbl("0" & IdAccount)
if Cdbl(IdAccount)>0 then 
   inCom = "(select IdCompagnia From AccountCompagnia Where IdAccount= " & IdAccount & " ) "
   MySql = "" 
   MySql = MySql & " Select a.*, isnull(b.IdProdotto,0) As idEventoB,isnull(b.CodiceProdotto,'') as CodProF "
   MySql = MySql & " From Prodotto A left join AccountProdotto b  "
   MySql = MySql & " on    A.IdProdotto = B.IdProdotto "
   MySql = MySql & " and   b.IdAccount = " & IdAccount
   MySql = MySql & " Where (B.IdProdotto is not null "
   MySql = MySql & " or A.IdCompagnia in " & inCom & " )"
   MySql = MySql & " order By A.DescProdotto"
   if flagDebug=true then 
      response.write MySql
   end if 
   
   Set RsDet = Server.CreateObject("ADODB.Recordset")
   RsDet.CursorLocation = 3 
   RsDet.Open MySql, ConnMsde 
   
   
   Do While Not RsDet.EOF 
      Id=RsDet("IdProdotto")
	  selezionato=""
	  if Cdbl(RsDet("idEventoB"))>0 then 
	     selezionato=" checked "
	  end if 
	  codiceProdotto=RsDet("CodiceProdotto")
	  if RsDet("CodProF")<>"" then 
	     codiceProdotto = RsDet("CodProF")
	  end if 
	  
	  
	  sendData="Oper=UpdProdAccount|IdProdotto=" & Id & "|" & "IdAccount=" & IdAccount
	  sendData=CryptWithKey(sendData,Session("CryptKey"))
	  'response.write sendData
   %>
   
	<tr scope="col"> 
		<td>
			<input class="form-control" type="text" readonly value="<%=RsDet("DescProdotto")%>">
		</td>
		<td>
			<input class="form-control" id="CodiceProdotto<%=id%>" name="CodiceProdotto<%=id%>" type="text" value="<%=codiceProdotto%>">
		</td>
		
		
		<td><div class="form-check">
				<input id="checkProd<%=Id%>" <%=selezionato%> name="checkProd<%=Id%>" 
				type="checkbox" value = "S" class="big-checkbox" onclick="AssProdAccount(<%=Id%>,'<%=sendData%>')">
			</div>		
		</td>
	</tr> 
	<%	
      RsDet.MoveNext
   Loop
   RsDet.close
else
   response.write sendDatao & "  " & Session("CryptKey") & " - " & sendData
end if 
    %>	
    </table>
