<%
'input
'
'IdAffidamentoRichiesta = richiesta di affidamento
'ShowAction             = mostra azioni  


'
'controllo se il cliente ha Certificazioni altrimenti non espongo 
QntaCert = "0" & LeggiCampo("Select * from AccountCertificazione Where IdAccount=" & idAccountcliente ,"IdAccount")

if cdbl(QntaCert)>0 then 

	MySqlSub = ""
	MySqlSub = MySqlSub & " Select IdAffidamentoRichiestaComp "
	MySqlSub = MySqlSub & " from AffidamentoRichiestaComp "
	MySqlSub = MySqlSub & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
	if Cdbl(IdCompagnia)>0 then 
	   MySqlSub = MySqlSub & " And IdCompagnia = " & IdCompagnia 
	end if 

	MySqlDoc = ""
	MySqlDoc = MySqlDoc & " select a.*,B.DescEstesaCertificazione,B.DescBreveCertificazione "
	MySqlDoc = MySqlDoc & " From AffidamentoRichiestaCompCert a, Certificazione B "
	MySqlDoc = MySqlDoc & " Where IdAffidamentoRichiestaComp in (" & MySqlSub & ") "
	MySqlDoc = MySqlDoc & " and a.IdCertificazione = B.idCertificazione "
	MySqlDoc = MySqlDoc & " Order By B.DescBreveCertificazione"

	NoteValidazione=""

	Set RsDoc = Server.CreateObject("ADODB.Recordset")
	'response.write MySqlDoc

%>
			<div class="row">
			   <div class="col-12 bg-primary text-white font-weight-bold">
			   Certificazioni 
			   </div>
			</div>
			
	<div class="table-responsive"><table class="table"><tbody>
		<thead>
		<tr>
		    <th scope="col">Certificazione
			<% if ShowAction = true and Instr(OpDocAmm,"I")>0 then
			   RiferimentoA="col-2;#;;2;plus;Carica Certificazione;;LocalAddRowCert();N"
			%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
			<%end if %>  
			</th>
            <th scope="col">Stato</th>
			<%if OpDocAmm<>"" then %>
			<th scope="col">azioni</th>
			<%end if %>
		</tr>
		</thead>
		<%

		err.clear
		FlagAction=""
        RsDoc.CursorLocation = 3 
        RsDoc.Open MySqlDoc, ConnMsde
		if err.number=0 then 
		   NumDocToLoad=0
		   FlagAction="RICHIEDI"
		   Do While Not RsDoc.EOF 
		      IdCert = RsDoc("IdAffidamentoRichiestaCompCert")
		     %>

			  <tr scope="col">
			    <td>
					<input class="form-control" type="text" readonly value="<%=RsDoc("DescBreveCertificazione")%>">
				</td>
				<td >
				    <input class="form-control" type="text" readonly value="<%=Rs("Note")%>">
				</td>				
				<td>

			    <%
				if ShowAction = true then 
					if Instr(OpDocAmm,"I")>0 and RsDoc("FlagValidato")=0 then
					   RiferimentoA="col-2;#;;2;dele;Rimuovi Certificato;;LocalRemRowCert(" & IdCert & ");N"
					%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					<%end if %>				

				</td>
                <% end if %>
			  </tr>
		      <%
		      RsDoc.MoveNext
	       Loop
		End if 
		RsDoc.close 
		err.clear 
		%>


	</tbody></table></div>

<%end if %>