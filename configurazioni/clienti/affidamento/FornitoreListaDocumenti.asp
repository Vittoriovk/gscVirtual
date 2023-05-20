<%
 
Set RsDoc = Server.CreateObject("ADODB.Recordset")
Set DsDoc = Server.CreateObject("ADODB.Recordset")

'response.write MySqlDoc

%>

	<div class="table-responsive"><table class="table"><tbody>
		<thead>
		<tr>
		    <th scope="col">Documento</th>
		    <th scope="col">Stato</th>
            <th scope="col">Richiesto</th>			
			<th scope="col">Valido Dal</th>
			<th scope="col">Valido Al</th>
		    <th scope="col">Note</th>
		</tr>
		</thead>
		<%
        ContaDocAssenti=0
        ContaDocKo=0
   
		err.clear
		FlagAction=""
        RsDoc.CursorLocation = 3 
        RsDoc.Open qListaDoc, ConnMsde
		if err.number=0 then 
		   NumDocToLoad=0
		   Do While Not RsDoc.EOF 
		      IdDocumento         = RsDoc("IdDocumento")
			  ElencoIdDocumenti = ElencoIdDocumenti & "," & IdDocumento
			  IdAccountDocumento  = RsDoc("IdAccountDocumento")
			  linkDocumento       = ""
			  PathDocumento       = ""
			  IdTipoValidazione   = ""
			  DescTipoValidazione = "Da Caricare"
			  IdUpload            = 0
              descBreve           = ""
			  NomeDocumento       = ""
			  ValidoDal           = 0
              ValidoAl            = 0
			  
			  if Cdbl(IdAccountDocumento)>0 then 
			     QSel = ""
                 Qsel = Qsel & " select A.IdTipoValidazione,A.NoteValidazione,b.* from AccountDocumento a,Upload B "
				 Qsel = Qsel & " Where A.IdAccountDocumento=" & IdAccountDocumento
				 Qsel = Qsel & " And A.IdUpload = b.IdUpload "
				 
                 'response.write Qsel 
	             DsDoc.CursorLocation = 3 
                 DsDoc.Open Qsel, ConnMsde
                 
			     if err.number = 0 then 
			        if Not Ds.EOF then 
					   IdTipoValidazione  = DsDoc("IdTipoValidazione")
					   NoteValidazione    = DsDoc("NoteValidazione")
					   DescTipoValidazione=funDoc_DescrizioneStatoDoc(IdTipoValidazione)
                       descBreve          = DsDoc("descBreve")
					   NomeDocumento      = DsDoc("NomeDocumento")
                       PathDocumento      = DsDoc("PathDocumento")
					   ValidoDal          = DsDoc("ValidoDal")
					   ValidoAl           = DsDoc("ValidoAl")
					   IdUpload           = DsDoc("IdUpload")
				    end if 
			     end if 
			     DsDoc.close 
			     err.clear   
			  end if 

			  
		     %>

			  <tr scope="col">
			    <td>
					<input class="form-control" type="text" readonly value="<%=RsDoc("DescDocumento")%>">
				</td>
				<td width='20%'>
				    <input class="form-control" type="text" readonly value="<%=DescTipoValidazione%>">
				</td>
				<Td width='7%'>
				   <%
				   Req="NO"
				   if RsDoc("FlagObbligatorio")=1 then 
				      Req="SI"
				   end if 
				   
				   %>
				   <input class="form-control" type="text" readonly value="<%=Req%>">
				</td>
				<td width='11%'>
				    <input class="form-control" type="text" readonly value="<%=Stod(ValidoDal)%>">
				</td>				
				<td width='11%'>
				    <input class="form-control" type="text" readonly value="<%=Stod(ValidoAl)%>">
				</td>				
				<td >
				    <input class="form-control" type="text" readonly value="<%=NoteValidazione%>">
				</td>				
				<td>
				<%Linkdocumento=PathDocumento%>
				<!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->
			  </tr>
		      <%
		      RsDoc.MoveNext
	       Loop
		End if 
		RsDoc.close 
		err.clear 
		%>


	</tbody></table></div>

