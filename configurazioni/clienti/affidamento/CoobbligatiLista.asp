<%
'input
'OpDocAmm               = Operazioni ammesse 
''                        C = permette di passare alla gestione coobbligati LocalAddNewCoob
'                         Q = mette + vicino a documento e chiama LocalAddRow
'                         O = mette ok sulla riga e chiama localVal
'                         G = mette manutenzione e chiama localGes passando IdDocumento da gestire
'                         X = forza cancellazione 
'                         N = Richiede il numero dei coobbligati 
'IdCompagnia                = mostra l'elenco dei documenti per compagnia 
'IdAffidamentoRichiestaComp = richiesta di affidamento
'NumCoobbligatiRichiesti    = quanti coobbligati devono essere inseriti
'ShowAction                 = mostra azioni oltre al documento se presente  
'ShowElencoCoob             = mostra elenco anche se non ci sono

MySqlDoc = ""
MySqlDoc = MySqlDoc & " select count(*) as tot "
MySqlDoc = MySqlDoc & " From AffidamentoRichiestaCompCoob"
MySqlDoc = MySqlDoc & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
NumCoobbligatiRichiestiCount = cdbl("0" & LeggiCampo(MySqlDoc,"tot"))

MySqlDoc = ""
MySqlDoc = MySqlDoc & " select * "
MySqlDoc = MySqlDoc & " From AffidamentoRichiestaCompCoob"
MySqlDoc = MySqlDoc & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
MySqlDoc = MySqlDoc & " Order By Ragsoc"
'response.write MySqlDoc 
CiSonoCoob = cdbl("0" & LeggiCampo(MySqlDoc,"IdAffidamentoRichiestaComp"))
CoobIsPresenteDocum = true 
if Cdbl(CiSonoCoob)>0 or ShowElencoCoob then 
   NoteValidazione=""
   Set RsDoc = Server.CreateObject("ADODB.Recordset")
%>
	<div class="row">
	   <div class="col-12 bg-primary text-white font-weight-bold">
	   Coobbligati
			<% 
			  if IsBackOffice() then %>
			  
			  <%if CoobCanAddRem then %>
			  
			  <a href="#" title="Aggiungi Coobbligato" onclick="LocalIncCoob();">
	             <i class="fa fa-2x fa-plus-square"></i>
              </a>
			  <%   if cdbl(NumCoobbligatiRichiesti)>0 then %>
			          <a href="#" title="Rimuovi Coobbligato" onclick="LocalDecCoob();">
	                     <i class="fa fa-2x fa-minus-square"></i>
                      </a>
			  <%   end if 
			    end if 
			  %>
			  richiesti: <%=NumCoobbligatiRichiesti%>
			  <%if NumCoobbligatiRichiesti>0 then %>
			  <div class="spinner-grow text-warning" role="status">
                 <span class="sr-only">Loading...</span>
              </div>
			  <%end if %>
			  inseriti:  <%=NumCoobbligatiRichiestiCount%>
			  <%if NumCoobbligatiRichiesti>0 then %>
			     <%if NumCoobbligatiRichiestiCount < NumCoobbligatiRichiesti then %>
			  <div class="spinner-grow text-danger" role="status">
                 <span class="sr-only">Loading...</span>
              </div>			  
			     <%else%>
			  <div class="spinner-grow text-success" role="status">
                 <span class="sr-only">Loading...</span>
              </div>			  
				 <%end if %>
			  <%end if %>
			  <%
			  else 
			     if cdbl(NumCoobbligatiRichiesti)>0 then 
			        response.write " richiesti: " & NumCoobbligatiRichiesti
					%>
			  <div class="spinner-grow text-warning" role="status">
                 <span class="sr-only">Loading...</span>
              </div>
					<%
				    response.write "  inseriti: " & NumCoobbligatiRichiestiCount
					if cdbl(NumCoobbligatiRichiesti)>cdbl(NumCoobbligatiRichiestiCount) then
					%>
			           <div class="spinner-grow text-danger" role="status">
                         <span class="sr-only">Loading...</span>
                       </div>			  
                    <%else%>
			           <div class="spinner-grow text-success" role="status">
                         <span class="sr-only">Loading...</span>
                       </div>			  
					<%
					end if 
			     end if 
			     if cdbl(NumCoobbligatiRichiesti)<=cdbl(NumCoobbligatiRichiestiCount) then
			        CoobPresentiTutti = true 
                 end if 
			     if Instr(OpDocAmm,"C")>0 then
                    RiferimentoA="col-2;#;;2;clie;Gestione Coobbligato;;LocalAddNewCoob();N"
			%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
			<% 
			     end if 
			  end if %>  
	   
	   </div>
	</div>
			
	<div class="table-responsive"><table class="table"><tbody>
		<thead>
		<tr>
		    <th scope="col">Nominativo
			<% if ShowAction = true and cdbl(IdCompagnia)>0 and Instr(OpDocAmm,"I")>0 then
			      IdRichiestaAffComp = "0" & LeggiCampo(MySqlSub,"IdAffidamentoRichiestaComp")
			      if esisteCoobbligatoxAccount (idAccountcliente,cdbl(IdRichiestaAffComp)) then 
                     RiferimentoA="col-2;#;;2;plus;Carica Coobbligato;;LocalAddRowCoob();N"
			%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
			<%    end if 
			  end if %>  
			</th>
		    <th scope="col">CF</th>
            <th scope="col">PI</th>
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
		      IdAccountCoobbligato = RsDoc("IdAccountCoobbligato")
			  FlagValidato         = Rsdoc("FlagValidato")
			  FlagDocCompleta      = true  
			  'controllo se è validato su account 
			  if Cdbl(FlagValidato) = 0 then 
			     qSelTemp = ""
				 qSelTemp = qSelTemp & "select * from AccountCoobbligato "
				 qSelTemp = qSelTemp & "where IdAccountCoobbligato = " & IdAccountCoobbligato
			     FlagValidato = "0" & LeggiCampo(qSelTemp,"FlagValidato")
			  end if
			  'se non è validato controllo che ci siano tutti i documenti per poter
			  'eventualmente mandare la richiesta 
			  if Cdbl(FlagValidato) = 0 then 
			     qSelTemp = ""
				 qSelTemp = qSelTemp & " Select A.IdAccountDocumento "
				 qSelTemp = qSelTemp & " From AccountDocumento A left join Upload C "
				 qSelTemp = qSelTemp & " on a.idUpload = C.IdUpload "
				 qSelTemp = qSelTemp & " inner join AccountCoobbligato D"
				 qSelTemp = qSelTemp & " on D.IdAccountCoobbligato = " & IdAccountCoobbligato
				 qSelTemp = qSelTemp & " Where a.IdAccount = D.IdAccount" 
				 qSelTemp = qSelTemp & " and A.TipoRife = 'COOB' and A.IdRife = D.IdAccountCoobbligato "
				 tempData = Cdbl("0" & leggiCampo(qSelTemp,"IdAccountDocumento"))
				 'manca la documentazione 
				 if Cdbl(tempData)=0 then 
				    FlagDocCompleta     = false 
					CoobIsPresenteDocum = false 
				 else 

			        qSelTemp = ""
				    qSelTemp = qSelTemp & " Select A.IdAccountDocumento "
				    qSelTemp = qSelTemp & " From AccountDocumento A left join Upload C "
				    qSelTemp = qSelTemp & " on a.idUpload = C.IdUpload "
				    qSelTemp = qSelTemp & " inner join AccountCoobbligato D"
				    qSelTemp = qSelTemp & " on D.IdAccountCoobbligato = " & IdAccountCoobbligato
				    qSelTemp = qSelTemp & " Where a.IdAccount = D.IdAccount" 
				    qSelTemp = qSelTemp & " and A.TipoRife = 'COOB' and A.IdRife = D.IdAccountCoobbligato "
                    qSelTemp = qSelTemp & " and A.FlagObbligatorio = 1 and IsNull(C.PathDocumento,'ND') = 'ND'"
				 
				    tempData = Cdbl("0" & leggiCampo(qSelTemp,"IdAccountDocumento"))
				    if Cdbl(tempData)>0 then 
				       FlagDocCompleta     = false 
					   CoobIsPresenteDocum = false 
			        end if 
			     end if 
			  end if 
		      IdCoob       = RsDoc("IdAffidamentoRichiestaCompCoob")
			  ElencoRagSoc = ElencoRagSoc & ",'" & apici(Rsdoc("RagSoc")) & "'"
			  
			  Note = ""
			  consentiVal = false
			  if flagValidato then 
			     Note = "Validato"
			  else 
			     consentiVal = true 
			     Note = "Da validare "
				 if FlagDocCompleta = false then 
				    Note = Note & " - documentazione da caricare "
				 end if 
			  end if 
		     %>

			  <tr scope="col">
			    <td>
					<input class="form-control" type="text" readonly value="<%=RsDoc("RagSoc")%>">
				</td>
			    <td>
					<input class="form-control" type="text" readonly value="<%=RsDoc("PI")%>">
				</td>
			    <td>
					<input class="form-control" type="text" readonly value="<%=RsDoc("CF")%>">
				</td>
				<td >
				    <input class="form-control" type="text" readonly value="<%=Note%>">
				</td>				
				<td>

			    <%
				if OpDocAmm="VAL_COOB" and consentiVal = true then 
				   RiferimentoA="col-2;#;;2;manu;Gestisci Coobbligato;;LocalGesRowCoob(" & RsDoc("IdAccountCoobbligato") & ");N"
					%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					<%
				end if 
				
				if ShowAction = true then 
					if Instr(OpDocAmm,"I")>0 and RsDoc("FlagValidato")=0 then
					   RiferimentoA="col-2;#;;2;dele;Rimuovi Coobbligato;;LocalRemRowCoob(" & IdCoob & ");N"
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
		CoobPresenteDocum = CoobIsPresenteDocum
		err.clear 
		%>


	</tbody></table></div>

<%end if %>