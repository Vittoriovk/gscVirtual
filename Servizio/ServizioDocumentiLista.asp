
<script>

function SDL_Attiva(id)  {
   modalAttivitaDocumento(id);
}

function SDL_ConfirmDelete(id)  {
   xx=myConfirm("Documento","Conferma Cancellazione","SDLDOCUDELE_" + id);
}
function SDL_ConfirmOk(id)  {
   xx=myConfirm("Documento","Conferma Approvazione","SDLDOCUOK_" + id);
}
function SDL_ConfirmKo(id)  {
   xx=myConfirmInfo("Documento","Documento non valido","SDLDOCUKO_" + id,"Motivo del ko");
}
function SDL_ConfirmReqS(id)  {
   xx=myConfirm("Documento","Documento obbligatorio","SDLDOCUREQS_" + id);
}
function SDL_ConfirmReqN(id)  {
   xx=myConfirm("Documento","Documento non obbligatorio","SDLDOCUREQN_" + id);
}
function SDL_ShowDocumento(idAtt,idNumAtt) {

  $("#DocSearchIdSelected").val("0");
  //attivo la ricerca
  
  $("#myInputSearchDoc").on("keyup", function() {
    var value = $(this).val().toLowerCase();
    $("#tableDocSearch tr").filter(function() {
      $(this).toggle($(this).text().toLowerCase().indexOf(value) > -1)
    });
  });

   xx=$('#confirmModalDocumento').modal('toggle');
}

function mySearchDocConfirm(idDocumento)
  {
    ImpostaValoreDi("Oper","SDLDOCUADD");
	ImpostaValoreDi("ItemToRemove",idDocumento);
    document.Fdati.submit();
}
</script>
<%
on error goto 0
'input
'OpDocAmm               = Operazioni ammesse 
'                         A = All : tutte le operazioni
'                         Q = mette il segno piu per aggiungere un documento
'                         V = pulsante di validazione se non validato
'                         D = pulsante di cancellazione 
'                         U = pulsante di upload 
'                         N = pulsante di non validazione 
'
'IdAttivita            = mostra l'elenco dei documenti per compagnia 
'IdNumAttivita         = richiesta di affidamento
'ShowAction            = mostra azioni oltre al documento se presente

MySqlDoc = ""
MySqlDoc = MySqlDoc & " select a.*,B.DescDocumento "
MySqlDoc = MySqlDoc & " from AttivitaDocumento A, Documento B "
MySqlDoc = MySqlDoc & " Where A.IdAttivita = '" & IdAttivita & "'"
MySqlDoc = MySqlDoc & " and A.IdNumAttivita = " & NumForDb(IdNumAttivita)
MySqlDoc = MySqlDoc & " And A.IdDocumento = B.IdDocumento"
MySqlDoc = MySqlDoc & " order By A.IdAttivitaDocumento"
'response.write MySqlDoc 
Set RsDoc = Server.CreateObject("ADODB.Recordset")
Set DsDoc = Server.CreateObject("ADODB.Recordset")

%>

	<div class="table-responsive"><table class="table"><tbody>
		<thead>
		<tr>
		    <th scope="col">Documento
			<%
			if Instr(OpDocAmm,"Q")>0 or instr(OpDocAmm,"A")>0 then
			   RiferimentoA="col-2;#;;2;plus;Carica Documento;;SDL_ShowDocumento('" & IdAttivita & "','" & IdNumAttivita & "');N"
			%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
			<%
			end if 
			%>
			
			</th>
		    <th scope="col">Stato</th>
            <th scope="col">Richiesto</th>			
			<th scope="col">Valido Dal</th>
			<th scope="col">Valido Al</th>
		    <th scope="col">Note</th>
			<th scope="col">azioni</th>
		</tr>
		</thead>
		<%
		if ShowResponsabilita=true	 then 
		%>
		<tr>
		    <td colspan="7" class="bg-info" scope="col">
			<b>La documentazione acquisita &egrave a solo titolo informativo.E' responsabilit&agrave del richiedente verificarne la correttezza.</b></td>
		<tr>

		<%
		end if 
        ContaDocAssenti=0
        ContaDocKo=0
   
		err.clear
		FlagAction=""
        RsDoc.CursorLocation = 3 
        RsDoc.Open MySqlDoc, ConnMsde
		if err.number=0 then 
		   NumDocToLoad=0
		   FlagAction="RICHIEDI"
		   Do While Not RsDoc.EOF and err.number = 0
		      IdAttivitaDocumento = RsDoc("IdAttivitaDocumento")
		      IdDocumento         = RsDoc("IdDocumento")
			  if ElencoIdDocumenti="" then 
			     ElencoIdDocumenti="0"
			  end if 
			  ElencoIdDocumenti   = ElencoIdDocumenti & "," & IdDocumento
			  linkDocumento       = ""
			  PathDocumento       = ""
			  IdTipoValidazione   = RsDoc("IdTipoValidazione")
			  NoteValidazione     = RsDoc("NoteValidazione")
			  DescTipoValidazione = "Da Caricare"
			  IdUpload            = RsDoc("IdUpload")
              descBreve           = ""
			  NomeDocumento       = ""
			  ValidoDal           = 0
              ValidoAl            = 0

			  if Cdbl(IdUpload)>0 then 
			     QSel = ""
                 Qsel = Qsel & " select * from Upload "
				 Qsel = Qsel & " Where IdUpload = " & IdUpload 
 
	             DsDoc.CursorLocation = 3 
                 DsDoc.Open Qsel, ConnMsde
                 response.write err.description
			     if err.number = 0 then 
			        if Not DsDoc.EOF then 
					   NomeDocumento      = DsDoc("NomeDocumento")
                       PathDocumento      = DsDoc("PathDocumento")
					   ValidoDal          = DsDoc("ValidoDal")
					   ValidoAl           = DsDoc("ValidoAl")
					   DescTipoValidazione = "Caricato"
				    end if 
			     end if 
			     DsDoc.close 
			     err.clear  
              else
			     IdTipoValidazione  ="NONRIC"
			     DescTipoValidazione="Da Caricare"
			  end if 
			  if IdTipoValidazione="VALIDO" then 
			     DescTipoValidazione="Validato"
			  end if 
			  if IdTipoValidazione="NONVAL" then 
			     DescTipoValidazione="Non Valido"
			  end if 

			  DescDocumento=RsDoc("DescDocumento")
			  if ucase(trim(DescDocumento))<>ucase(trim(descBreve)) and len(DescBreve)>0 then 
			     DescDocumento=DescDocumento & ":" & DescBreve
			  end if 
			  DescDocumento = DescDocumento & " " & Descrife
			  flagInvalid = false 
			  
		     %>


				   <%
				   Req="NO"
				   'response.write RsDoc("FlagObbligatorio") & " " & PathDocumento & " " & idTipoValidazione
				   if RsDoc("FlagObbligatorio")=1 then 
				      Req="SI"
					  'manca il documento
					  if PathDocumento="" then 
					     ContaDocAssenti=ContaDocAssenti+1
						 ContaDocKo     =ContaDocKo+1
						 flagInvalid = true 
					  else 
					     if idTipoValidazione<>"VALIDO" then 
						    'response.write "noo"
					        ContaDocKo=ContaDocKo + 1
							flagInvalid = true 
					     end if 
					  end if
                   else
                      if PathDocumento<>"" and idTipoValidazione<>"VALIDO" then 
					     ContaDocKo=ContaDocKo + 1
						 flagInvalid = true 
					  end if 
				   end if 
				   classeTR="" 
				   if flagInvalid then 
				      classeTR= "class='table-danger'"
				   end if 
				   %>
			  <tr <%=classeTR%> scope="col">
			    <td>
					<input class="form-control" type="text" readonly value="<%=DescDocumento%>">
				</td>
				<td width='20%'>
				    <input class="form-control" type="text" readonly value="<%=DescTipoValidazione%>">
				</td>
				<Td width='7%'>				   
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
				
			    <%if ShowAction = true then %>
				     <%if (instr(OpDocAmm,"U") > 0 or instr(OpDocAmm,"A")>0 ) and idTipoValidazione<>"VALIDO"  then 'upload 
				         RiferimentoA="col-2;#;;2;uplo;Carica Documento;;SDL_Attiva(" & IdAttivitaDocumento & ");N"
				     %>
				         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					 <%end if %>
					 
				     <%
				     if IsBackOffice() and (instr(OpDocAmm,"V") > 0 or instr(OpDocAmm,"A")>0 ) and idTipoValidazione<>"VALIDO" then 'validazione 
					    if Req="NO" then 
				           RiferimentoA=";#!;;2;requ;Richiedi;;SDL_ConfirmReqS(" & IdAttivitaDocumento & ");N"
						else
						   RiferimentoA=";#!;;2;remo;Non Richiedere;;SDL_ConfirmReqN(" & IdAttivitaDocumento & ");N"
						end if 
					 %>
				         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
 				     <%end if %>
				 
				     <%if (instr(OpDocAmm,"V") > 0  or instr(OpDocAmm,"A")>0 ) and PathDocumento<>"" and idTipoValidazione<>"VALIDO" then 'validazione 
				         RiferimentoA="col-2;#;;2;ok;Segna come valido;;SDL_ConfirmOk(" & IdAttivitaDocumento & ");N"
				     %>
				         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					 <%else 
				        if (instr(OpDocAmm,"N") > 0  or instr(OpDocAmm,"A")>0 ) then 'validazione 
				           RiferimentoA="col-2;#;;2;ko;Segna come non valido;;SDL_ConfirmKo(" & IdAttivitaDocumento & ");N"
				     %>
				         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					 <%   end if 
					   end if %>
				     <%if (instr(OpDocAmm,"D") > 0  or instr(OpDocAmm,"A")>0 ) then 'validazione 
				         RiferimentoA="col-2;#;;2;dele;Rimuovi;;SDL_ConfirmDelete(" & IdAttivitaDocumento & ");N"
				     %>
				         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					 <%end if %>
					 
  				  <%end if %>
				</td>
			  </tr>
		      <%
		      RsDoc.MoveNext
	       Loop
		End if 
		RsDoc.close 
		err.clear 
		%>


	</tbody></table></div>

<div id="divModalAttivitaDocumento"></div>
<script>

function MAD_Reload()
{
   document.Fdati.submit();
}

function modalAttivitaDocumento(id)
{
   var vp=$("#hiddenVirtualPath").val(); 
   var dataIn = "IdAttivitaDocumento=" + id;
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/modalAttivitaDocumento/modalCreate.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	    $("#divModalAttivitaDocumento").html(msg); 
		$('.mydatepicker').datepicker({
		inputFormat: ["dd/MM/yyy"],
		outputFormat: 'dd/MM/yyyy'
	    });
		$('#confirmModalAttivitaDocumento').modal('toggle');
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata upload create fallita, si prega di riprovare..." + thrownError);
      }
    });  
  
}

</script>
