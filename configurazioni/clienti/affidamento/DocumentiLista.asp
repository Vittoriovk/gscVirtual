<%
'input
'OpDocAmm               = Operazioni ammesse 
'                         Q = mette + vicino a documento e chiama LocalAddRow
'                         O = mette ok sulla riga e chiama localVal
'                         G = mette manutenzione e chiama localGes passando IdDocumento da gestire
'                         X = forza cancellazione 
'
'IdCompagnia            = mostra l'elenco dei documenti per compagnia 
'IdAffidamentoRichiesta = richiesta di affidamento
'ShowAction             = mostra azioni oltre al documento se presente 
'ShowColors             = mostra i colori per gli elementi
'                         verdino = ok 
'                         rosso   = ko
'                         giallo  = da gestire     

MySqlSub = ""
MySqlSub = MySqlSub & " Select IdAffidamentoRichiestaComp "
MySqlSub = MySqlSub & " from AffidamentoRichiestaComp "
MySqlSub = MySqlSub & " Where IdAffidamentoRichiesta = " & IdAffidamentoRichiesta
MySqlSub = MySqlSub & " And IdCompagnia = " & IdCompagnia 
 

MySqlDoc = ""
MySqlDoc = MySqlDoc & "select C.IdDocumento,C.DescDocumento,A.TipoRife,A.IdRife "
MySqlDoc = MySqlDoc & ",Max(FlagObbligatorio) as FlagObbligatorio"
MySqlDoc = MySqlDoc & ",Max(FlagDataScadenza) as FlagDataScadenza"
MySqlDoc = MySqlDoc & ",Max(IdAccountDocumento) as IdAccountDocumento"
MySqlDoc = MySqlDoc & " From AffidamentoRichiestaCompDoc a, Documento C"
MySqlDoc = MySqlDoc & " Where A.IdAffidamentoRichiestaComp in (" & MySqlSub & ") "
MySqlDoc = MySqlDoc & " and   A.IdDocumento = C.IdDocumento"
if FiltraDocumenti<>""  then 
   MySqlDoc = MySqlDoc & " and   A.IdDocumento in (" & FiltraDocumenti & ")"
end if  
MySqlDoc = MySqlDoc & " group by C.IdDocumento,C.DescDocumento,A.TipoRife,A.IdRife"
  
Set RsDoc = Server.CreateObject("ADODB.Recordset")
Set DsDoc = Server.CreateObject("ADODB.Recordset")

'response.write MySqlDoc

%>
            <div class="row">
               <div class="col-12 bg-primary text-white font-weight-bold">
               Documentazione richiesta per affidamento
               </div>
            </div>
            
    <div class="table-responsive"><table class="table"><tbody>
        <thead>
        <tr>
            <th scope="col">Documento
            <%
            if Instr(OpDocAmm,"Q")>0 then
               RiferimentoA="col-2;#;;2;plus;Carica Documento;;LocalAddRow();N"
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
        ContaDocAssenti=0
        ContaDocKo=0
   
        err.clear
        FlagAction=""
        RsDoc.CursorLocation = 3 
        RsDoc.Open MySqlDoc, ConnMsde
        if err.number=0 then 
           NumDocToLoad=0
           FlagAction="RICHIEDI"
           Do While Not RsDoc.EOF 
              IdDocumento         = RsDoc("IdDocumento")
              if TipoRife="" then 
                 ElencoIdDocumenti = ElencoIdDocumenti & "," & IdDocumento
              end if 
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
              TipoRife            = RsDoc("TipoRife")
              IdRife              = RsDoc("Idrife")
              DescRife            = ""
              NoteValidazione     = ""
              'response.write "ttt=" & TipoRife 
              if TipoRife="COOB" then 
                 qSelRif = ""
                 qSelRif = qSelRif & " select * from AffidamentoRichiestaCompCoob "
                 qSelRif = qSelRif & " Where IdAffidamentoRichiestaComp in (" & MySqlSub & ") "
                 qSelRif = qSelRif & " and IdAccountCoobbligato = " & IdRife  
                 'response.write qSelRif
                 DescRife = LeggiCampo(qSelRif,"RagSoc") 
              end if 
              
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
              else
                 IdTipoValidazione  ="NONRIC"
                 DescTipoValidazione="Da Caricare"
              end if 

              DescDocumento=RsDoc("DescDocumento")
              'response.write "ecco1:" & DescDocumento & "-" & DescBreve 
              'se il documento non contiene la descrizione breve 
              if instr(ucase(trim(DescDocumento)),ucase(trim(descBreve)))=0 then 
                 if instr(ucase(trim(descBreve)),ucase(trim(DescDocumento)))=0 then 
                    DescDocumento=DescDocumento & ":" & DescBreve
                 else
                    DescDocumento=DescBreve
                 end if 
              end if 
              
              'la descrizione non contiente il riferimento 
              'response.write "ecco1:" & DescDocumento & ".." & DescRife 
              if Descrife<>"" and instr(ucase(trim(DescDocumento)),ucase(trim(Descrife)))=0 then
                 DescDocumento = DescDocumento & " " & Descrife              
              end if 

              bgColor = ""
			  if ShowColors then 
			     if IdTipoValidazione="VALIDO" then 
				    bgcolor="bgcolor='#CAFFE0'"
				 elseif IdTipoValidazione="NONVAL" then 
				    bgcolor="bgcolor='#FF9A9A'"
				 else 
				    bgcolor="bgcolor='#FFFFE0'"
				 end if 
			  end if 
             %>

              <tr scope="col" <%=bgColor%>>
                <td>
                    <input class="form-control" type="text" readonly value="<%=DescDocumento%>">
                </td>
                <td width='20%'>
                    <input class="form-control" type="text" readonly value="<%=DescTipoValidazione%>">
                </td>
                <Td width='7%'>
                   <%
                   Req="NO"
                   'response.write RsDoc("FlagObbligatorio") & " " & PathDocumento & " " & idTipoValidazione
                   if RsDoc("FlagObbligatorio")=1 then 
                      Req="SI"
                      'manca il documento
                      if PathDocumento="" then 
                         ContaDocAssenti=ContaDocAssenti+1
                         ContaDocKo     =ContaDocKo+1
                      else 
                         if idTipoValidazione<>"VALIDO" then 
                            'response.write "noo"
                            ContaDocKo=ContaDocKo + 1
                         end if 
                      end if
                   else
                      if PathDocumento<>"" and idTipoValidazione<>"VALIDO" then 
                         ContaDocKo=ContaDocKo + 1
                      end if 
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
                <td bgcolor="#F5F5F5">
                <%Linkdocumento=PathDocumento%>
                <!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->
                
                <%
                
                if ShowAction = true then 
                    if Instr(OpDocAmm,"I")>0 and cdbl(IdUpload)=0 then
                       RiferimentoA="col-2;#;;2;uplo;Carica Documento;;localInsDoc(" & idDocumento & "," & RsDoc("FlagObbligatorio") & "," & RsDoc("FlagDataScadenza") & ");N"
                    %>
                    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                    <%
                       RiferimentoA="col-2;#;;2;comp;Seleziona da Cassetto;;localSelCas(" & idDocumento & "," & IdAccountCliente & ",'" & TipoRife & "','" & IdRife & "');N"
                    %>
                    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                    
                    <%end if %>
                    <%
                    if Instr(OpDocAmm,"U")>0 and Cdbl(IdAccountDocumento)>0 and IdTipoValidazione="NONRIC" then
                       RiferimentoA="col-2;#;;2;upda;Aggiorna;;localUpd(" & IdDocumento & "," & IdAccountDocumento & ");N"
                    %>
                    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->                
                    <%end if %>
                    <%
                    if Instr(OpDocAmm,"X")>0 or (Instr(OpDocAmm,"D")>0 and Cdbl(IdAccountDocumento)>0 and IdTipoValidazione="NONRIC") then
                       if Instr(OpDocAmm,"X")>0 then 
                          RiferimentoA="col-2;#;;2;dele;Cancella;;localDel(" & IdDocumento & ");N"
                       else 
                          RiferimentoA="col-2;#;;2;dele;Cancella;;localDel(" & IdAccountDocumento & ");N"
                       end if 
                    %>
                    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->                
                    <%end if %>                
                    <%
                    ' validazione 
                    if Instr(OpDocAmm,"O")>0 and Cdbl(IdAccountDocumento)>0 and instr("_DAVALI_INVALI_NONRIC_","_" & IdTipoValidazione & "_")>0  and PathDocumento<>"" then
                       RiferimentoA="col-2;#;;2;ok;Valido;;localVal(" & IdAccountDocumento & ");N"
                    %>
                    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->                
                    <%end if %>                                    
                    <%
                    ' validazione 
                    vai=0
                    reg=0
                    if Session("LoginTipoUtente")=ucase("BackO") then 
                       vai=1
                    else 
                       if (isCliente() or isCollaboratore()) and (PathDocumento="" or instr("NONRIC_NONVAL",IdTipoValidazione)>0) then
                          vai=1
                       end if 
                       if (isCliente() or isCollaboratore()) and instr("NONVAL",IdTipoValidazione)>0 and Linkdocumento<>"" then
                          reg=0
                       end if    
                    end if 
                    
                    if Instr(OpDocAmm,"G")>0 and vai=1 then 'and cdbl(IdAccountDocumento)>0 then
                       RiferimentoA="col-2;#;;2;manu;Gestione;;localGes(" & IdDocumento & ",'" & TipoRife & "','" & IdRife & "');N"
                    %>
                    <!--#include virtual="/gscVirtual/include/Anchor.asp"-->

                    <%
                        if cdbl(IdUpload)=0 then
                           RiferimentoA="col-2;#;;2;comp;Seleziona da Cassetto;;localSelCas(" & idDocumento & "," & IdAccountCliente & ",'" & TipoRife & "','" & IdRife & "');N"
                    %>
                          <!--#include virtual="/gscVirtual/include/Anchor.asp"-->

                    
                    <%  end if 
                    end if %>                                    
                    <%if Instr(OpDocAmm,"G")>0 and reg=1 then
                         RiferimentoA="col-2;#;;2;ok;Reinvia;;reinviaGes(" & IdAccountDocumento & ");N"
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

