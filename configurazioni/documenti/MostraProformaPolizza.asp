<%
   idDocProF = LeggiCampo("Select * from Documento Where IdDocumentoInterno='PROFORMA'","IdDocumento")
   idDocProF = TestNumeroPos(idDocProF)
   idDocPoli = LeggiCampo("Select * from Documento Where IdDocumentoInterno='POLIZZA'","IdDocumento")
   idDocPoli = TestNumeroPos(idDocPoli)	 
   'leggo il proforma e la polizza se presenti 
   idUplDocProF=0
   PathDocProF =""
   qSel = ""
   qSel = qSel & " select * from Upload "
   qSel = qSel & " Where IdTabella='" & idTabella & "'" 
   qSel = qSel & " and IdTabellaKeyInt = " & IdTabellaKeyInt 
   qSel = qSel & " and IdTipoDocumento in (" & idDocProF & ")"
   Rs.CursorLocation = 3 
   Rs.Open qSel, ConnMsde
   if Err.number=0 then 
      do while not rs.eof
         if cdbl(Rs("IdTipoDocumento"))=cdbl(idDocProF) then
            idUplDocProF=Rs("IdUpload")
            PathDocProF =RS("PathDocumento")
         end if 
         rs.moveNext 
      loop
   end if 
   Rs.close 
   err.clear   
   
   idUplDocPoli=0
   PathDocPoli =""
   qSel = ""
   qSel = qSel & " select * from Upload "
   qSel = qSel & " Where IdTabella='" & idTabella & "'" 
   qSel = qSel & " and IdTabellaKeyInt = " & IdTabellaKeyInt 
   qSel = qSel & " and IdTipoDocumento in (" & idDocPoli & ")"
   Rs.CursorLocation = 3 
   Rs.Open qSel, ConnMsde
   if Err.number=0 then 
      do while not rs.eof
         if cdbl(Rs("IdTipoDocumento"))=cdbl(idDocPoli) then
            idUplDocPoli=Rs("IdUpload")
            PathDocPoli =RS("PathDocumento")
         end if 
         rs.moveNext 
      loop
   end if 
   Rs.close 
   err.clear   
%>
   <div class="table-responsive"><table class="table"><tbody>
   <thead>
   <tr>
   	<th scope="col">Documento</th>
   	<th scope="col">PDF</th>
       <th scope="col">Azioni</th>
   </tr>
   </thead>
   <tr>
      <td>Proforma</td>
      <td>
      <% if cdbl(idUplDocProF) = 0 then
            response.write "N.D." 
   	     else
            Linkdocumento=PathDocProF
   	  %>
   	  <!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->   	  

      <% end if%>
      </td>
      <td>
        <%
        RiferimentoA=";#;;2;uplo;Carica;;localGesAction('CALL_UPL','" & idUplDocProf & "');N"
   	    if cdbl(0 & FlagStatoFinale) = 0 then 	%>
   	         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
        <%end if %> 
      </td>
   </tr>
   <% if idUplDocPoli>0 or IdStatoServizio="PAGA" then %>
   <tr>
      <td>Polizza</td>
      <td>
      <% if cdbl(idUplDocPoli) = 0 then
                 response.write "N.D." 
       Linkdocumento=""
   	  else
       Linkdocumento=PathDocPoli
   	  %>
   	  <!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->   	  

           <% end if%>
      </td>
      <td>
        <%
        RiferimentoA=";#;;2;uplo;Carica;;localGesAction('CALL_POL','" & idUplDocPoli & "');N"
   	 if cdbl(0 & FlagStatoFinale) = 0 then %>
   	<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
   	<%end if %>
  
      </td>
   </tr>   
   <% end if %>

   </tbody></table></div>
   
