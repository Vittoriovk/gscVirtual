<%

   'modulo trasferimento da scaricare : lo carica il back office 
   IdDocTraD = LeggiCampo("Select * from Documento Where IdDocumentoInterno='MODULO_TRASF_D'","IdDocumento")
   IdDocTraD = TestNumeroPos(IdDocTraD)	    
   
   'modulo trasferimento da caricare  : lo carica il cliente    
   IdDocTraU = LeggiCampo("Select * from Documento Where IdDocumentoInterno='MODULO_TRASF_U'","IdDocumento")
   IdDocTraU = TestNumeroPos(IdDocTraU)

   'leggo upload e download se presenti  
   idUplDocTraU=0
   PathDocTraU =""
   qSel = ""
   qSel = qSel & " select * from Upload "
   qSel = qSel & " Where IdTabella='" & idTabella & "'" 
   qSel = qSel & " and IdTabellaKeyInt = " & IdTabellaKeyInt 
   qSel = qSel & " and IdTipoDocumento in (" & IdDocTraU & ")"
   Rs.CursorLocation = 3 
   Rs.Open qSel, ConnMsde
   if Err.number=0 then 
      do while not rs.eof
         if cdbl(Rs("IdTipoDocumento"))=cdbl(IdDocTraU) then
            idUplDocTraU=Rs("IdUpload")
            PathDocTraU =RS("PathDocumento")
         end if 
         rs.moveNext 
      loop
   end if 
   Rs.close 
   err.clear   
   
   idUplDocTraD=0
   PathDocTraD =""
   qSel = ""
   qSel = qSel & " select * from Upload "
   qSel = qSel & " Where IdTabella='" & idTabella & "'" 
   qSel = qSel & " and IdTabellaKeyInt = " & IdTabellaKeyInt 
   qSel = qSel & " and IdTipoDocumento in (" & IdDocTraD & ")"
   'response.write qSel 
   
   Rs.CursorLocation = 3 
   Rs.Open qSel, ConnMsde
   if Err.number=0 then 
      do while not rs.eof
         if cdbl(Rs("IdTipoDocumento"))=cdbl(IdDocTraD) then
            idUplDocTraD=Rs("IdUpload")
            PathDocTraD =RS("PathDocumento")
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
   <%if ShowDownload then %>
   <tr>
      <td>Trasferimento Da Scaricare</td>
      <td>
      <% if cdbl(idUplDocTraD) = 0 then
            response.write "N.D." 
   	     else
            Linkdocumento=PathDocTraD
   	  %>
   	  <!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->   	  

      <% end if%>
      </td>
      <td>
        <%if IsBackOffice() then
        RiferimentoA=";#;;2;uplo;Carica;;upload('" & idUplDocTraD & "');N"
   	    if cdbl(0 & FlagStatoFinale) = 0 then 	%>
   	         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
        <%end if 
		end if %> 
      </td>
   </tr>
   <%end if %>
   <%if showUpload = true then %>
   
   <tr>
      <td>Trasferimento Firmato</td>
      <td>
      <% if cdbl(idUplDocTraU) = 0 then
            response.write "N.D." 
   	     else
            Linkdocumento=PathDocTraU
   	  %>
   	  <!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->   	  

      <% end if%>
      </td>
      <td>
        <%if IsBackOffice() = false then
        RiferimentoA=";#;;2;uplo;Carica;;uploadd('" & idUplDocTraU & "');N"
   	    if cdbl(0 & FlagStatoFinale) = 0 then 	%>
   	         <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
        <%end if 
		end if %> 
      </td>
   </tr>
   
   
   
   <%end if %>

   </tbody></table></div>
   
