<script>
function GestioneAllegato()
{
    ImpostaValoreDi("Oper","CALL_GALL");
    document.Fdati.submit();
}
function CancellaAllegato()
{
    ImpostaValoreDi("Oper","DELE_GALL");
    document.Fdati.submit();
}
</script>

<%
   if Oper="DELE_GALL" then 
      xx=RemoveSwap()
      IdUpload    = cdbl("0" & Request("idUplDocToShow")) 
	  if Cdbl(IdUpload)>0 then 
	     ConnMsde.execute "Delete From Upload Where IdUpload=" & IdUpload
      end if   
   end if 
   'deve ricevere IdDocToShow e NomeDocToShow
   'leggo il doc se presenti 
   idUplDocToShow=0
   PathDocToShow =""
   qSel = ""
   qSel = qSel & " select * from Upload "
   qSel = qSel & " Where IdTabella='" & idTabella & "'" 
   qSel = qSel & " and IdTabellaKeyInt = " & IdTabellaKeyInt 
   qSel = qSel & " and IdTipoDocumento in (" & IdDocToShow & ")"
   Rs.CursorLocation = 3 
   Rs.Open qSel, ConnMsde
   if Err.number=0 then 
      do while not rs.eof
         if cdbl(Rs("IdTipoDocumento"))=cdbl(IdDocToShow) then
            idUplDocToShow=Rs("IdUpload")
            PathDocToShow =RS("PathDocumento")
         end if 
         rs.moveNext 
      loop
   end if 
   Rs.close 
   err.clear 
   
%>
   <input type="hidden" name="IdDocToShow"    value="<%=IdDocToShow%>">
   <input type="hidden" name="idUplDocToShow" value="<%=idUplDocToShow%>">
   <%if mostraDoc=true or Cdbl(idUplDocToShow)>0 then%>
   <div class="table-responsive"><table class="table"><tbody>
   <thead>
   <tr>
   	<th scope="col">Documento</th>
   	<th scope="col">PDF</th>
	<%if mostraDoc=true then %>
       <th scope="col">Azioni</th>
	<%end if %>
   </tr>
   </thead>
   <tr>
      <td><%=NomeDocToShow%></td>
      <td>
      <% if cdbl(idUplDocToShow) = 0 then
            response.write "N.D." 
   	     else
            Linkdocumento=PathDocToShow
   	  %>
   	  <!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->   	  

      <% end if%>
      </td>
	  <%if mostraDoc=true then %>
      <td>
        
   	<%if cdbl(0 & FlagGestisci) = 1 then %>
	   <%RiferimentoA=";#;;2;uplo;Carica;;GestioneAllegato();N"%>
   	   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
	   <% if cdbl(idUplDocToShow) > 0 then
	      RiferimentoA=";#;;2;dele;Rimuovi;;CancellaAllegato();N"%>
   	      <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
   	<%    end if
	  end if  
	%>
  
      </td>	  
	  <%end if %>
   </tr>   
   </tbody></table></div>
   <%end if %>
   
