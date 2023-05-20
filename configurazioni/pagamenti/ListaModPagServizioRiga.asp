<tr scope="col">
   <td>
   <%
    idPP=TipoPaga & "_" & CodePaga
	disab = " disabled "
	check = " "
    if (cdbl(ImptDaPagare) < cdbl(ImptDisp)) and canSelModPaga=true then
	   CheckOk = true 
	   disab = " "
	end if 
    if (TipoPaga="ESTE") and canSelModPaga=true then
	   CheckOk = true 
	   disab = " "
	end if 	
	if TipoUten = "CLIE" and TipoPagaClieSelected=TipoPaga then 
	   check = " Checked "
	end if 
	if TipoUten = "CLIE" and TipoPagaClieSelected="" and check=" " then 
	   check = " Checked "
	end if 	
	if TipoUten = "REQU" and TipoPagaRequSelected=TipoPaga then 
	   check = " Checked "
	end if 
	if TipoUten = "REQU" and TipoPagaRequSelected="" and check=" " then 
	   check = " Checked "
	end if 	
	
   %>
   <div class="form-check">
       <input class="form-check-input" <%=disab%> type="radio" <%=check%>
	   name="ListaModPagServizio<%=TipoUten%>" id="ListaModPagServizio<%=idPP%>" value="<%=TipoPaga%>">
   </div>
      

   </td>
   <td>
        <input class="form-control" type="text" readonly value="<%=LMP_opt%>">
   </td>
   <%if TipoPaga="ESTE" then %>
   <td colspan="6">
        <input class="form-control " type="text" readonly value="Pagamento su sistema convenzionato">
   </td>		 
   
   <%else%>
   <td>
        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptTota,2)%>">
   </td>		 
   <td>
        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptImpe,2)%>">
   </td>		 
   <td>
        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptUtil,2)%>">
   </td>	   
   <td>
        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptDisp,2)%>">
   </td>		 
   <td>
        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptValo,2)%>">
   </td>
   <%end if %>   
</tr>