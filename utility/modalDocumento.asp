<script>
function SearchDocConfirm()
{
   var d = $("#DocSearchIdSelected").val(); 
   if (d==0) {
     alert("Selezionare un documento");
	 return false;
   }
   // implementata nel modulo chiamante 
   mySearchDocConfirm(d);
}
</script>
<div class="modal fade" id="confirmModalDocumento"  aria-hidden="true" role="dialog">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">

        <h2>Documento Da Caricare </h2> 
        <button type="button" class="close" data-dismiss="modal">
          <span aria-hidden="true">×</span><span class="sr-only">Chiudi</span>
        </button>
      </div>

      <div class="modal-body"> 
        <div class="row  bg-light">
		   <div class="col-1"></div>
		   <div class="col-11">
              <div class="form-group ">
			     <%xx=ShowLabel("Ricerca")%>
					 <input type="text"  name="myInputSearchDoc" id="myInputSearchDoc" class="form-control" placeholder="Ricerca..">
                  </div>		
		   </div>
      </div>
	  
        <div>
             <table class="table table-bordered table-striped">
	         <thead>
		       <tr>
		        <th>Sel</th>
			    <th>Documento</th>
		       </tr>
	         </thead>
             <tbody id="tableDocSearch"> 
          <%
             Conta=0
             Rs.CursorLocation = 3 
			 q = "select * from Documento Where IdDocumentoInterno='' "
			 if ElencoIdDocumenti<>"" then 
			    Q = q & " and idDocumento not in (" & ElencoIdDocumenti & ")"  
			 end if 
			 q = Q & " Order By Descdocumento"

             Rs.Open q  , ConnMsde
			 'response.write q & " " & err.description 
		     'response.end 
			 if err.number > 0 then 
			    response.write err.description 
			 else			
                Do while not rs.eof and err.number=0
			       conta=conta+1
             %>
			 <tr>
			    <td><input name="MDOC_Selection" type="radio" onclick="$('#DocSearchIdSelected').val(<%=Rs("IdDocumento")%>)"></td>
			    <td><%=Rs("DescDocumento")%></td>
             </tr>
             <%
                   Rs.movenext
                loop 
                Rs.close 
             end if 
          %>
		     </tbody>
           </table>
        </div>          
      </div> 
      <input type="hidden" name="DocSearchIdSelected" id="DocSearchIdSelected" value="0">
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
        <%if conta>0 then %>
        <button type="button" class="btn btn-primary" onclick="SearchDocConfirm();";>Seleziona</button>
        <%end if %>
      </div>
    </div>
  </div>
</div>
