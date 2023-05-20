<script>

function localContattoSel()
{
	var d=$('#contattoCampoDest').val();
	var s=$('input[name="contattoCampoNome"]:checked').val();
	$('#'+d).val(s);
	$('#btnContattoClose').click();
}
</script>

<input type="hidden" name="contattoCampoDest" id="contattoCampoDest" value="<%=campoxValore%>"> 
<div class="modal fade" id="contattoModal"  aria-hidden="true" role="dialog">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">

        <h2>Selezionare Contatto</h2> 
        <button type="button" class="close" id="btnContattoClose" data-dismiss="modal">
          <span aria-hidden="true">×</span><span class="sr-only">Chiudi</span>
        </button>
      </div>
      <div class="modal-body"> 
		<div>	  
      <%
	  Set RsContatto = Server.CreateObject("ADODB.Recordset")
	  MyContQ = ""
	  MyContQ = MyContQ & " select * from AccountContatto "
	  MyContQ = MyContQ & " where IdAccount = " & contattoIdAccount
      MyContQ = MyContQ & " and   IdTipoContatto in (" & contattoTipo & ")"
      MyContQ = MyContQ & " order by flagPrincipale desc"  
'response.write MyContQ
      RsContatto.CursorLocation = 3
      RsContatto.Open MyContQ, ConnMsde 
      LeggiContatti=true 
	  Conta=0
      If Err.number<>0 then	
       	 LeggiContatti=false
      elseIf RsContatto.EOF then	
         LeggiContatti=false
		 RsContatto.close 
      End if
	  if LeggiContatti then 
	     
	     do while not RsContatto.eof 
		    conta=conta+1
			checked=""
			if conta=1 then 
			   checked=" checked "
			end if 
		 %>
		  <div class="form-check">
			<input name="contattoCampoNome" type="radio" id="radio<%=conta%>"  value="<%=RsContatto("DescContatto")%>" <%=checked%>>
			<label for="radio<%=conta%>"><%=RsContatto("DescContatto")%></label>
		  </div>		 
		 <%
		    RsContatto.moveNext 
		 loop  
		 RsContatto.close
	  end if 
	  if conta=0 then 
	     response.write "<h2>Nessuna mail in archivio</h2> "
	  end if 
      
	  %>
		</div>		  
      </div> 

	  
 
      <%if Conta>0 then %>
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
        <button type="button" class="btn btn-primary" onclick="localContattoSel();";>Seleziona</button>
      </div>
	  <%end if %>
    </div>
  </div>
</div>
