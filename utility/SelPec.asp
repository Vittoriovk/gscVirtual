<!--#include virtual="/gscVirtual/include/includeStd.asp"-->

<%
session("EsitoCallSelIndirizzo") = "" 
IdAccount = session("params_IdAccount")
'lista dei campi da valorizzare tipo 
'  Campo=nome_delcampo;
ListaCampi= Session("params_ListaCampi")
ArCampi = split(ListaCampi,";")
c_Pec = ""

for J=lbound(ArCampi) to ubound(ArCampi)-1
    str=ArCampi(j)
	'response.write str
	ArDett = split(str,":")
	if ubound(ArDett)=1 then 
	   nome=ucase(trim(ArDett(0)))
	   valo=trim(ArDett(1))
	   if Nome="PEC" then 
	      c_Pec=valo
	   end if 
   end if 
next 


prefix = "params_"

%>

<script>

function confirmModalProcedi()
{
	
	var s=$('input[name="sel"]:checked').val();

	var pec = $('#params_Pec' + s).val();
	d=$('#c_Pec').val();
	try {
		$('#'+d).val(pec);
	} catch (error) {
	}	
	$('#dismissModalGenerica').click();
}
</script>
      <input type="hidden" name="c_Pec" id="c_Pec" value="<%=c_Pec%>">
      <%
	  prefix = "params_"
	  Set Rs = Server.CreateObject("ADODB.Recordset")
	  MyContQ = ""
	  MyContQ = MyContQ & " select * from AccountContatto A "
	  MyContQ = MyContQ & " inner join TipoContatto B on A.IdTipoContatto = B.IdTipoContatto "
	  MyContQ = MyContQ & " Where A.IdAccount = " & IdAccount
	  MyContQ = MyContQ & " and   A.IdTipoContatto = 'PECC'"
      MyContQ = MyContQ & " order by Ordine"  
'response.write MyContQ
      Rs.CursorLocation = 3
      Rs.Open MyContQ, ConnMsde 
      LeggiContatti=true 
	  Conta=0
      If Err.number<>0 then	
       	 LeggiContatti=false
      elseIf Rs.EOF then	
         LeggiContatti=false
		 Rs.close 
      End if
	  if LeggiContatti then 
	     
	     do while not Rs.eof 
		    conta=conta+1
			checked=""
			if conta=1 then 
			   checked=" checked "
			end if 
			id = Rs("IdAccountContatto")
		 %>
			<div class="row">
	
			   <div class="col-5">
                  <div class="form-group ">
				     <input name="sel" type="radio" id="radio<%=conta%>"  value="<%=id%>" <%=checked%>>
				     <%xx=ShowLabel("PEC")
					   nn=prefix & "Pec" & id
					 %>
					 
					 <input type="text" class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("DescContatto")%>" >
                  </div>		
			   </div>
			   <div class="col-5">
                  <div class="form-group ">
				     <%xx=ShowLabel("Note")
					   nn=prefix & "NoteContatto" & id
					 %>
					 <input type="text" class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("NoteContatto")%>" >
                  </div>		
			   </div>

			</div>
	 
		 <%
		    Rs.moveNext 
		 loop  
		 Rs.close
	  end if 
	  if conta=0 then 
	     session("EsitoCallSelIndirizzo")="KO"
	     response.write "<h3>Nessuna pec in archivio</h3> "
	  end if 
      
	  %>


