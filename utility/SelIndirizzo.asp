<!--#include virtual="/gscVirtual/include/includeStd.asp"-->

<%
session("EsitoCallSelIndirizzo") = "" 
IdAccount = session("params_IdAccount")
'lista dei campi da valorizzare tipo 
'  Campo=nome_delcampo;
ListaCampi= Session("params_ListaCampi")
ArCampi = split(ListaCampi,";")
c_DescStato = ""
c_Indirizzo = ""
c_Civico    = ""
c_Cap       = ""
c_Comune    = ""
c_Provincia = "" 
for J=lbound(ArCampi) to ubound(ArCampi)-1
    str=ArCampi(j)
	ArDett = split(str,":")
	if ubound(ArDett)=1 then 
	   nome=ucase(trim(ArDett(0)))
	   valo=trim(ArDett(1))
	   if Nome="DESCSTATO" then 
	      c_DescStato=valo
	   end if 
	   if Nome="INDIRIZZO" then 
	      c_Indirizzo=valo
	   end if 
	   if Nome="CIVICO" then 
	      c_Civico=valo
	   end if 
	   if Nome="CAP" then 
	      c_Cap=valo
	   end if 
	   if Nome="COMUNE" then 
	      c_Comune=valo
	   end if 
	   if Nome="PROVINCIA" then 
	      c_Provincia=valo
	   end if 	   
   end if 
next 


prefix = "params_"

%>

<script>

function confirmModalProcedi()
{
	
	var s=$('input[name="sel"]:checked').val();

	var DescStato = $('#params_DescStato' + s).val();
	var Indirizzo = $('#params_Indirizzo' + s).val();
	var Civico    = $('#params_Civico'    + s).val();
	var Cap       = $('#params_Cap'       + s).val();
	var Comune    = $('#params_Comune'    + s).val();
	var Provincia = $('#params_Provincia' + s).val();
	d=$('#c_DescStato').val();
	$('#'+d).val(DescStato);
	d=$('#c_Indirizzo').val();
	$('#'+d).val(Indirizzo);
	d=$('#c_Civico').val();
	$('#'+d).val(Civico);
	d=$('#c_Cap').val();
	$('#'+d).val(Cap);
	d=$('#c_Comune').val();
	$('#'+d).val(Comune);
	d=$('#c_Provincia').val();
	$('#'+d).val(Provincia);	
	$('#dismissModalGenerica').click();
}
</script>
      <input type="hidden" name="c_DescStato" id="c_DescStato" value="<%=c_DescStato%>">
	  <input type="hidden" name="c_Indirizzo" id="c_Indirizzo" value="<%=c_Indirizzo%>">
	  <input type="hidden" name="c_Civico"    id="c_Civico"    value="<%=c_Civico%>">
	  <input type="hidden" name="c_Cap"       id="c_Cap"       value="<%=c_Cap%>">
	  <input type="hidden" name="c_Comune"    id="c_Comune"    value="<%=c_Comune%>">
	  <input type="hidden" name="c_Provincia" id="c_Provincia" value="<%=c_Provincia%>">
      <%
	  prefix = "params_"
	  Set Rs = Server.CreateObject("ADODB.Recordset")
	  MyContQ = ""
	  MyContQ = MyContQ & " select * from AccountSede A "
	  MyContQ = MyContQ & " inner join TipoSede B on A.IdTipoSede = B.IdTipoSede "
	  MyContQ = MyContQ & " left  join Stato C on A.idStato = C.IdStato"
	  MyContQ = MyContQ & " Where A.IdAccount = " & IdAccount
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
			id = Rs("IdAccountSede")
		 %>
			<div class="row">
	
			   <div class="col-2">
                  <div class="form-group ">
				     <input name="sel" type="radio" id="radio<%=conta%>"  value="<%=id%>" <%=checked%>>
				     <%xx=ShowLabel("Stato")
					   nn=prefix & "DescStato" & id
					 %>
					 
					 <input type="text" class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("DescStato")%>" >
                  </div>		
			   </div>
			   <div class="col-3">
                  <div class="form-group ">
				     <%xx=ShowLabel("Comune")
					 nn=prefix & "Comune" & id
					 %>
					 <input type="text" class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("Comune")%>" >
                  </div>		
			   </div>			   
			   <div class="col-3">
                  <div class="form-group ">
				     <%xx=ShowLabel("Indirizzo")
					   nn=prefix & "Indirizzo" & id
					 %>
					 <input type="text" class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("Indirizzo")%>" >
                  </div>		
			   </div>
			   <div class="col-1">
                  <div class="form-group ">
				     <%xx=ShowLabel("Civico")
					 nn=prefix & "Civico" & id
					 %>
					 <input type="text" class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("Civico")%>" >
                  </div>		
			   </div>	
			   <div class="col-1">
                  <div class="form-group ">
				     <%xx=ShowLabel("Cap")
					 nn=prefix & "Cap" & id
					 %>
					 <input type="text" class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("Cap")%>" >
                  </div>		
			   </div>

			   <div class="col-2">
                  <div class="form-group ">
				     <%xx=ShowLabel("Provincia")
					 nn=prefix & "Provincia" & id
					 %>
					 <input type="text" class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=Rs("Provincia")%>" >
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
	     response.write "<h2>Nessun indirizzo in archivio</h2> "
	  end if 
      
	  %>


