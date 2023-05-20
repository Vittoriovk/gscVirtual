    <%if cercaComp<>"N" then %>
	<div class="row">

      <div class="col-1 s1 no-margin font-weight-bold">
	     Compagnia
	  </div>
	  
      <div class="col-9 no-margin">
	  <%
	    stdClass="class='form-control form-control-sm'"
	    q="Select * from Compagnia "
		q=Q & " order by DescCompagnia "
		'Where 
	    response.write ListaDbChangeCompleta(q,"IdCompagnia0",IdCompagnia ,"IdCompagnia","DescCompagnia" ,1,"","","","","",stdClass)
	  
	  %>
	  </div>	
      <div class="col-2 no-margin">
	  </div>	

	</div>
	<%end if %>
	
	<div class="row">

      <div class="col-1 s1 no-margin font-weight-bold">
	     Collaboratore
	  </div>
	  
      <div class="col-9 no-margin">
	  <%
	    stdClass="class='form-control form-control-sm'"
	    
		q="Select * from Collaboratore Where 1=1 "
		if IsBackOffice() then 
	       q = q & " and IdAccountLivello1=0 "
		else
		   q = q & "and " & trim(getCondForLevel(session("LivelloAccount"),Session("LoginIdAccount")))
		end if 
		q=Q & " order by Denominazione "
		'Where 
	    response.write ListaDbChangeCompleta(q,"IdCollaboratore0",IdAccountCollaboratore ,"IdAccount","Denominazione" ,1,"","","","","",stdClass)
	  
	  %>
	  </div>	
      <div class="col-2 no-margin">
	  </div>	

	</div>



	<%
	AddRow=true
	dim CampoDb(10)
	ElencoOption = ";0;Denominazione;1;Codice Fiscale;2;Partita Iva;3"
    CampoDB(1)   = "Denominazione"
	CampoDB(2)   = "CodiceFiscale"
    CampoDB(3)   = "PartitaIva"

	%>
	<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->