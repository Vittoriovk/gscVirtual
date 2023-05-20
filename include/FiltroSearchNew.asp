

<%if AddRow then %>

<%end if%>
<%
if funSearchEsegui="" then 
   funSearchEsegui="Sottometti();"
end if 

check_inizia=""
if v_inizia_per <> "" then 
   check_inizia = "checked=""checked"""  
end if 


%> 
		
		<div class="col-3 no-margin">
			<% 
			if Len(ElencoOption)>0 then 
			response.write OptionListaValori("TipoRicerca",ElencoOption,v_tipoRicerca)
			end if
			%>
		</div>
	

		<div class="d-flex">
			<input type="text" name="cerca_testo" id="cerca_testo"  class="form-control form-control-lg" style="font-size: 1.25rem;" value="<%=v_cercatesto%>" placeholder="Cerca">
			<div class="btn-group-vertical">
				<a title="Cerca" href="#" id="FiltroSearchNewButton" onclick="<%=funSearchEsegui%>">
					<i class="fa fa-search" style="font-size: 1.5rem; margin-left: 1rem;"></i>
				</a>  
			</div> 
		</div>
		<% if AddRow then %>

<% end if %>

<%
'valutazione della ricerca 
TipoRicerca=Request("TipoRicerca")
cerca_testo = Request.form("cerca_testo")
if cerca_testo = "" then
	cerca_testo = Request.querystring("cerca_testo")
end if
inizia_per = Request.form("inizia_per")
if inizia_per = "" then
	inizia_per = Request.querystring("inizia_per")
end if

If Inizia_per = "" then 
   PrimoC="%"
else
   PrimoC=""
end if 

CampoR=""
if isnumeric(TipoRicerca) then 
   CampoR=CampoDB(TipoRicerca)
end if 

Condizione = ""

if cerca_testo <> "" and CampoR<>"" then
	Condizione = CampoR & " like '" & PrimoC & apici(cerca_testo) & "%' " 
end if

%>


<script>
	var selectElement = document.querySelector('select[name="TipoRicerca"]');
if (selectElement && selectElement.childElementCount === 0) {
  selectElement.parentNode.removeChild(selectElement);
}
</script>



