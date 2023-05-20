<tbody>

<%
Set Rs = Server.CreateObject("ADODB.Recordset")

if Condizione<>"" then 
	Condizione=" and " & Condizione
end if 

MySql = "" 
MySql = MySql & " Select a.* From TrattamentoFiscaleStorico A "
MySql = MySql & " Left join Regione B on A.IdRegione = B.IdRegione"
MySql = MySql & " Left join Provincia c on A.IdProvincia = c.IdProvincia"
MySql = MySql & " Where IdTrattamentoFiscale=" & IdTrattamentoFiscale
MySql = MySql & Condizione & " order By DataInizio"

'response.write mysql 

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

DescLoaded=""
NumCols = 4
NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
%>
	<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->


<div class="table-responsive">
	<table class="table">
		<thead>
			<tr>
			  <th scope="col">Valido Dal</th>
			  <th scope="col">Per La Regione</th>
			  <th scope="col">Per La Provincia</th>
			  <th scope="col">Perc. %</th>
			  <th scope="col">Azioni</th>
			</tr>
		</thead>

<%
if MsgNoData="" then 
	if PageSize>0 then 
		Rs.PageSize = PageSize
		pageTotali = rs.PageCount
		NumRec=0
		if Cpag<=0 then 
			Cpag =1
		end if 
		if Cpag>PageTotali then 
			CPag=PageTotali
		end if  
		Rs.absolutepage=CPag
	end if
	NumRec=0
	Do While Not rs.EOF and (NumRec<PageSize or Pagesize<=0)
		Primo=Primo+1
		NumRec=NumRec+1
		Id=Rs("IdTrattamentoFiscaleStorico")
		DescLoaded=DescLoaded & Id & ";"
		IdRef="DataInizio" & Id 
		%> 
	<tr scope="col"> 
		<td>
			<input class="form-control mydatepicker" type="text" name="<%=IdRef%>" id="<%=IdRef%>" placeholder="<%=pHold%>" 
			maxlength="90" value="<%=StoD(Rs("DataInizio"))%>">
		</td>
		<td>
		   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
		   <%
		   ao_row = ""
		   ao_lbd = ""             'descrizione label 
           ao_3ld = ""             'descrizione terzo elemento
		   ao_div = ""   
		   ao_nid = "IdRegione" & Id            'nome ed id
		   ao_val = Rs("IdRegione") 'valore di default	
		   ao_Tex = "select * from Regione order By DescRegione"
		   ao_ids = "IdRegione"			  'valore della select 
		   ao_des = "DescRegione"           'valore del testo da mostrare 
		   ao_cla = ""                        'azzero classe
		   ao_Eve = ""                        'azzero evento
		   ao_Att = "1"                       'indica se deve mettere vuoto 
		   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
		   ao_Cla = "class='form-control form-control-sm'"
		   %>
		   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->		
		
		</td>
		<td>
		   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
		   <%
		   ao_row = ""
		   ao_lbd = ""             'descrizione label 
           ao_3ld = ""             'descrizione terzo elemento
		   ao_div = ""   
		   ao_nid = "IdProvincia" & Id           'nome ed id
		   ao_val = Rs("IdProvincia")	
		   ao_Tex = "select * from Provincia order By DescProvincia"
		   ao_ids = "IdProvincia"			  'valore della select 
		   ao_des = "DescProvincia"           'valore del testo da mostrare 
		   ao_cla = ""                        'azzero classe
		   ao_Eve = ""                        'azzero evento
		   ao_Att = "1"                       'indica se deve mettere vuoto 
		   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
		   ao_Cla = "class='form-control form-control-sm'"
		   %>
		   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->		
		
		</td>		
		<td>
			<%Idref="PercImposta" & Id %>
			<input class="form-control" type="text" name="<%=IdRef%>" id="<%=IdRef%>" placeholder="<%=pHold%>" 
			maxlength="90" value="<%=Rs("PercImposta")%>">
		
		</td>
		<td>
		
			<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;SalvaSingoloEdAttiva('UPD'," & Id & ",true,'','','');N"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
			<%RiferimentoA="col-2;#;;2;dele;Cancella;;SalvaSingoloEdAttiva('DEL'," & Id & ",true,'','','');N"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
		</td>
	</tr> 
	<%
		rs.MoveNext
	Loop
end if 
rs.close
%>
<%if ShowNew then 
	Id=0
%>
	<tr> 
		<td>
			<% 	IdRef="DataInizio" & Id 
				pHold="Inserire Data"
			%>	
			<input type="text" class="mydatepicker form-control" id="<%=IdRef%>" name="<%=IdRef%>" 
					placeholder="gg/mm/aaaa" title="format : gg/mm/aaaa" value="<%=StoD(Dtos())%>"/>
		</td>

		<td>
		   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
		   <%
		   ao_row = ""
		   ao_lbd = ""             'descrizione label 
           ao_3ld = ""             'descrizione terzo elemento
		   ao_div = ""   
		   ao_nid = "IdRegione0"            'nome ed id
		   ao_val = "" 'valore di default	
		   ao_Tex = "select * from Regione order By DescRegione"
		   ao_ids = "IdRegione"			  'valore della select 
		   ao_des = "DescRegione"           'valore del testo da mostrare 
		   ao_cla = ""                        'azzero classe
		   ao_Eve = ""                        'azzero evento
		   ao_Att = "1"                       'indica se deve mettere vuoto 
		   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
		   ao_Cla = "class='form-control form-control-sm'"
		   %>
		   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->		
		
		</td>		
		<td>
		   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
		   <%
		   ao_row = ""
		   ao_lbd = ""             'descrizione label 
           ao_3ld = ""             'descrizione terzo elemento
		   ao_div = ""   
		   ao_nid = "IdProvincia0"            'nome ed id
		   ao_val = "" 'valore di default	
		   ao_Tex = "select * from Provincia order By DescProvincia"
		   ao_ids = "IdProvincia"			  'valore della select 
		   ao_des = "DescProvincia"           'valore del testo da mostrare 
		   ao_cla = ""                        'azzero classe
		   ao_Eve = ""                        'azzero evento
		   ao_Att = "1"                       'indica se deve mettere vuoto 
		   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
		   ao_Cla = "class='form-control form-control-sm'"
		   %>
		   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->		
		
		</td>		
		<td>
			<%Idref="PercImposta" & Id %>
			<input class="form-control" type="text" name="<%=IdRef%>" id="<%=IdRef%>" placeholder="<%=pHold%>" 
			maxlength="10" value="0">
		
		</td>		
		<td align="left">
			<%RiferimentoA="col-2;#;;2;insert;Inserisci;;SaveWithOper('INS')"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
		</td>
	</tr>

	   
<%end if%>
</tbody></table></div> <!-- table responsive fluid -->
