<%
  NomePagina="DocumentiCliente.asp"
  titolo="Menu - Documenti Clienti"
  default_check_profile="Clie,Coll"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAffidamento.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionEvento.asp"-->
<%
  livelloPagina="00"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include virtual="/gscVirtual/include/head.asp"-->
<!-- Custom styles for this template -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">

</head>
<script>
function localIns()
{
	xx=ImpostaValoreDi("Oper","CALL_INS");
	document.Fdati.submit();
}
function localUpd(id)
{
	xx=ImpostaValoreDi("ItemToRemove",id);
	xx=ImpostaValoreDi("Oper","CALL_UPD");
	document.Fdati.submit();
}
</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%

IdAccount         = Session("LoginIdAccount")
IdAccountCliente  = 0
PaginaReturn = ""

if FirstLoad then 
   v_attivi="S"
   v_Cessati=""
   TipoRife     = getCurrentValueFor("TipoRife")
   IdRife       = getCurrentValueFor("IdRife")
   PaginaReturn = getCurrentValueFor("PaginaReturn")
else
   v_attivi     = Request("attivi")
   v_Cessati    = Request("cessati")
   TipoRife     = getValueOfDic(Pagedic,"TipoRife")
   IdRife       = getValueOfDic(Pagedic,"IdRife") 
   PaginaReturn = getValueOfDic(Pagedic,"PaginaReturn") 
end if 

xx=setCurrent(NomePagina,livelloPagina) 

qSelDitta=""
TipoRife=Trim(TipoRife)

if IsCliente() then 
   IdAccountCliente = IdAccount 
else 
   If TipoRife="COOB" then 
      qSelDitta = qSelDitta & " Select * "
      qSelDitta = qSelDitta & " from AccountCoobbligato"
      qSelDitta = qSelDitta & " where IdAccountCoobbligato = " & NumforDb(IdRife)
   end if 
   If TipoRife="ATI" then 
      qSelDitta = qSelDitta & " Select * "
      qSelDitta = qSelDitta & " from AccountATI"
      qSelDitta = qSelDitta & " where IdAccountATI = " & NumforDb(IdRife)
   end if 
   if qSelDitta<>"" then 
      IdAccountCliente   = LeggiCampo(qSelDitta,"IdAccount")
   end if 
end if 

if Oper="VALIDA" and TipoRife="COOB"  then 
   qq = "select * from AccountCoobbligato where IdAccountCoobbligato = " & NumforDb(IdRife)
   IdStatoValidazione = LeggiCampo(qq,"IdStatoValidazione")
   NomeCoob = LeggiCampo(qq,"RagSoc")
   qUpd = ""
   qUpd = qUpd & " update AccountCoobbligato "
   if IdStatoValidazione="" then
      tipoGestore = Getdiz(session("Login_Parametri"),"VAL_COB")
      qUpd = qUpd & " set IdStatoValidazione = 'RICH'"
	  qUpd = qUpd & " ,TipoGestore = '" & tipoGestore & "'"
	  qUpd = qUpd & " ,IdAccountRichiedente=" & Session("LoginIdAccount")
      qUpd = qUpd & " ,DataRichiesta = " & Dtos()
   else 
      qUpd = qUpd & " set IdStatoValidazione = 'LAVO'"
   end if 
   qUpd = qUpd & " where IdAccountCoobbligato = " & NumforDb(IdRife)
   ConnMsde.execute qUpd 
   descInfo="Validazione coobbligato " & NomeCoob
   XX=createEvento("COOB","VALI",Session("LoginIdAccount"),descInfo,"AccountCoobbligato","IdAccountCoobbligato=" & IdRife,true,0)
   response.redirect RitornaA(PaginaReturn)
   response.end   
end if 
if Oper="VALIDA" and TipoRife="ATI"  then 
   IdStatoValidazione = LeggiCampo("select * from AccountATI where IdAccountAti = " & NumforDb(IdRife),"IdStatoValidazione")
   qUpd = ""
   qUpd = qUpd & " update AccountATI "
   if IdStatoValidazione="" then 
      tipoGestore = Getdiz(session("Login_Parametri"),"VAL_ATI")
      qUpd = qUpd & " set IdStatoValidazione = 'RICH'"
	  qUpd = qUpd & " ,TipoGestore = '" & tipoGestore & "'"
	  qUpd = qUpd & " ,IdAccountRichiedente=" & Session("LoginIdAccount")
	  qUpd = qUpd & " ,DataRichiesta = " & Dtos()
   else 
      qUpd = qUpd & " set IdStatoValidazione = 'LAVO'"
   end if 
   
   qUpd = qUpd & " where IdAccountATI = " & NumforDb(IdRife)
   ConnMsde.execute qUpd   
   descInfo="Validazione azienda per A.T.I." 
   XX=createEvento("ATI","VALI",Session("LoginIdAccount"),descInfo,"AccountCoobbligato","IdAccountCoobbligato=" & IdRife,true,0)
   'response.write qUpd
   response.redirect RitornaA(PaginaReturn)
   response.end   
end if 


if Oper="CALL_INS" or Oper="CALL_UPD" then
   xx=RemoveSwap()
   Session("swap_IdTabella")          = "CLIENTE_DOC"
   Session("swap_IdTabellaKeyInt")    = IdAccountCliente
   Session("swap_OperTabella")        = Oper
   Session("swap_IdAccount")          = IdAccountCliente
   Session("swap_IdAccountDocumento") = Cdbl("0" & Request("ItemToRemove"))
   Session("swap_PaginaReturn")       = "configurazioni/Clienti/" & NomePagina
   Session("swap_OperAmmesse")        = "CRUD"
   Session("swap_TipoRife")           = TipoRife
   Session("swap_IdRife")             = IdRife 
   'response.write IdAccountCliente
   'response.end 
   response.redirect RitornaA("configurazioni/Clienti/DocumentoClienteUpload.asp")
   response.end 
end if  

if Oper="DEL" then 
    Session("TimeStamp")=TimePage
    KK  = Cdbl("0" & Request("ItemToRemove"))
    MsgErrore = funDoc_DelAccDoc(kk)
End if

xx=setValueOfDic(Pagedic,"PaginaReturn",PaginaReturn)
xx=setValueOfDic(Pagedic,"TipoRife"    ,TipoRife)
xx=setValueOfDic(Pagedic,"IdRife"      ,IdRife)
xx=setCurrent(NomePagina,livelloPagina) 

IdRife=cdbl("0" & IdRife) 


IdTipoDitta       = ""
IdStatoValidazione= ""
descDitta         = ""
TipoRife=Trim(TipoRife)

if qSelDitta<>"" then 
    IdAccountCliente   = LeggiCampo(qSelDitta,"IdAccount")
    IdTipoDitta        = LeggiCampo(qSelDitta,"IdTipoDitta")
	descDitta          = LeggiCampo(qSelDitta,"RagSoc")
	IdStatoValidazione = LeggiCampo(qSelDitta,"IdStatoValidazione")
	NoteValidazione    = LeggiCampo(qSelDitta,"NoteValidazione")
end if
if DescDitta<>"" then 
   If TipoRife="COOB" then 
      DescTipoDitta = "Coobbligato "
   end if 
   If TipoRife="ATI" then 
      DescTipoDitta = "A.T.I. "   
   end if 
   
end if
CanModify=false  
if IdStatoValidazione="" or IdStatoValidazione="RICH" or IdStatoValidazione="DOCU" then 
   CanModify=true 
end if 
%>

<div class="d-flex" id="wrapper">
	<%
	  
      callP=VirtualPath & "bar/" & Session("sideBar_" & Session("LoginIdAccount")) 
	  
	  'response.write callP
	  'response.end 
	  
      Server.Execute(callP) 
	  
	%>
    <!-- Page Content -->
	<div id="page-content-wrapper">
	<%
      callP=VirtualPath & "bar/" & Session("TopBar_" & Session("LoginIdAccount")) 
      Server.Execute(callP) 
	%>	

		<div class="container-fluid">
			<form name="Fdati" Action="<%=NomePagina%>" method="post">

			<div class="row">
				<%
				if PaginaReturn<>"" then 
				   RiferimentoA="col-1  text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"
				%>
				   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<% 
				else 
				   response.write "<div class='col-1'></div>"
				end if %>
				<div class="col-11"><h3>Cassetto Documentale</h3>
				</div>
			</div>
            <%if Iscliente()=false or DescDitta<>"" then %>
                <div class="row">
				   <%if Iscliente()=false then %>
                   <div class = "col-5">
                        <div class="form-group ">
                        <%xx=ShowLabel("Cliente")
						DescCliente=LeggiCampo("select * from Cliente Where IdAccount=" & IdAccountCliente,"Denominazione")
						%>
                        <input value="<%=DescCliente%>" readonly class="form-control"  >
                        </div>
                  </div>
                  <%end if %>				
  			      <%if DescDitta<>"" then %>
				   <div class = "col-5"> 
                        <div class="form-group ">
                        <%xx=ShowLabel(DescTipoDitta)%>
                        <input value="<%=DescDitta%>" readonly class="form-control"  >
                        </div>
				   </div>
				   <%end if %>
              </div>			
			
			<%end if %>			
			<%
			AddRow=true
			dim CampoDb(10)
			ElencoOption = ";0;Documento;1"
            CampoDB(1)   = "DescDocumento"
			
			%>
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
		<%
		check_Attivi=""
		if v_attivi <> "" then 
		   check_Attivi = "checked=""checked"""  
		end if
		check_cessati=""
		if v_cessati <> "" then 
		   check_Cessati = "checked=""checked"""  
		end if
		%>

	<div class="row no-row-margin " style="margin-top: 10px;margin-bottom: 10px;" >

      <div class="col-1 s1 no-margin font-weight-bold">
	     mostra
	  </div>
	  
      <div class="col-1 no-margin">
      <label>
	    <input type="checkbox" name="attivi"  id="attivi" <%=check_attivi%> value="on">
        <span class="font-weight-bold">Attivi</span>
      </label>
	  </div>	
      <div class="col-1 no-margin">
      <label>
	    <input type="checkbox" name="cessati"  id="cessati" <%=check_cessati%> value="on">
        <span class="font-weight-bold">Tutti</span>
      </label>
	  </div>	
 
	  
	</div>
	<%if IdStatoValidazione<>"" or NoteValidazione<>"" then 
	     if IdStatoValidazione<>"" then 
		    DescStatoValidazione=LeggiCampoTabellaText("StatoServizio",IdStatoValidazione)
		 end if 
	
	%>
		<div class="row">
		   <div class="col-3">
               <div class="form-group ">
		       <%xx=ShowLabel("Stato Validazione")%>
			   <input type="text" readonly class="form-control" value="<%=descStatoValidazione%>" >
             </div>		
		   </div>
		   <div class="col-3">
               <div class="form-group ">
		       <%xx=ShowLabel("Note Validazione")%>
			   <input type="text" readonly class="form-control" value="<%=NoteValidazione%>" >
             </div>		
		   </div>
		   
		</div>
			
	<%end if %>
			
<%
'caricamento tabella 
   if Condizione<>"" then 
      Condizione = " And " & Condizione
   end if 

   Oggi = Dtos()
   
   Set Rs = Server.CreateObject("ADODB.Recordset")
   MySql = "" 
   MySql = MySql & " Select A.*,B.DescDocumento,D.DescTipoValidazione"
   MySql = MySql & ",IsNull(C.DescBreve,'') as DescBreve,isnull(C.NomeDocumento,'') as NomeDocumento"
   MySql = MySql & ",IsNull(C.PathDocumento,'') as PathDocumento"
   MySql = MySql & ",IsNull(C.ValidoDal,0) as ValidoDal"   
   MySql = MySql & ",IsNull(C.ValidoAl,99991231) as ValidoAl"   
   MySql = MySql & " From AccountDocumento A  "
   MySql = MySql & " inner join Documento B on a.idDocumento = B.IdDocumento"
   MySql = MySql & " left  join Upload C on a.idUpload = C.IdUpload"
   MySql = MySql & " left  join TipoValidazione D on a.idTipoValidazione = D.idTipoValidazione"
   MySql = MySql & " Where a.IdAccount = " & IdAccountCliente
   MySql = MySql & " and A.TipoRife = '" & TipoRife & "'"
   MySql = MySql & " and A.IdRife = " & NumForDb(IdRife)
   if (v_attivi <>"" and v_Cessati = "") or (v_attivi = "" and v_Cessati <> "") then
      if v_attivi <>"" then
         MySql = MySql & " and IsNull(C.ValidoDal,0) <= " & oggi & " and IsNull(C.ValidoAl,99991231) > " & Oggi
      else
         MySql = MySql & " and IsNull(C.ValidoAl,99991231) <= " & Oggi
      end if 
   end if 
   MySql = MySql & Condizione & " order By DescDocumento"
   'response.write MySql 
   'controllo se non ci sono documenti 
   if FirstLoad and (TipoRife="COOB" or TipoRife="ATI") then 
      Vuoto=LeggiCampo(MySql,"DescDocumento")
      if Vuoto="" then 
		 if IdTipoDitta="" then 
		    IdTipoDitta="xxxx"
		 end if 
         MyQ = "" 
         MyQ = MyQ & " Insert into AccountDocumento ("
         MyQ = MyQ & " IdAccount,IdDocumento,TipoRife,IdRife,FlagObbligatorio)"
         MyQ = MyQ & " select " & idAccountCliente & " as IdAccount,IdDocumento"
         MyQ = MyQ & ",IdTipoUtenza," & idRife & " as IdRife,FlagObbligatorio"   
         MyQ = MyQ & " from ServizioDocumento "
		 MyQ = MyQ & " Where IdAnagServizio='CAUZ_PROV' "
		 MyQ = MyQ & " and IdTipoUtenza='" & TipoRife & "'"
		 MyQ = MyQ & " and (DITT='" & IdTipoDitta & "' or PEGI='" & IdTipoDitta & "' or PEFI='" & IdTipoDitta & "')"
		 'response.write MyQ
         ConnMsde.execute MyQ
      end if 
   end if 
   'response.write MySql 
   Rs.CursorLocation = 3 
   Rs.Open MySql, ConnMsde

   DescLoaded=""
   NumCols = numC + 1
   NumRec  = 0
   ShowNew    = true
   ShowUpdate = false
   MsgNoData  = ""
   
   ShowRichiesto = false 
   if TipoRife="COOB" or TipoRife="ATI" then 
      ShowRichiesto = true 
   end if 
%>

<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

		<div class="table-responsive"><table class="table"><tbody>
		<thead>
		<tr>
			<th scope="col"> Documento
			<%
			if IdStatoValidazione="" or IdStatoValidazione="DOCU" then 
               RiferimentoA="col-2;#;;2;inse;Inserisci;;localIns();N"
			
			%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
			<%end if %>
		    </th>
			<%if ShowRichiesto = true then %>
			<th scope="col" width="10%">Richiesto</th>
			<%end if %>
		    <th scope="col" width="15%">Valido Dal</th>
		    <th scope="col" width="15%">Valido Al</th>
		    <th scope="col">Validazione</th>
		    <th scope="col">Azioni</th>
		</tr>
		</thead>

<%
        TuttiCaricati=true 
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
			Primo=0
			Do While Not rs.EOF and (NumRec<PageSize or Pagesize<=0)
				Primo=Primo+1
				NumRec=NumRec+1
				Id=Rs("IdAccountDocumento")
				PathDocumento=trim(RS("PathDocumento"))
				if PathDocumento="" and Rs("FlagObbligatorio")=1 then 
				   TuttiCaricati=false
				end if 
				DescLoaded=DescLoaded & Id & ";"
		%>
			<tr scope="col">
				<td>
				    <%
					DescDocumento=Rs("DescBreve")
					if DescDocumento="" then 
					   DescDocumento=Rs("DescDocumento")
					end if 
					TipoRife=Rs("TipoRife")
					if TipoRife<>"" then 
					   IdRife=Rs("IdRife") 
					   if Cdbl(IdRife)>0 then 
					      qs = ""
						  if ucase(TipoRife)="COOB" then 
						     qs = "Select Ragsoc  as descInfo from AccountCoobbligato Where IdAccountCoobbligato=" & IdRife 
						  end if 
						  if ucase(TipoRife)="ATI" then 
						     qs = "Select Ragsoc  as descInfo from AccountATI Where IdAccountATI=" & IdRife 
						  end if 						  
						  if qs<>"" then 
						     'response.write qs
						     addInfo=LeggiCampo(qs,"descInfo")
							 if instr(ucase(DescDocumento),ucase(addInfo))=0 then 
							    DescDocumento = DescDocumento & " " & addInfo 
							 end if 
						  end if 
					   end if 
					end if 
					
					%>
					<input class="form-control" type="text" readonly value="<%=DescDocumento%>">
				</td>
				<%if ShowRichiesto = true then %>
				<td><%
                      if Rs("FlagObbligatorio")=0 then 
					     SiNo = "NO"
					  else
					     SiNo = "SI"
					  end if 
                    %>
					<input class="form-control" type="text" readonly value="<%=SiNo%>">
                </td>
				<%end if %>
				<td>
				    <%
					
					ValidoDal=Rs("ValidoDal")
					if ValidoDal="0" then 
					   ValidoDal=""
					else 
					   ValidoDal=StoD(ValidoDal)
					end if 
					%>
					<input class="form-control" type="text" readonly value="<%=ValidoDal%>">
				</td>
				<td>
				    <%
					ValidoAl=Rs("ValidoAl")
					if ValidoAl="99991231" then 
					   ValidoAl=""
					else 
					   ValidoAl=StoD(ValidoAl)					   
					end if 
					%>
					<input class="form-control" type="text" readonly value="<%=ValidoAl%>">
				</td>				

				<td>
				   <%
				       if PathDocumento="" then 
					      Validazione="Da caricare"
					   else 
					      Validazione=funDoc_DescrizioneStatoDoc(Rs("IdTipoValidazione"))
					   end if 
				   %>
					<input class="form-control" type="text" readonly value="<%=Validazione%>">
				</td>
		
			<td>
			    <% if CanModify=true then %>
					<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;localUpd('" & id & "');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->		
					<%
					if instr(OperAmmesse,"D")>=0 then 
						RiferimentoA="col-2;#;;2;dele;Cancella;;SalvaSingoloEdAttiva('DEL'," & Id & ",true,'','','');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					<%end if %>
				<%end if %>
				
				<%
				Linkdocumento=Rs("PathDocumento")%>
				<!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->	 				
			</td>
			</tr>
		<%	
		    rs.MoveNext
	    Loop
   else
      TuttiCaricati=false 
   end if 
   rs.close


%>

</tbody></table></div> <!-- table responsive fluid -->
   <%
   if TuttiCaricati=true and ShowRichiesto = true then 
      if IdStatoValidazione="" or instr("DOCU",IdStatoValidazione)>0 then %>
        <div class="row">
          <div class="col-12 text-center">
              <button type="button" onclick="AttivaFunzione('VALIDA',0)" class="btn btn-success">Richiedi Validazione</button>
          </div>
        </div>
	  <%
      end if 
   end if 
   %>
			<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
			<!--#include virtual="/gscVirtual/include/paginazione.asp"-->

			</form>
		</div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->

<!--#include virtual="/gscVirtual/include/scripts.asp"-->

  <!-- Menu Toggle Script -->
  <script>
    $("#menu-toggle").click(function(e) {
      e.preventDefault();
      $("#wrapper").toggleClass("toggled");
    });
  </script>
  <script>
    $(document).ready(function(){
      $('[data-toggle="tooltip"]').tooltip();   
    });
  </script>
  <script>
$('.btn').onClick(function(e){
  e.preventDefault();
});  
</script>
</body>

</html>
