<%
  NomePagina="ValidazioneBackODettaglio.asp"
  titolo="Menu - Dettaglio Validazione"
  default_check_profile="COLL,BackO"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
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
function ValidaRichiesta(stato)
{
   var oldN=ValoreDi("NameLoaded");
   var oldD=ValoreDi("DescLoaded");
   yy=ImpostaValoreDi("NameLoaded","ValidoDal,DTO;ValidoAl,DTO");
   yy=ImpostaValoreDi("DescLoaded",'0');
   xx=ElaboraControlli();
   yy=ImpostaValoreDi("NameLoaded",oldN);
   yy=ImpostaValoreDi("DescLoaded",oldD);
   if (xx==false)
      return false;
   yy=ImpostaValoreDi("DescLoaded",'0');
   yy=ImpostaValoreDi("NameRangeDT","ValidoDal;ValidoAl");
   xx=CheckRangeDT(0);
   yy=ImpostaValoreDi("DescLoaded",oldD);
   if (xx==false)
      return false;
   
   ImpostaValoreDi("ItemToModify",'');
   xx=ImpostaValoreDi("Oper","CALL_RIC");
   xx=ImpostaValoreDi("IdParm",stato);
   document.Fdati.submit();   
}



function localRicN(stato)
{
   var oldN=ValoreDi("NameLoaded");
   var oldD=ValoreDi("DescLoaded");
   yy=ImpostaValoreDi("NameLoaded","NoteValidazione,TE");
   yy=ImpostaValoreDi("DescLoaded",'0');
   xx=ElaboraControlli();
   yy=ImpostaValoreDi("NameLoaded",oldN);
   yy=ImpostaValoreDi("DescLoaded",oldD);
   if (xx==false)
      return false;

   xx=ImpostaValoreDi("Oper","CALL_RIC");
   xx=ImpostaValoreDi("IdParm",stato);
   document.Fdati.submit(); 

}

function registraStato(stato)
  {
	ImpostaValoreDi("ItemToModify",'');
    xx=ImpostaValoreDi("Oper","CALL_RIC");
	xx=ImpostaValoreDi("IdParm",stato);
    document.Fdati.submit();  
}

function localValN(id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    xx=$('#myConfirmInvalid').modal('toggle');
}

function myConfirmIvalidYes()
  {
	var inf = $("#myConfirmIvalidinfo").val();
	ImpostaValoreDi("ItemToModify",inf);
    xx=ImpostaValoreDi("Oper","CALL_VAL");
	xx=ImpostaValoreDi("IdParm","N");
    document.Fdati.submit();  

}

function localVal(stato,id)
{
    xx=ImpostaValoreDi("Oper","CALL_VAL");
	xx=ImpostaValoreDi("ItemToRemove",id);
	xx=ImpostaValoreDi("IdParm",stato);
    document.Fdati.submit();  
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%
   PaginaReturn = ""
   if FirstLoad then 
      PaginaReturn     = getCurrentValueFor("PaginaReturn")
	  TipoRife         = getCurrentValueFor("TipoRife")
	  IdRife           = getCurrentValueFor("IdRife")
   else
      PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
      TipoRife         = getValueOfDic(Pagedic,"TipoRife")
      IdRife           = getValueOfDic(Pagedic,"IdRife")
	  DescCliente      = getValueOfDic(Pagedic,"DescCliente")
   end if 
   
   tabella=""
   descAzione=""
   if TipoRife="COOB" then 
      tabella="AccountCoobbligato"
	  descAzione="Coobbligato"
   elseif TipoRife="ATI" then 
      tabella="AccountATI"
	  descAzione="A.T.I"
   end if 
   if tabella="" then 
      response.redirect RitornaA(paginaReturn)
   end if 

   'evita il refresh 
   if CheckTimePageLoad()=false then 
      Oper=""
   end if 
  'registrazione dei dati :
   IdSt = ucase(trim(Request("IdParm")))
   IdUp = cdbl("0" & Request("ItemToRemove"))
   if Oper = "CALL_VAL" and idSt="O" and idUp>0 then 
      MySql = "" 
      MySql = MySql & " update AccountDocumento set "
      MySql = MySql & " FlagObbligatorio=1"
      MySql = MySql & " Where IdAccountDocumento=" & IdUp
      ConnMsde.execute MySql 
   end if    
   if Oper = "CALL_VAL" and idSt="X" and idUp>0 then 
      MySql = "" 
      MySql = MySql & " update AccountDocumento set "
      MySql = MySql & " FlagObbligatorio=0"
      MySql = MySql & " Where IdAccountDocumento=" & IdUp
      ConnMsde.execute MySql 
   end if    
   
   if Oper = "CALL_VAL" and idSt="S" and idUp>0 then 
      MySql = "" 
      MySql = MySql & " update AccountDocumento set "
      MySql = MySql & " IdTipoValidazione='VALIDO'"
      MySql = MySql & ",NoteValidazione=''"
      MySql = MySql & " Where IdAccountDocumento=" & IdUp
      ConnMsde.execute MySql 
	  'response.write MySql
   end if 
   if Oper = "CALL_VAL" and idSt="N" and idUp>0 then 
      note=trim(Request("ItemToModify"))
      MySql = "" 
      MySql = MySql & " update AccountDocumento set "
      MySql = MySql & " IdTipoValidazione='NONVAL'"
      MySql = MySql & ",NoteValidazione='" & apici(note) & "'"
      MySql = MySql & " Where IdAccountDocumento=" & IdUp
      ConnMsde.execute MySql 
	  'response.write MySql
   end if    

   if Oper = "CALL_RIC" then 
      note =trim(Request("NoteValidazione0"))
	  qUpd = ""
	  qUpd = qUpd & " update " & Tabella & " set "
	  qUpd = qUpd & " IdStatoValidazione='" & IdSt & "'" 
	  qUpd = qUpd & ",NoteValidazione='" & apici(note) & "'" 
	  qUpd = qUpd & ",DataUltimaModifica=" & Dtos()
	  
      if IdSt="AFFI" then 
	     ValidoDal=DataStringa(Request("ValidoDal0"))
		 ValidoAl =DataStringa(Request("ValidoAl0"))
	     qUpd = qUpd & ",ValidoDal=" & NumForDb(ValidoDal)
	     qUpd = qUpd & ",ValidoAl="  & NumForDb(ValidoAL) 
		 qUpd = qUpd & ",FlagValidato=1"
      end if 
	  qUpd = qUpd & " Where  Id" & Tabella & "=" & IdRife 
      ConnMsde.execute qUpd 
	  qSel = "select * from " & Tabella & " Where  Id" & Tabella & "=" & IdRife 
	  DescRiferimento = LeggiCampo(,"Ragsoc")
      descInfo="Richiesta di Validazione per " & descAzione & ":" & DescRiferimento
      if IdSt="AFFI" then 
         descInfo = descInfo & " : conclusa"
      end if 
      if IdSt="ANNU" then 
         descInfo = descInfo & " : annullata"
      end if 
      if IdSt="DOCU" then 
         descInfo = descInfo & " : Integrazione documentazione"
      end if 
      if IdSt="RIFI" then 
         descInfo = descInfo & " : Rifiutata"
      end if 
   
      XX=createEvento(TipoRife,IdSt,Session("LoginIdAccount"),descInfo,"AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,0)
      'response.redirect RitornaA(PaginaReturn)
   end if 
 if Oper = "CALL_VALxxx" then 
    IdSt = trim(Request("IdParm"))
	'validato
	if idSt="X" then 
       DeSt = Request("NewDescStato0")
       ClSt = Request("NewDescStatoClie0")
       qUpd = ""
       qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
       qUpd = qUpd & " IdStatoAffidamento = '" & apici(IdSt) & "'"
       qUpd = qUpd & ",NoteAffidamento = '"    & apici(deSt) & "'"
       qUpd = qUpd & ",NoteAffidamentoCliente = '"    & apici(clSt) & "'"
       qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
	   'response.write qUpd 
       ConnMsde.execute qUpd 

	   
   end if    
 
 end if 
 

   xx=setValueOfDic(Pagedic,"PaginaReturn"     ,PaginaReturn)
   xx=setValueOfDic(Pagedic,"IdRife"           ,IdRife)
   xx=setValueOfDic(Pagedic,"TipoRife"         ,TipoRife)
   xx=setCurrent(NomePagina,livelloPagina) 

   %>
   <!-- ricarico la pagina senza fare logica  -->
   <!--#include virtual="/gscVirtual/include/ReloadPage.asp"-->
   <%

   
   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
   
   MySql = "" 
   MySql = MySql & " select a.*,B.Denominazione,C.DescStatoServizio,C.FlagStatoFinale" 
   MySql = MySql & " from " & Tabella & " A, Cliente B, StatoServizio C "
   MySql = MySql & " Where A.IdStatoValidazione = C.IdStatoServizio "
   MySql = MySql & " and   A.IdStatoValidazione <>''"
   MySql = MySql & " and   A.IdAccount = B.IdAccount" 
   MySql = MySql & " and   A.Id" & Tabella & "=" & IdRife 
   'response.write MySql 
   xx=GetInfoRecordset(DizDatabase,MySql)
   IdTipoPers=Getdiz(DizDatabase,"IdTipoDitta")
   OperAmmesse="R"
   IdStatoValidazione = Getdiz(DizDatabase,"IdStatoValidazione")
   FlagStatoFinale    =cdbl("0" & Getdiz(DizDatabase,"FlagStatoFinale"))
   DescLoaded=""
   MsgNoData  = ""
  
%>

<div class="d-flex" id="wrapper">
    <%
      callP=VirtualPath & "bar/" & Session("sideBar_" & Session("LoginIdAccount")) 
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
                   RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"
                %>
                   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <% end if %>
                <div class="col-11"><h3>Validazione <%=descAzione%></h3>
                </div>
            </div>
			
            <div class="row">
               <div class="col-3">
                  <div class="form-group ">
                     <%xx=ShowLabel("Utente")%>
                     <input type="text" readonly class="form-control" value="<%=GetDiz(DizDatabase,"Denominazione")%>" >
                  </div>        
               </div>
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Richiesta Del")%>
                     <input type="text" readonly class="form-control" value="<%=Stod(GetDiz(DizDatabase,"DataRichiesta"))%>" >
                  </div>        
               </div>
               <div class="col-3">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Stato Richiesta")
                     %>
                     <input type="text" readonly class="form-control" value="<%=GetDiz(DizDatabase,"DescStatoServizio")%>" >
                  </div>        
               </div>
               <div class="col-2">
                  <div class="form-group">
                     <%xx=ShowLabel("Elaborata il")%>
                     <input type="text" readonly class="form-control" value="<%=Stod(GetDiz(DizDatabase,"DataUltimaModifica"))%>" >
                  </div>                       
               </div> 			   
            </div>
            
			<%if FlagStatoFinale = 1 then %>
            <div class="row">
               <div class="col-8">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Annotazioni")
                     DescNote = NoteAffidamentoComp
                     if isCliente() then 
                        DescNote = noteAffidamentoClie
                     end if 
                     
                     %>
                  <input type="textArea" readonly class="form-control" value="<%=GetDiz(DizDatabase,"NoteValidazione")%>" >      
                  </div>               
               </div> 
			   <% if IdStatoValidazione="AFFI" then %>
               <div class="col-1">
                     <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Valido Dal")
                     valo=Stod(Getdiz(DizDatabase,"ValidoDal"))
                     %>
                     <input type="text" readonly class="form-control" value="<%=valo%>" >      
                     </div>	  
               </div> 			   
               <div class="col-1">
                     <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Valido Al")
                     valo=Stod(Getdiz(DizDatabase,"ValidoAl"))
                     %>
                     <input type="text" readonly class="form-control" value="<%=valo%>" >      
                     </div>	  
               </div> 			   
			   
			   <% end if %>
            </div> 
			<%end if %>
			
   <%readonly=" readonly "%>
   <div class="row">
      <div class="col-2"> 
           <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Tipo Utenza")
                q = "select * from TipoDitta where IdTipoditta='" & apici(IdTipoPers) & "'"
                valo = LeggiCampo(q,"DescTipoDitta")                     
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>        	  
      </div> 
      <%if IdTipoPers="PEGI" or IdTipoPers="DITT" then%>
	  <div class="col-2">
           <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Denominazione")
                valo=Getdiz(DizDatabase,"RagSoc")                     
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>  
       </div>
     <%end if%> 
     <%if IdTipoPers="DITT" or  IdTipoPers="PEFI" then %>
      <div class="col-2">
           <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Cognome")
                valo=Getdiz(DizDatabase,"Cognome")                     
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>  
      </div> 
     <div class="col-2">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Nome")
                valo=Getdiz(DizDatabase,"Nome")                     
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>  
      </div> 
   <%end if %>
      <div class="col-2">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Codice Fiscale")
                valo=Getdiz(DizDatabase,"CF")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  
      </div>  
	  <%if IdTipoPers="DITT" or IdTipoPers="PEGI" then%>
      <div class="col-2">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Partita Iva")
                valo=Getdiz(DizDatabase,"PI")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  	  
         <p class="font-weight-bold"></p>
      </div> 
      <% end if  %>
  </div>
  
   <%if IdTipoPers="DITT" or IdTipoPers="PEFI" then%>
   <div class="row">
      <div class="col-2">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Stato Nascita")
                valo=Getdiz(DizDatabase,"StatoNascita")
				q = "Select * from Stato Where IdStato='" & valo & "'" 
				Valo=LeggiCampo(q,"DescStato")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  
      </div>
      <div class="col-2">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Provincia Nascita")
                valo=Getdiz(DizDatabase,"ProvinciaNascita")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  
      </div> 

      <div class="col-2">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Comune Nascita")
                valo=Getdiz(DizDatabase,"ComuneNascita")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  
      </div> 
      <div class="col-1">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Data Nascita")
                valo=Stod(Getdiz(DizDatabase,"DataNascita"))
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  
      </div> 
      <div class="col-1">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Sesso")
                valo=Getdiz(DizDatabase,"IdSesso")
				q = "Select * from Sesso Where IdSesso='" & valo & "'" 
				Valo = LeggiCampo(q,"DescSesso")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  
 
      </div>
   </div>
    
   <%end if %>
   
  
   <div class="row">
      <div class="col-2">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Stato")
                valo=Getdiz(DizDatabase,"IdStato")
				q = "Select * from Stato Where IdStato='" & valo & "'" 
				Valo=LeggiCampo(q,"DescStato")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  
      </div>
      <div class="col-2">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Provincia")
                valo=Getdiz(DizDatabase,"Provincia")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  
      </div> 

      <div class="col-2">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Comune")
                valo=Getdiz(DizDatabase,"Comune")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  
      </div> 	  
      <div class="col-2">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("CAP")
                valo=Getdiz(DizDatabase,"CAP")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	  
      </div> 	  	  
   </div>
   
   <div class="row">
     <div class="col-6">
          <div class="form-group font-weight-bold">
              <%xx=ShowLabel("Indirizzo")
                valo=Getdiz(DizDatabase,"Indirizzo")
              %>
              <input type="textArea" readonly class="form-control" value="<%=valo%>" >      
           </div>	     
   
      </div>    
   </div>  
 
   <%
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
   MySql = MySql & " Where a.IdAccount = " & GetDiz(DizDatabase,"IdAccount")
   MySql = MySql & " and A.TipoRife = '" & TipoRife & "'"
   MySql = MySql & " and A.IdRife = " & NumForDb(IdRife)  
   Rs.CursorLocation = 3 
   Rs.Open MySql, ConnMsde
   'response.write MySql 
   
   DescLoaded=""
   NumCols = numC + 1
   NumRec  = 0
   ShowNew    = true
   ShowUpdate = false
   MsgNoData  = ""  
   Pagesize=0   
   'response.write MySql 
   %>
  
		<div class="table-responsive"><table class="table"><tbody>
		<thead>
		<tr>
			<th scope="col"> Documento</th>
			<th scope="col" width="10%">Richiesto</th>
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
				if PathDocumento="" then 
				   TuttiCaricati=false
				end if 
				DescLoaded=DescLoaded & Id & ";"
				idTipoValidazione = Rs("idTipoValidazione")
				'decido il colore del rigo 
				stileTr=""
				if IdTipoValidazione="VALIDO" then 
				   stileTr="style='background-color:#99FF99'"
				elseif IdTipoValidazione="NONVAL" then 
				   response.write "qui 1"
				   stileTr="style='background-color:#FFCC99'"
				elseif Rs("FlagObbligatorio")=1 and PathDocumento="" then 
				   response.write "qui 2"
				   stileTr="style='background-color:#FFCC99'"
				end if 				
		%>
			<tr scope="col">
				<td  <%=stileTr%>>
				    <%
					DescDocumento=Rs("DescBreve")
					if DescDocumento="" then 
					   DescDocumento=Rs("DescDocumento")
					end if 
					%>
					<input class="form-control" type="text" readonly value="<%=DescDocumento%>">
				</td>
				<td <%=stileTr%>><%
                      if Rs("FlagObbligatorio")=0 then 
					     SiNo = "NO"
					  else
					     SiNo = "SI"
					  end if 
                    %>
					<input class="form-control" type="text" readonly value="<%=SiNo%>">
                </td>
				<td <%=stileTr%>>
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
				<td <%=stileTr%>>
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

				<td  <%=stileTr%>>
				   <%
				       if PathDocumento="" then 
					      Validazione="Da caricare"
					   else 
					      Validazione=Rs("DescTipoValidazione")
					   end if 
					   if Rs("NoteValidazione")<>"" then 
					      Validazione = Validazione & " - " & Rs("NoteValidazione")
					   end if 
				   %>
					<input class="form-control" type="text" readonly value="<%=Validazione%>">
				</td>
		
			<td>
				<%Linkdocumento=Rs("PathDocumento")%>
				<!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->
				
			    <% if FlagStatoFinale=0 then %>
				
					<%if Rs("FlagObbligatorio")=0 then
					     RiferimentoA="col-2;#;;2;plus;Obbligatorio;;localVal('O','" & id & "');N"
					  else
					     RiferimentoA="col-2;#;;2;minu;Non Obbligatorio;;localVal('X','" & id & "');N"
					  end if 
					%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				    <%if Rs("IdTipoValidazione")<>"VALIDO" then 
					  RiferimentoA="col-2;#;;2;ok;Validare;;localVal('S','" & id & "');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->		
					<%end if %>

					
					<%RiferimentoA="col-2;#;;2;ko;Non Valido;;localValN('" & id & "');N"%>
					<!--#include virtual="/gscVirtual/include/Anchor.asp"-->		
					
				<%end if%> 
	 				
			</td>
			</tr>
		<%	
		    rs.MoveNext
	    Loop
   else
      TuttiCaricati=false 
   end if 
   rs.close

'response.write "nnn:" & ContaDocAssenti & " " & ContaDocKo
'gestisco i dati se è possibile modificare 
flag=false 
flag = Session("LoginTipoUtente")=ucase("BackO") or Session("LoginTipoUtente")=ucase("COLL")

if flag=true and cdbl(FlagStatoFinale)=0  and instr("DOCU_LAVO",IdStatoValidazione)>0 then 
%>
</tbody></table></div> <!-- table responsive fluid -->

            <div class="row">
               <div class="col-1">
               </div>			

               <div class="col-1">
			         <div class="form-group font-weight-bold">Valido Dal
                     </div>	  
               </div>            
               <div class="col-1">
                     <div class="form-group font-weight-bold">
                     <%
                     valo=Stod(Getdiz(DizDatabase,"ValidoDal"))
                     %>
                     <input type="text" name="ValidoDal0" id="ValidoDal0" class="form-control" value="<%=valo%>" >      
                     </div>	  
               </div> 			   
               <div class="col-1">
			         <div class="form-group font-weight-bold">Valido Al
                     </div>	  
               </div> 

               <div class="col-1">
                     <div class="form-group font-weight-bold">
                     <%
                     valo=Stod(Getdiz(DizDatabase,"ValidoAl"))
                     %>
					 <input type="text" name="ValidoAl0" id="ValidoAl0" class="form-control" value="<%=valo%>" >      
                     </div>	  
               </div> 
               <div class="col-2">
                    <button type="button" onclick="ValidaRichiesta('AFFI')"   class="btn btn-success">Valida Richiesta</button>
               </div>			   
           </div>
		   <br>
		   
            <div class="row">
               <div class="col-1">
               </div>			
                 <div class="col-1">
			         <div class="form-group font-weight-bold">Annotazioni
                     </div>	  
               </div> 
		       <div class="col-8">
                  <div class="form-group font-weight-bold">
                     <%
                     DescNote = NoteAffidamentoComp
                     if isCliente() then 
                        DescNote = noteAffidamentoClie
                     end if 
                     
                     %>
                  <input type="textArea" name="NoteValidazione0" id="NoteValidazione0" class="form-control" value="<%=GetDiz(DizDatabase,"NoteValidazione")%>" >      
                  </div>               
               </div> 
            </div>   
            <div class="row">
               <div class="col-1">
               </div>			
   
               <%if IdStatoValidazione="DOCU" then %>
                    <div class="col-2">
                         <button type="button" onclick="registraStato('LAVO')"   class="btn btn-info">Lavorazione</button>
                    </div>
			   <%end if %>
               <%if IdStatoValidazione="LAVO" then %>
                    <div class="col-2">
                         <button type="button" onclick="localRicN('DOCU')"   class="btn btn-info">Integra docum.</button>
                    </div>
			   <%end if %>

					<div class="col-2">
                         <button type="button" onclick="localRicN('RIFI')"   class="btn btn-warning">Rifiuta Richiesta</button>
                    </div>
					<div class="col-2">
                         <button type="button" onclick="localRicN('ANNU')"   class="btn btn-danger">Annulla Richiesta</button>
                    </div>

             </div>            
   
<%end if %>
            <input type="hidden" name="localVirtualPath" id="localVirtualPath" value = "<%=VirtualPath%>">

            
            <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
            <!--#include virtual="/gscVirtual/include/paginazione.asp"-->
            
            </form>
        </div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->


<div class="modal fade" id="myConfirmInvalid"  aria-hidden="true" role="dialog">
  <!-- a pieno schermo -->
  <div class="modal-dialog">
    <div class="modal-content">
        <div class="row bg-warning">
			<div class="col-10 "><h3 class="bg-warning ">Documento non valido</h3>
			</div>
			<div class="col-2">
			</div>
			<div class="col-1"></div>
		</div>
        <div class="row bg-light">
		   <div class="col-1"></div>
		   <div class="col-11">
		   <h4>Motivo per cui il documento non e' valido</h4>  
		   </div>
		</div>
		
        <div class="row  bg-light">
		   <div class="col-1"></div>
		   <div class="col-11">
              <div class="form-group ">
			     <label class='form-check-label font-weight-bold'  style='font-size:11px; margin-top:0px; margin-bottom:0px;'   >informazioni aggiuntive</label>
					 <input type="text"  name="myConfirmIvalidinfo" id="myConfirmIvalidinfo" class="form-control">
                  </div>		
		   </div>
      </div>  
        <div class="row  bg-light">
		   <div class="col-12"></div>
        </div>
      <div class="row bg-light text-center">
	     <div class="col-6">
         <button type="button" onclick="myConfirmIvalidYes();" class="btn btn-success" data-dismiss="modal">Conferma</button>
		 </div>
         <div class="col-6">
         <button type="button" class="btn btn-danger" data-dismiss="modal">Annulla</button>
		 </div>
      </div>

    </div>
  </div>
</div>


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
      $('[data-toggle="tooltip" = Rs("")').tooltip();   
    });
  </script>

</body>

</html>
