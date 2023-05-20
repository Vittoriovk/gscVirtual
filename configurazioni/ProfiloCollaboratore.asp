<%
  NomePagina="ProfiloCollaboratore.asp"
  titolo="Profilo"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
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
<!--#include virtual="/gscVirtual/js/functionTable.js"-->
<script language="JavaScript">

function localFun(Op,Id)
{
	xx=ImpostaValoreDi("DescLoaded","0");
	xx=ElaboraControlli();
	
 	if (xx==false)
	   return false;

	ImpostaValoreDi("Oper","update");
	document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

 
 <!-- javascript locale -->
<script>
function localSubmit(Op)
{
var xx;
    xx=false;
	if (Op=="submit")
	   xx=ElaboraControlli();
   	
 	if (xx==false)
	   return false;
		
	ImpostaValoreDi("Oper","update");
	document.Fdati.submit(); 
}
</script>

<%

  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdCollaboratore = cdbl(Session("LoginIdCollaboratore"))
  IdAccount       = cdbl(Session("LoginIdAccount"))
 
  'puo' aggiornare solo password 
   if Oper=ucase("update") then 

	  if request("checkCrea0")="S" then 
	     FlagGeneraCollaboratore=1
	  else
	     FlagGeneraCollaboratore=0
	  end if 



      if Cdbl(IdCollaboratore)=0 then 
		 if Cdbl(IdAccount)=0 then 
		    MsgErrore="errore di sistema : contattare assistenza"
         else 
		    myUpd = ""
		    myUpd = myUpd & " Update Account Set IdAzienda=" & Session("IdAziendaWork") 
			myUpd = myUpd & ",IdTipoAccount='Coll'"
			myUpd = myUpd & ",FlagAttivo='S',Abilitato=0"
			myUpd = myUpd & ",Nominativo='" & apici(Nominativo) & "'"
			myUpd = myUpd & " Where IdAccount=" & IdAccount
			ConnMsde.execute = MyUpd
			NextL=cdbl(session("LivelloAccount") + 1)
            MyQ = "" 
            MyQ = MyQ & " Insert into Collaboratore (IdAccount,IdAzienda,Denominazione,IdTipoDitta,Livello)"
            MyQ = MyQ & " values (" & IdAccount & "," & Session("IdAziendaWork") & ",'" & Apici(Nominativo) & "'"
            MyQ = MyQ & ",'" & apici(IdTipoPers) & "'," & NextL & ")"
            ConnMsde.execute MyQ 
			'response.write MyQ
            If Err.Number <> 0 Then 
               MsgErrore = ErroreDb(Err.description)
			   IdAccount=0
            else
			   Ritorna=true 
               IdCollaboratore = GetTableIdentity("Collaboratore")    
            end if 
         end if 
      end if 
      'aggiorno Collaboratore 

   end if 

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
	 
   'recupero i dati 
   if cdbl(IdCollaboratore)>0 then
      MySql = ""
      MySql = MySql & " Select * From  Collaboratore "
      MySql = MySql & " Where IdCollaboratore=" & IdCollaboratore
      xx=GetInfoRecordset(DizDatabase,MySql)
      IdTipoPers=Getdiz(DizDatabase,"IdTipoDitta")
      IdTipoColl=Getdiz(DizDatabase,"IdTipoCollaboratore")
      IdAccount =Cdbl(Getdiz(DizDatabase,"IdAccount"))
  end if 
     
   DescPageOper="Aggiornamento"

  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdCollaboratore"  ,IdCollaboratore)
  xx=setValueOfDic(Pagedic,"IdAccount"        ,IdAccount)
 
  xx=setCurrent(NomePagina,livelloPagina) 
  DescLoaded="0"  
  %>
<% 
  'xx=DumpDic(SessionDic,NomePagina)
%>
<div class="d-flex" id="wrapper">

	<%
	  Session("opzioneSidebar")="prof"
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
				<div class="col-11"><h3>Profilo Collaboratore :</b> <%=DescPageOper%> </h3>
				</div>
			</div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

    <%
	  stdClass="class='form-control form-control-sm'"
      l_Id = "0"
	  err.clear
      SoloLettura=true
      ReadOnly=" readonly "
   
   %>
   <div class="row">
      <div class="col-2"><p class="font-weight-bold">Tipo Collaboratore</p></div>   
      <div class = "col-3">
	     <%
		 q = ""
		 q = q & "Select * from TipoCollaboratore "
		 if Cdbl(IdCollaboratore)>0 then 
            IdTipoColl=Getdiz(DizDatabase,"IdTipoCollaboratore") 
		 end if 
		 if IdTipoColl<>"" then 
		    q = q & " where IdTipoCollaboratore='" & IdTipoColl & "'"
		 else 
		    NextL=cdbl(session("LivelloAccount") + 1)
		    q = q & " where LivelloMinimo>=" & NextL & " and LivelloMassimo<= " & NextL 
		 end if 
		 q = q & " order By DescTipoCollaboratore"
	     response.write ListaDbChangeCompleta(q,"IdTipoCollaboratore" & l_Id,IdTipoColl ,"IdTipoCollaboratore","DescTipoCollaboratore" ,0,"","","","","",stdClass)
	     %>
      </div>

      <div class="col-2"><p class="font-weight-bold">Tipo Utenza</p></div>   
      <div class = "col-3">
	     <%
		 q = "select * from TipoDitta where IdTipoditta='" & apici(IdTipoPers) & "'"
		 valo = LeggiCampo(q,"DescTipoDitta")
	     %>
		 <input type="text" readonly class="form-control" value="<%=valo%>" >	 
      </div>
      <div class="col-2"><p class="font-weight-bold"> </p>
      </div> 
   </div>
   <%
   NameLoaded= ""
   if IdTipoPers="PEGI" then
      NameLoaded= NameLoaded & "Denominazione,TE"   		  
   elseif  IdTipoPers="DITT" then 
      NameLoaded= NameLoaded & "Denominazione,TE"   		  
      NameLoaded= NameLoaded & ";Cognome,TE"   		  
	  NameLoaded= NameLoaded & ";Nome,TE"   		  
   else 
      NameLoaded= NameLoaded & "Cognome,TE"   		  
	  NameLoaded= NameLoaded & ";Nome,TE"   		  
   end if 
   %>
   
   
   <%if IdTipoPers="PEGI" or IdTipoPers="DITT" then%>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Denominazione</p>
      </div> 
	  <div class="col-8">
	  	  <%
 
		  nome="Denominazione" & l_id
		  valo=Getdiz(DizDatabase,"Denominazione")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>    
   <%end if%> 
   
   <%if IdTipoPers="PEFI" or IdTipoPers="DITT" then
      lblC0 = "Cognome" 
	  lblC1 = "Nome"
      colss = "col-3"
   %>

   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold"><%=lblC0%></p>
      </div> 
	  <div class="<%=colss%>">
	  	  <%
		  nome="Cognome" & l_id
		  valo=Getdiz(DizDatabase,"Cognome")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div>
	  <%if IdTipoPers<>"PEGI" then%>
      <div class="col-2">
         <p class="font-weight-bold"><%=lblC1%></p>
      </div> 
	  <%end if %>
	  <div class="<%=colss%>">
	  	  <%
		  nome="Nome" & l_id
		  valo=Getdiz(DizDatabase,"Nome")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>   
   <%end if %>
   <%if IdTipoPers="PEFI" or IdTipoPers="DITT" then%>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Stato Nascita</p>
      </div>
	  <div class="col-3">
	  	  <%
          NameLoaded= NameLoaded & ";StatoNascita,TE"   		  
		  nome="StatoNascita" & l_id
		  valo=Getdiz(DizDatabase,"StatoNascita")
		  if valo="" then 
		     valo="Italia"
		  end if 
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div> 
      <div class="col-2">
         <p class="font-weight-bold">Comune Nascita</p>
      </div> 
	  <div class="col-3">
	  	  <%
          NameLoaded= NameLoaded & ";ComuneNascita,TE"   		  
		  nome="ComuneNascita" & l_id
		  valo=Getdiz(DizDatabase,"ComuneNascita")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div> 
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Provincia Nascita</p>
      </div> 
	  <div class="col-3">
	  	  <%
          NameLoaded= NameLoaded & ";ProvinciaNascita,TE"   		  
		  nome="ProvinciaNascita" & l_id
		  valo=Getdiz(DizDatabase,"ProvinciaNascita")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div> 	  
      <div class="col-2">
         <p class="font-weight-bold">Data Di Nascita</p>
      </div> 
	  <div class="col-3">
	  	  <%
          NameLoaded= NameLoaded & ";DataNascita,TE"   		  
		  nome="DataNascita" & l_id
		  valo=StoD(Getdiz(DizDatabase,"DataNascita"))
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
	  </div> 	  	  
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>	  
   </div>

	  
   <%end if %>
   
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Codice fiscale</p>
      </div> 
	  <div class="col-3">
		  <%
		  NameLoaded= NameLoaded & ";CodiceFiscale,CF" 
		  nome="CodiceFiscale" & l_id
		  valo=Getdiz(DizDatabase,"CodiceFiscale")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
	  <%if IdTipoPers="PEFI" then%>
      <div class="col-7">
         <p class="font-weight-bold"> </p>
      </div>
	  
	  <% else %>
      <div class="col-2">
         <p class="font-weight-bold">Partita Iva</p>
      </div> 
	  	  <div class="col-3">
		  <%
		  NameLoaded= NameLoaded & ";PartitaIva,PI" 
		  nome="PartitaIva" & l_id
		  valo=Getdiz(DizDatabase,"PartitaIva")
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
	  <% end if  %>
   </div> 

   <div class="row">
      <div class="col-2"><p class="font-weight-bold">Sezione Rui</p></div>   
      <div class = "col-2">
	     <%
		 valo=Getdiz(DizDatabase,"SezioneRui")
		 NextL=cdbl(session("LivelloAccount") + 1)
		 q = ""
		 q = q & " SELECT * From TipoRui where LivelloMinimo>=" & NextL & " and LivelloMassimo<= " & NextL & " order By DescTipoRui  "
	     response.write ListaDbChangeCompleta(q,"SezioneRui" & l_Id,valo ,"IdTipoRui","DescTipoRui" ,0,"","","","","",stdClass)
	     %>
      </div>

      <div class="col-1"><p class="font-weight-bold">Num. RUI</p></div>   
      <div class = "col-2">
		  <%
		  NameLoaded= NameLoaded & ";NumeroRui,TE" 
		  nome="NumeroRui" & l_id
		  valo=Getdiz(DizDatabase,"NumeroRui")
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
      </div>
	  <div class="col-1"><p class="font-weight-bold">Iscritto il</p></div>
      <div class = "col-2">
		  <%
		  NameLoaded= NameLoaded & ";DataIscrizioneRui,DTO" 
		  nome="DataIscrizioneRui" & l_id
		  valo=StoD(Getdiz(DizDatabase,"DataIscrizioneRui"))
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >		  
      </div>
   </div>
   

    <%if IdTipoPers="PEGI" then%>
	
	  <a class="btn btn-info" data-toggle="collapse" href="#collapseAmministratore" role="button" 
		 aria-expanded="false" aria-controls="collapseAmministratore">
		 <span Id="Amministratore_Plus"><i class="fa fa-1x fa-plus-circle"></i></span>
		 <span Id="Amministratore_Minus" style= "display:none"><i class="fa fa-1x fa-minus-circle"></i></span>
		 <input type="hidden" id="Amministratore_plusMinus" value = "+">
		 </a>
		 <B> Amministratore </B>
	  </p> 
	  

		<div class="collapse" id="collapseAmministratore">
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Cognome Amministratore</p></div> 
			  <div class="col-3">
				  <%
				  nome="CognomeAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"CognomeAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Nome Amministratore</p></div> 
			  <div class="col-3">
				  <%
				  nome="NomeAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"NomeAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p>
			  </div>
		   </div> 

		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Codice fiscale</p></div> 
			  <div class="col-3">
				  <%
				  nome="CodiceFiscaleAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"CodiceFiscaleAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Partita Iva</p></div> 
				  <div class="col-3">
				  <%
				  nome="PartitaIvaAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"PartitaIvaAmministratore")
				  %>
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
   
		   <div class="row">
			  <div class="col-2">
				 <p class="font-weight-bold">Stato Nascita</p>
			  </div>
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "StatoNascitaAmministratore,TE"   		  
				  nome="StatoNascitaAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"StatoNascitaAmministratore")
				  if valo="" then 
					 valo="Italia"
				  end if 
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div> 
			  <div class="col-2">
				 <p class="font-weight-bold">Comune Nascita</p>
			  </div> 
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "ComuneNascitaAmministratore,TE"   		  
				  nome="ComuneNascitaAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"ComuneNascitaAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div> 
			  <div class="col-2">
				 <p class="font-weight-bold"> </p>
			  </div>
		   </div>
		   <div class="row">
			  <div class="col-2">
				 <p class="font-weight-bold">Provincia Nascita</p>
			  </div> 
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "ProvinciaNascitaAmministratore,TE"   		  
				  nome="ProvinciaNascitaAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"ProvinciaNascitaAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div> 	  
			  <div class="col-2">
				 <p class="font-weight-bold">Data Di Nascita</p>
			  </div> 
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "DataNascitaAmministratore,TE"   		  
				  nome="DataNascitaAmministratore" & l_id
				  valo=StoD(Getdiz(DizDatabase,"DataNascitaAmministratore"))
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
						 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
			  </div> 	  	  
			  <div class="col-2">
				 <p class="font-weight-bold"> </p>
			  </div>	  
		   </div>
		
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Indirizzo</p></div> 
			  <div class="col-3">
				  <%
				  nome="IndirizzoAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"IndirizzoAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Citta'</p></div> 
				  <div class="col-3">
				  <%
				  nome="CittaAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"CittaAmministratore")
				  %>
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Provincia</p></div> 
			  <div class="col-3">
				  <%
				  nome="ProvinciaAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"ProvinciaAmministratore")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Cap</p></div> 
				  <div class="col-3">
				  <%
				  nome="CapAmministratore" & l_id
				  valo=Getdiz(DizDatabase,"CapAmministratore")
				  %>
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
		
		</div>  <!-- fine sezione amministratore -->

	  <a class="btn btn-info" data-toggle="collapse" href="#collapsePreposto" role="button" 
		 aria-expanded="false" aria-controls="collapsePreposto">
		 <span Id="Preposto_Plus"><i class="fa fa-1x fa-plus-circle"></i></span>
		 <span Id="Preposto_Minus" style= "display:none"><i class="fa fa-1x fa-minus-circle"></i></span>
		 <input type="hidden" id="Preposto_plusMinus" value = "+">
		 </a>
		 <B> Preposto </B>
	  </p> 
	  

		<div class="collapse" id="collapsePreposto">
   <div class="row">
      <div class="col-2"><p class="font-weight-bold">Cognome Preposto</p></div> 
	  <div class="col-3">
	  	  <%
		  nome="CognomePreposto" & l_id
		  valo=Getdiz(DizDatabase,"CognomePreposto")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
	  </div>
      <div class="col-2"><p class="font-weight-bold">Nome Preposto</p></div> 
	  <div class="col-3">
	  	  <%
		  nome="NomePreposto" & l_id
		  valo=Getdiz(DizDatabase,"NomePreposto")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
      <div class="col-2"><p class="font-weight-bold"> </p>
      </div>
   </div> 

   <div class="row">
      <div class="col-2"><p class="font-weight-bold">Codice fiscale Preposto</p></div> 
	  <div class="col-3">
		  <%
		  nome="CodiceFiscalePreposto" & l_id
		  valo=Getdiz(DizDatabase,"CodiceFiscalePreposto")
		  %>	  
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
      <div class="col-2"><p class="font-weight-bold">Partita Iva Preposto</p></div> 
	  	  <div class="col-3">
		  <%
		  nome="PartitaIvaPreposto" & l_id
		  valo=Getdiz(DizDatabase,"PartitaIvaPreposto")
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
	  </div>
      <div class="col-2"><p class="font-weight-bold"> </p></div>
   </div> 
		   <div class="row">
			  <div class="col-2">
				 <p class="font-weight-bold">Stato Nascita</p>
			  </div>
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "StatoNascitaPreposto,TE"   		  
				  nome="StatoNascitaPreposto" & l_id
				  valo=Getdiz(DizDatabase,"StatoNascitaPreposto")
				  if valo="" then 
					 valo="Italia"
				  end if 
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div> 
			  <div class="col-2">
				 <p class="font-weight-bold">Comune Nascita</p>
			  </div> 
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "ComuneNascitaPreposto,TE"   		  
				  nome="ComuneNascitaPreposto" & l_id
				  valo=Getdiz(DizDatabase,"ComuneNascitaPreposto")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div> 
			  <div class="col-2">
				 <p class="font-weight-bold"> </p>
			  </div>
		   </div>
		   <div class="row">
			  <div class="col-2">
				 <p class="font-weight-bold">Provincia Nascita</p>
			  </div> 
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "ProvinciaNascitaPreposto,TE"   		  
				  nome="ProvinciaNascitaPreposto" & l_id
				  valo=Getdiz(DizDatabase,"ProvinciaNascitaPreposto")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >	  
			  </div> 	  
			  <div class="col-2">
				 <p class="font-weight-bold">Data Di Nascita</p>
			  </div> 
			  <div class="col-3">
				  <%
				  NameLoaded= ""
				  NameLoaded= NameLoaded & "DataNascitaPreposto,TE"   		  
				  nome="DataNascitaPreposto" & l_id
				  valo=StoD(Getdiz(DizDatabase,"DataNascitaPreposto"))
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
						 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
			  </div> 	  	  
			  <div class="col-2">
				 <p class="font-weight-bold"> </p>
			  </div>	  
		   </div>
  
   
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Indirizzo Preposto</p></div> 
			  <div class="col-3">
				  <%
				  nome="IndirizzoPreposto" & l_id
				  valo=Getdiz(DizDatabase,"IndirizzoPreposto")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Citta' Preposto</p></div> 
				  <div class="col-3">
				  <%
				  nome="CittaPreposto" & l_id
				  valo=Getdiz(DizDatabase,"CittaPreposto")
				  %>
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
		   <div class="row">
			  <div class="col-2"><p class="font-weight-bold">Provincia Preposto</p></div> 
			  <div class="col-3">
				  <%
				  nome="ProvinciaPreposto" & l_id
				  valo=Getdiz(DizDatabase,"ProvinciaPreposto")
				  %>	  
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold">Cap Preposto</p></div> 
				  <div class="col-3">
				  <%
				  nome="CapPreposto" & l_id
				  valo=Getdiz(DizDatabase,"CapPreposto")
				  %>
				  <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
			  </div>
			  <div class="col-2"><p class="font-weight-bold"> </p></div>
		   </div> 
		   
   <div class="row">
      <div class="col-2"><p class="font-weight-bold">Sezione Rui</p></div>   
      <div class = "col-2">
	     <%
		 valo=Getdiz(DizDatabase,"SezioneRuiPreposto")
		 q = ""
		 q = q & " SELECT * From TipoRui where LivelloMinimo>=0 and LivelloMassimo<=99 order By DescTipoRui  "
	     response.write ListaDbChangeCompleta(q,"SezioneRuiPreposto" & l_Id,valo ,"IdTipoRui","DescTipoRui" ,0,"","","","","",stdClass)
	     %>
      </div>

      <div class="col-1"><p class="font-weight-bold">Num. RUI</p></div>   
      <div class = "col-2">
		  <%
		  NameLoaded= NameLoaded & ";NumeroRuiPreposto,TE" 
		  nome="NumeroRuiPreposto" & l_id
		  valo=Getdiz(DizDatabase,"NumeroRuiPreposto")
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
      </div>
	  <div class="col-1"><p class="font-weight-bold">Iscritto il</p></div>
      <div class = "col-2">
		  <%
		  NameLoaded= NameLoaded & ";DataIscrizioneRuiPreposto,DTO" 
		  nome="DataIscrizioneRuiPreposto" & l_id
		  valo=StoD(Getdiz(DizDatabase,"DataIscrizioneRuiPreposto"))
		  %>
	      <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >		  
      </div>
   </div>
</div>

   <%end if %>

   <!--#include virtual="/gscVirtual/include/setDataForCall.asp"--> 
   
   <%
   NomeStruttura     = "SEDI_COLLABORATORE"
   DescStruttura     = "Sedi Collaboratore"
   flagOperStruttura = ""
   ProfiloAccount    = "COLL"
   %>
   <!--#include virtual="/gscVirtual/configurazioni/sedi/StrutturaSede.asp"-->    

   <%
   NomeStruttura     = "CONTATTI_COLLABORATORE"
   DescStruttura     = "Contatti Collaboratore"
   flagOperStruttura = ""
   ProfiloAccount    = "COLL"
   %>
   <!--#include virtual="/gscVirtual/configurazioni/contatti/StrutturaContatto.asp"--> 
   
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div></div>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
		</div>
   
			<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
			</form>
<!--#include virtual="/gscVirtual/include/FormSoggetti.asp"-->
		</div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

</body>

</html>
