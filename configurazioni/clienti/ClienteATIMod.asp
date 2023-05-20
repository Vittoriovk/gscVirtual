<%
  NomePagina="ClienteATIMod.asp"
  titolo="Modifica ATI per cliente"
  default_check_profile="Clie,Coll"
  
  act_call_dett = CryptAction("CALL_DETT")
  
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/functionCf.asp"-->
<!--#include virtual="/gscVirtual/common/functionCfScript.asp"-->
<!--#include virtual="/gscVirtual/common/functionDataList.asp"-->
<!--#include virtual="/gscVirtual/common/functionDataListScript.asp"-->
<!--#include virtual="/gscVirtual/js/functionLocalita.js"-->
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
function localDett()
{
    xx=ImpostaValoreDi("Oper",'<%=act_call_dett%>');
    document.Fdati.submit();
}

function localFun(Op,Id)
{
    xx=ImpostaValoreDi("DescLoaded","0");
    xx=ElaboraControlli();
    
     if (xx==false)
       return false;

    ImpostaValoreDi("Oper","update");
    document.Fdati.submit();
}

function calcolaCFClie()
{
   var nome=$("#Cognome0").val();
   var cogn=$("#Nome0").val();
   var sess=$("#IdSesso0").val();;
  
   var dtna=$("#DataNascita0").val();
   var stat=$("#StatoNascita0").val();
   var comu=$("#ComuneNascita0").val();
   var prov=$("#ProvinciaNascita0").val();
   var cf = calcolaCF(nome,cogn,sess,dtna,stat,prov,comu);
   if (cf.length>0)
      $("#CodiceFiscale0").val(cf);
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<!--#include virtual="/gscVirtual/modelli/FunctionAccount.asp"-->
  
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
  IdAccountCliente=0
  idATI=0
  if FirstLoad then 
     idATI      = "0" & Session("swap_idATI")
     if Cdbl(idATI)=0 then 
        idATI   = cdbl("0" & getValueOfDic(Pagedic,"idATI"))
     end if   
     IdAccountCliente   = "0" & Session("swap_IdAccCliente")
     if Cdbl(IdAccountCliente)=0 then 
        IdAccountCliente = cdbl("0" & getValueOfDic(Pagedic,"IdAccountCliente"))
     end if      
     OperAmmesse   = Session("swap_OperAmmesse")
     OperTabella   = Session("swap_OperTabella")
     PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
     if PaginaReturn="" then 
        PaginaReturn = Session("swap_PaginaReturn")
     end if 
     if Cdbl(idATI)=0 then 
        IdTipoPers=Session("swap_IdPersCliente")
     end if 

  else
     idATI            = "0" & getValueOfDic(Pagedic,"idATI")
     IdAccountCliente = "0" & getValueOfDic(Pagedic,"IdAccountCliente")
     OperAmmesse      = getValueOfDic(Pagedic,"OperAmmesse")
     OperTabella      = getValueOfDic(Pagedic,"OperTabella") 
     PaginaReturn     = getValueOfDic(Pagedic,"PaginaReturn")
     IdTipoPers       = getValueOfDic(Pagedic,"IdTipoPers")
   end if 
   idATI = cdbl(idATI)

   if OperAmmesse="" then 
      if idATI = 0 then 
         OperAmmesse="CRUD"
      end if 
   end if 
   Oper = DecryptAction(Oper)
   if Oper="CALL_DETT" then 
      xx=RemoveSwap()
      Session("TimeStamp")=TimePage
      KK=idATI
      if Cdbl("0" & KK ) > 0 then 
         Session("swap_IdListaDocumento")= KK
         Session("swap_OperTabella")     = Oper
         Session("swap_TipoRife") = "ATI"
         Session("swap_IdRife")   = KK
         Session("swap_PaginaReturn")    = PaginaReturn
         response.redirect virtualPath   & "configurazioni/clienti/AffidamentoAtiCoob.asp"
         response.end 
      end if 
   End if 
   
  'inserisco account 
   if Oper=ucase("update") then 
      Ritorna=false 
      OperAmmesse="U"
      Session("TimeStamp")=TimePage
      MsgErrore=""
      Cognome    = Request("Cognome0")
      Nome       = Request("Nome0")
      Nominativo = Request("RagSoc0")
      if Nominativo = "" then 
         Nominativo =  trim(trim(Cognome) & " " & Trim(Nome)) 
      end if 

      StatoNascita     = Request("StatoNascita0")
      ComuneNascita    = Request("ComuneNascita0")
      ProvinciaNascita = Request("ProvinciaNascita0")
      DataNascita      = Request("DataNascita0")
      if len(DataNascita)<>10 then 
         DataNascita=0
      else 
         DataNascita=DataStringa(DataNascita)
      end if 
      IdSesso       = Request("IdSesso0")
      CodiceFiscale = Request("CodiceFiscale0")
      PartititaIva  = Request("PartitaIva0")
      Indirizzo     = trim(Request("Indirizzo0"))
	  IdStato       = trim(Request("IdStato0"))
      Comune        = trim(Request("Comune0"))
      Provincia     = trim(Request("Provincia0"))
	  Cap           = trim(Request("Cap0"))
      TipoSocieta   = Request("TipoSocietaGR")
      if Cdbl(idATI)=0 then 
         MyQ = MyQ & " insert into AccountATI"
         MyQ = MyQ & "(IdAccount,IdTipoDitta,PI,CF,RagSoc,flagValidato,Note) values "
         MyQ = MyQ & "(" & IdAccountCliente & ",'" & IdTipoPers & "','" & PI & "','" & CF & "','" & Nominativo & "',0,'')"  
		 ConnMsde.execute MyQ 
		 if Err.Number=0 then 
		    idATI=GetTableIdentity("AccountATI")
		 end if 
	  end if 
      if Cdbl(idATI)>0 then 
	     MyQ = ""
         MyQ = MyQ & " update AccountATI set "
         MyQ = MyQ & " PI ='" & apici(PartititaIva) & "'"  
         MyQ = MyQ & ",CF ='" & apici(CodiceFiscale) & "'"   
         MyQ = MyQ & ",RagSoc = '" & apici(Nominativo) & "'"   
         MyQ = MyQ & ",Indirizzo = '" & apici(Indirizzo) & "'"   
         MyQ = MyQ & ",Cap = '" & apici(Cap) & "'"   
         MyQ = MyQ & ",Comune = '" & apici(Comune) & "'"   
         MyQ = MyQ & ",Provincia = '" & apici(Provincia) & "'" 
         MyQ = MyQ & ",IdStato= '" & apici(IdStato) & "'" 
         MyQ = MyQ & ",Cognome= '" & apici(Cognome) & "'"  
         MyQ = MyQ & ",Nome= '" & apici(Nome) & "'" 
         MyQ = MyQ & ",StatoNascita= '" & apici(StatoNascita) & "'" 
         MyQ = MyQ & ",ComuneNascita= '" & apici(ComuneNascita) & "'" 
         MyQ = MyQ & ",ProvinciaNascita= '" & apici(ProvinciaNascita) & "'" 
         MyQ = MyQ & ",DataNascita=" & NumForDb(DataNascita)
         MyQ = MyQ & ",IdSesso= '" & apici(IdSesso) & "'" 
		 MyQ = MyQ & ",TipoSocieta= '" & apici(TipoSocieta) & "'" 
         MyQ = MyQ & " Where IdAccountATI = " & idATI 
         MyQ = MyQ & " and   IdAccount=" & IdAccountCliente
		 'response.write MyQ 
		 
		 ConnMsde.execute MyQ 
      end if 

   end if 

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
     
   'recupero i dati 
   if cdbl(idATI)>0 then
      MySql = ""
      MySql = MySql & " Select * From  AccountATI "
      MySql = MySql & " Where IdAccountATI=" & idATI
      xx=GetInfoRecordset(DizDatabase,MySql)
      IdTipoPers  = Getdiz(DizDatabase,"IdTipoDitta")
	  TipoSocieta = Getdiz(DizDatabase,"TipoSocieta")
	  if Cdbl("0" & Getdiz(DizDatabase,"FlagValidato"))=1 then 
	     OperAmmesse="R"
	  end if 
   end if 
     
   DescPageOper="Aggiornamento"
   if OperAmmesse="R" then 
      DescPageOper = "Consultazione"
   elseIf cdbl(idATI)=0 then 
      DescPageOper = "Inserimento"
   end if
  'registro i dati della pagina 
   xx=setValueOfDic(Pagedic,"idATI"            ,idATI)
   xx=setValueOfDic(Pagedic,"IdAccountCliente" ,IdAccountCliente)
   xx=setValueOfDic(Pagedic,"OperAmmesse"      ,OperAmmesse)
   xx=setValueOfDic(Pagedic,"IdTipoPers"       ,IdTipoPers)  
   xx=setValueOfDic(Pagedic,"PaginaReturn"     ,PaginaReturn)
   xx=setValueOfDic(Pagedic,"OperTabella"      ,OperTabella)
 
   xx=setCurrent(NomePagina,livelloPagina) 
   DescLoaded="0"  
  %>
<% 
  'xx=DumpDic(SessionDic,NomePagina)
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
            <%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <div class="col-11"><h3>Gestione A.T.I. :</b> <%=DescPageOper%> </h3>
                </div>
            </div>
            <%if Iscliente()=false then %>
                <div class="row">
                   <div class = "col-6">
                        <div class="form-group ">
                        <%xx=ShowLabel("Cliente")
						DescCliente=LeggiCampo("select * from Cliente Where IdAccount=" & IdAccountCliente,"Denominazione")
						%>
                        <input value="<%=DescCliente%>" readonly class="form-control"  >
                        </div>
                  </div>
               </div>
			   <br>
            <%end if %>			
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

    <%
      stdClass="class='form-control form-control-sm'"
      l_Id = "0"
      err.clear
      ReadOnly=""
      SoloLettura=false
      if instr(OperAmmesse,"U")=0 or (instr(OperAmmesse,"I")>0 and cdbl("0" & idATI)>0) then 
         SoloLettura=true
         ReadOnly=" readonly "
      end if 
   
   %>
   <div class="row">
      <div class="col-2"><p class="font-weight-bold">Tipo Utenza</p></div>   
      <div class = "col-3">
         <%
         q = "select * from TipoDitta where IdTipoditta='" & apici(IdTipoPers) & "'"
         valo = LeggiCampo(q,"DescTipoDitta")
         %>
         <input type="text" readonly class="form-control" value="<%=valo%>" >     
      </div>
	  <%if IdTipoPers = "PEGI" then
           TipoCapi=""
		   TipoPers=""
		   if TipoSocieta="CAPI" then 
		      TipoCapi = " checked "
		   else
		      TipoPers = " checked "
		   end if 
	  %>
	  <div class="col-3"><B>Societ&agrave; di </B>
	        &nbsp;&nbsp;
            <input name="TipoSocietaGR" type="radio" <%=TipoCapi%> id="TipoSocieta1" value="CAPI" >&nbsp;<B>Capitale</B>
			&nbsp;&nbsp;
            <input name="TipoSocietaGR" type="radio" <%=TipoPers%> id="TipoSocieta2" value="PERS" >&nbsp;<B>Persone</B>
      </div> 
	  <%end if %>	  
      <div class="col-2"><p class="font-weight-bold"> </p>
      </div> 
   </div>
   <%
   NameLoaded= ""
   NameLoaded= NameLoaded & "RagSoc,TE"
   if IdTipoPers="DITT" then
      NameLoaded= NameLoaded & "Cognome,TE"             
      NameLoaded= NameLoaded & ";Nome,TE"             
   end if 
   showElabCf = false 
   %>
   
   
   <%if IdTipoPers="PEGI" or IdTipoPers="DITT" then%>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Denominazione</p>
      </div> 
      <div class="col-8">
            <%
          nome="RagSoc" & l_id
          valo=Getdiz(DizDatabase,"RagSoc")
          %>      
          <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >      
      </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>    
   <%end if%> 
   
   <%if IdTipoPers="DITT" then
      lblC0 = "Cognome" 
      lblC1 = "Nome"
      colss = "col-3"
      showElabCf = true 
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
   <%if IdTipoPers="DITT" then%>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Stato Nascita</p>
      </div>
      <div class = "col-3">
         <%
         valo=Getdiz(DizDatabase,"StatoNascita")
         if valo="" then 
            valo="IT"
         end if
         IdStato=Valo 
         q = ""
         q = q & "Select * from Stato "
         if readonly<>"" then 
            q = q & " Where IdStato='" & valo & "'" 
         end if 
         q = q & " order By DescStato"
         onC = "absoluteChangeStato('StatoNascita0','ProvinciaNascita0');"
		 addClass=""
		 if readonly<>"" then 
		    addclass=" disabled "
		 end if 		 
         response.write ListaDbChangeCompleta(q,"StatoNascita" & l_Id,valo ,"IdStato","DescStato" ,0,onC,"","","","",stdClass & addClass)
         %>
       </div>
      <div class="col-2">
         <p class="font-weight-bold">Provincia Nascita</p>
      </div> 
      <div class="col-3">
            <%
          NameLoaded= NameLoaded & ";ProvinciaNascita,TE"             
          nome="ProvinciaNascita" & l_id
          valo=Getdiz(DizDatabase,"ProvinciaNascita")
          idProvincia=valo
          onC = "absoluteChangeProvincia('StatoNascita0','ProvinciaNascita0','ComuneNascita0');"
          listDataProv=""
          listDataComu=""
          if IdStato="IT" then 
             listDataProv="absoluteProvinciaIT"
             if IdProvincia<>"" then 
                IdProvincia = getSiglaProvinciaDaProvincia(IdProvincia)
                if IdProvincia<>"" then 
                   listDataComu="absoluteComune" & IdProvincia
                end if 
             end if 
             
          end if 
          
          %>      
          <input type="text" list="<%=listDataProv%>" onchange="<%=onC%>" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >      
      </div>         
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Comune Nascita</p>
      </div> 
      <div class="col-3">
            <%
          NameLoaded= NameLoaded & ";ComuneNascita,TE"             
          nome="ComuneNascita" & l_id
          valo=Getdiz(DizDatabase,"ComuneNascita")
          %>      
          <input type="text" list="<%=listDataComu%>" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >      
      </div> 
      <div class="col-1">
         <p class="font-weight-bold">Data Nascita</p>
      </div> 
      <div class="col-2">
            <%
          NameLoaded= NameLoaded & ";DataNascita,TE"             
          nome="DataNascita" & l_id
          valo=StoD(Getdiz(DizDatabase,"DataNascita"))
          %>      
          <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" value="<%=valo%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
      </div>
      <div class="col-1">
         <p class="font-weight-bold">Sesso</p>
      </div>
      <div class="col-1">
               <%
         valo=Getdiz(DizDatabase,"IdSesso")
         if valo="" then 
            valo="M"
         end if
         q = ""
         q = q & "Select * from Sesso "
         if readonly<>"" then 
            q = q & " Where IdSesso='" & valo & "'" 
         end if 
         q = q & " order By DescSesso"
		 addClass=""
		 if readonly<>"" then 
		    addclass=" disabled "
		 end if 
         response.write ListaDbChangeCompleta(q,"IdSesso" & l_Id,valo ,"IdSesso","DescSesso" ,0,"","","","","",stdClass & addClass)
         %>
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
		  if IdTipoPers = "PEGI" then
		     NameLoaded= NameLoaded & ";CodiceFiscale,PG" 
			 showElabCf=false
		  else 
		     NameLoaded= NameLoaded & ";CodiceFiscale,CF" 
		  end if 		  
          nome="CodiceFiscale" & l_id
          valo=Getdiz(DizDatabase,"CF")
          %>      
          <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
      </div>
          <%if Readonly="" and showElabCf then %>
            <a href="#" title="Traduci" onclick="reverseCF('CodiceFiscale0','','','ComuneNascita0','ProvinciaNascita0','','IdSesso0','DataNascita0','','StatoNascita0')">  
               <i class="fa fa-2x fa-retweet"></i>
            </a>
            <a href="#" title="Calcola" onclick="calcolaCFClie()">  
               <i class="fa fa-2x fa-id-card-o"></i>
            </a>            
          <%end if %>      
      <%if IdTipoPers="DITT" then%>
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
          valo=Getdiz(DizDatabase,"PI")
          %>
          <input type="text" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >
      </div>
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
      <% end if  %>
   </div> 

   
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Stato</p>
      </div>
      <div class = "col-3">
         <%
         valo=Getdiz(DizDatabase,"IdStato")
         if valo="" then 
            valo="IT"
         end if
         IdStato=Valo 
         q = ""
         q = q & "Select * from Stato "
         if readonly<>"" then 
            q = q & " Where IdStato='" & valo & "'" 
         end if 
         q = q & " order By DescStato"
         onC = "absoluteChangeStato('IdStato0','Provincia0');"
		 addClass=""
		 if readonly<>"" then 
		    addclass=" disabled "
		 end if 		 
         response.write ListaDbChangeCompleta(q,"IdStato" & l_Id,valo ,"IdStato","DescStato" ,0,onC,"","","","",stdClass & addClass)
         %>
       </div>
      <div class="col-1">
         <p class="font-weight-bold">Provincia</p>
      </div> 
      <div class="col-4">
            <%
          NameLoaded= NameLoaded & ";Provincia,TE"             
          nome="Provincia" & l_id
          valo=Getdiz(DizDatabase,"Provincia")
          idProvinciaResi=valo
          onC = "absoluteChangeProvincia('IdStato0','Provincia0','Comune0');"
          listDataProvResi=""
          listDataComuResi=""
          if IdStato="IT" then 
             listDataProvResi="absoluteProvinciaIT"
             if idProvinciaResi<>"" then 
                idProvinciaResi = getSiglaProvinciaDaProvincia(idProvinciaResi)
                if idProvinciaResi<>"" then 
                   listDataComuResi="absoluteComune" & idProvinciaResi
                end if 
             end if 
             
          end if 
          
          %>      
          <input type="text" list="<%=listDataProvResi%>" onchange="<%=onC%>" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >      
      </div>         
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>
   </div>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Comune</p>
      </div> 
      <div class="col-3">
            <%
          NameLoaded= NameLoaded & ";Comune,TE"             
          nome="Comune" & l_id
          valo=Getdiz(DizDatabase,"Comune")
          %>      
          <input type="text" list="<%=listDataComuResi%>" <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >      
      </div> 
      <div class="col-1">
         <p class="font-weight-bold">CAP</p>
      </div>
      <div class="col-1">
            <%
          NameLoaded= NameLoaded & ";Cap,TE"             
          nome="Cap" & l_id
          valo=Getdiz(DizDatabase,"Cap")
          %>      
          <input type="text"  <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >      
      </div> 
      <div class="col-2">
         <p class="font-weight-bold"> </p>
      </div>      
   </div>
   <div class="row">
      <div class="col-2">
         <p class="font-weight-bold">Indirizzo</p>
      </div> 
      <div class="col-6">
            <%
          NameLoaded= NameLoaded & ";Indirizzo,TE"             
          nome="Indirizzo" & l_id
          valo=Getdiz(DizDatabase,"Indirizzo")
          %>      
          <input type="text"  <%=readonly%> name="<%=nome%>" id="<%=nome%>" class="form-control" value="<%=valo%>" >      
      </div>    
   </div>
   
     <%if SoloLettura=false then%>
        <div class="row">
		    <div class="mx-auto">
        <%RiferimentoA=";#;;2;save;Registra; Registra;localFun('submit','0');N"%>
        <!--#include virtual="/gscVirtual/include/Anchor.asp"--> 
        <%if cdbl(idATI)>0 then 
		  RiferimentoA=";#;;2;hand;Documenti; Documenti;localDett();N" 
		%>
		   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
		<%end if %>
                        

        </div></div>
        <div class="row">
            <div class="col">
                &nbsp;
            </div>
        </div>
   <%end if %>
   
            <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
<%
  xx=createDataList("STATO","dataList_Stato","")
  xx=createDataList("COMUNE_IT","dataList_ComuneIT","")
  xx=createDataList("PROVINCIA_IT","absoluteProvinciaIT","")
  if IdProvincia<>"" then 
     xx=createDataList("COMUNE_BYSIGLAPROV_IT","absoluteComune" & IdProvincia,idProvincia)
  end if 
  if idProvinciaResi<>"" then 
     xx=createDataList("COMUNE_BYSIGLAPROV_IT","absoluteComune" & idProvinciaResi,idProvinciaResi)
  end if 
  
  
%>            


            </form>
        </div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

</body>

</html>
