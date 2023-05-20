<%
  NomePagina="DocumentoClienteUploadGestione.asp"
  titolo="Menu - Dashboard"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/common/clsupload.asp"-->
<%
  livelloPagina="00"
  set o = new clsUpload
    
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
function localDelFile(aHref)
{
var xx;
   xx=ImpostaValoreDi("FilePrecedente0","");
   $('#' + aHref).empty();
   $('#' + aHref).contents().unwrap();
   $('#' + aHref + "_1").empty();
   $('#' + aHref + "_1").contents().unwrap();
}

function cambiaId()
{
var xx;

	xx=ValoreDi("IdTipoDocumento0");
	if (xx=="-1") {
		xx=ImpostaValoreDi("DescBreve0","");
		xx=ImpostaValoreDi("DescEstesa0","");
	}
	else {
	    yy=$("#IdTipoDocumento0 option:selected" ).text();
		xx=ImpostaValoreDi("DescBreve0",yy);
		xx=ImpostaValoreDi("DescEstesa0",yy);
	}
}
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
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<%
PageSize=0
CPag=1 

Oper = o.ValueOf("Oper")
Oper = ucase(Oper)
'SERVE  A GESTIRE UN EVENTUALE REFRESH DELLA PAGINA 
TimeStamp = Dtos() & TimeTos()
TimePage = Request("TimePage")

If (Oper="INS" or OPER="UPD" or OPER=ucase("RemoveItem")) and Session("TimeStamp")<>"" then  
	If Session("TimeStamp") = TimePage Then
		Oper=" "
	End If
end if 
%>
<%
IdTabella=""
idTabellaDesc=""
IdUpload=0
IdAccount=0
IdAccountDocumento=0
IdTabellaKeyInt=0
IdTabellaKeyString=""
FlagFileUpload  ="S"
FlagDescEstesa  ="S"
IdTipoValidazione="DAVALI"
NoteValidazione  =""
IdDocumento=0 

IdDocFir=0
if FirstLoad then 
   IdTabella          = Session("swap_IdTabella")
   IdTabellaDesc      = Session("swap_IdTabellaDesc")
   IdTabellaKeyInt    = cdbl("0" & Session("swap_IdTabellaKeyInt"))
   IdUpload           = cdbl("0" & Session("swap_IdUpload"))
   IdTabellaKeyString = Session("swap_IdTabellaKeyString")
   FlagFileUpload     = Session("swap_FlagFileUpload")
   IdRichiesta        = cdbl("0" & Session("swap_IdRichiesta"))
   IdDocumento        = cdbl("0" & Session("swap_IdDocumento"))
   TipoRife           = Session("swap_TipoRife")
   IdRife             = cdbl("0" & Session("swap_IdRife"))
   if TipoRife="" then 
      IdRife=0
   end if 
   
   if idTabella="" then 
      IdTabella = getValueOfDic(Pagedic,"IdTabella")
   end if 
   if Cdbl(IdTabellaKeyInt)=0 then 
      IdTabellaKeyInt = cdbl("0" & getValueOfDic(Pagedic,"IdTabellaKeyInt"))
   end if
   if Cdbl(IdUpload)=0 then 
      IdUpload = cdbl("0" & getValueOfDic(Pagedic,"IdUpload"))
   end if
   if IdTabellaKeyString="" then 
      IdTabellaKeyString = getValueOfDic(Pagedic,"IdTabellaKeyString")
   end if    
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
else
   IdTabella          = getValueOfDic(Pagedic,"IdTabella")
   IdTabellaDesc      = getValueOfDic(Pagedic,"IdTabellaDesc")
   IdTabellaKeyInt    = cdbl("0" & getValueOfDic(Pagedic,"IdTabellaKeyInt"))
   IdUpload           = cdbl("0" & getValueOfDic(Pagedic,"IdUpload"))
   IdTabellaKeyString = getValueOfDic(Pagedic,"IdTabellaKeyString")
   PaginaReturn       = getValueOfDic(Pagedic,"PaginaReturn")
   IdAccount          = cdbl("0" & getValueOfDic(Pagedic,"IdAccount"))
   IdRichiesta        = cdbl("0" & getValueOfDic(Pagedic,"IdRichiesta"))
   IdDocumento        = cdbl("0" & getValueOfDic(Pagedic,"IdDocumento"))
   TipoRife           = getValueOfDic(Pagedic,"TipoRife")
   IdRife             = cdbl("0" & getValueOfDic(Pagedic,"IdRife"))
end if 
'IdRichiesta = IdAffidamentoRichiestaComp
'IdDocumento = documento da caricare 

Set RsRec = Server.CreateObject("ADODB.Recordset")
qDati = ""
qDati = qDati & " Select * From AffidamentoRichiestaCompDoc "
qDati = qDati & " where IdAffidamentoRichiestaComp =" & IdRichiesta
qDati = qDati & " And   IdDocumento = " & IdDocumento

'response.write TipoRife & IdRife
'response.end 

RsRec.CursorLocation = 3
RsRec.Open qDati, ConnMsde 

IdDocFir           = IdDocumento
IdAccountDocumento = 0
FlagDataScadenza   = 1
FlagRichiesto      = 0
If Err.number<>0 then	
   IdDocFir = 0
elseIf RsRec.EOF then	
   RsRec.close 
else 
   
   IdAccountDocumento = RsRec("IdAccountDocumento")
   FlagRichiesto      = RsRec("FlagObbligatorio")
   FlagScadenza       = RsRec("FlagScadenza")
   'response.write " trovato " & IdAccountDocumento & " "
   RsRec.close 
End if

if IdDocFir = 0 then 
   response.redirect virtualpath & PaginaReturn
   response.end
end if 

if Cdbl(IdAccountDocumento) > 0 then 
   qDati = "Select * From AccountDocumento where IdAccountDocumento =" & IdAccountDocumento
   RsRec.CursorLocation = 3
   RsRec.Open qDati, ConnMsde 
   IdTipoValidazione = RsRec("IdTipoValidazione")
   NoteValidazione   = RsRec("NoteValidazione")
   IdUpload          = RsRec("IdUpload")
   RsRec.close 
end if 

if cdbl(IdAccount)=0 then 
   qDati = ""
   qDati = qDati & " Select * From AffidamentoRichiestaComp a, AffidamentoRichiesta B  "
   qDati = qDati & " where A.IdAffidamentoRichiestaComp = " & IdRichiesta
   qDati = qDati & " and   A.IdAffidamentoRichiesta = b.IdAffidamentoRichiesta "
   'response.write qDati
   IdAccount = "0" & LeggiCampo(qDati,"IdAccountCliente")
end if 

if FlagFileUpload="" then 
   FlagFileUpload  ="S"
end if 

'response.write "QUII" & IdTabella & " " & IdAccount & " " & IdRichiesta
'response.end 
if IdTabella="" or Cdbl(IdAccount)=0 then 
   response.redirect virtualpath & PaginaReturn
   response.end
end if 
ShowValidoDal=True
ShowValidoAl =True

OperAmmesse="IUD"
FilePrecedente=o.ValueOf("FilePrecedente0")
FileAttuale   =o.FileNameOf("FileIn0")
on error resume next
If Oper="UPDATE" then
   if Cdbl(IdUpload)=0 then 
      Oper="INS"
   else
      Oper="UPD"
   end if 
end if 

DescElenco = IdTabellaDesc
if DescElenco="" then 
   DescElenco="Documento cliente : "
   if Cdbl(IdUpload)=0 then 
      DescElenco=DescElenco & " Inserimento "
   else
      DescElenco=DescElenco & " Aggiornamento "
   end if 
end if 

FileCambiato=false
if Oper="INS" then 
    Session("TimeStamp")=TimePage
	KK=0
	MyQ = "" 
	MyQ = MyQ & " insert into Upload (IdTabella,IdTabellaKeyInt,IdTabellaKeyString,DataUpload"
	MyQ = MyQ & " ,TimeUpload,IdTipoDocumento,DescBreve,DescEstesa,NomeDocumento,PathDocumento,ValidoDal,ValidoAl) "
	MyQ = MyQ & " values ("
	MyQ = MyQ & " '" & Apici(IdTabella) & "'"
	MyQ = MyQ & ", " & IdTabellaKeyInt
	MyQ = MyQ & ",'" & Apici(IdTabellaKeyString) & "'"
	MyQ = MyQ & ",0,0,'','','','','',0,20991231)" 
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else
	    FileCambiato=true
	    IdUpload=GetTableIdentity("Upload")
		Oper="UPD"
	End If
end if

If FileCambiato=false and o.FileNameOf("FileIn0")<>"" then 
   FileCambiato=true
end if 

'il file non è presente 
FileAssente = false 
if Oper="UPD" and o.FileNameOf("FileIn0")="" and FilePrecedente="" then 
   FileAssente=true 
   qUpd = ""
   qUpd = qUpd & " Update Upload Set "
   qUpd = qUpd & " NomeDocumento=''"
   qUpd = qUpd & ",PathDocumento=''"
   qUpd = qUpd & " where IdUpload=" & IdUpload   
   'response.write qUpd
   FileCambiato = false
   ConnMsde.execute qUpd 
end if 

if Oper="UPD" and FileCambiato then 
   NomeFilFull = o.FileNameOf("FileIn0")
   sFileSplit = split(NomeFilFull, "\")
   sFile = sFileSplit(Ubound(sFileSplit))

   sFileWrite = "CX" & IdUpload & "_" & Year(Now()) & Month(Now()) & Day(Now()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now()) &  "_" & sFile
  
   o.FileInputName = "FileIn0"
   o.FileFullPath = PathBaseUpload  & sFileWrite
   o.save
   if o.Error <> ""  then
       MsgErrore= "Caricamento Fallito: " & o.Error & o.FileFullPath
   elseif err.number<>0 then
       MsgErrore= "Caricamento Errore : " & Err.Description
   else
       qUpd = ""
	   qUpd = qUpd & " Update Upload Set "
	   qUpd = qUpd & " NomeDocumento='" & apici(sFile) & "'"
	   qUpd = qUpd & ",PathDocumento='" & apici(sFileWrite) & "'"
	   qUpd = qUpd & " where IdUpload=" & IdUpload   
       ConnMsde.execute qUpd 
   end if 
end if 

'se aggiorno Upload e non esiste il cassetto lo inserisco 
if Oper="UPD" then 
   if Cdbl(IdAccountDocumento=0) then 
      MyQ = "" 
      MyQ = MyQ & "insert into AccountDocumento (IdAccount,IdDocumento,IdUpload"
      MyQ = MyQ & ",NoteValidazione,IdTipoValidazione,TipoRife,IdRife) values ("
      MyQ = MyQ & "  " & idAccount 
      MyQ = MyQ & ", " & o.ValueOf("IdTipoDocumento0")
      MyQ = MyQ & ", " & IdUpload
      MyQ = MyQ & ", " & "'','NONRIC'"
	  MyQ = MyQ & ",'" & Apici(TipoRife) & "'," & IdRife
	  MyQ = MyQ & ")"
      ConnMsde.execute MyQ 
      IdAccountDocumento = GetTableIdentity("AccountDocumento")
   elseif FileCambiato then  
      MyQ = "" 
      MyQ = MyQ & "update AccountDocumento Set IdUpload=" & IdUpload
	  MyQ = MyQ & ",IdTipoValidazione = 'NONRIC'"
      MyQ = MyQ & "where IdAccountDocumento=" & IdAccountDocumento
      ConnMsde.execute MyQ 
   end if 
   'response.write MyQ
end if 

'aggiorno gli attributi
if Oper="UPD" then 
    Session("TimeStamp")=TimePage
	KK=o.ValueOf("ItemToRemove")
    ValidoDal = "0" & DataStringa(o.ValueOf("ValidoDal0"))
    if isnumeric(ValidoDal)=false or FileAssente then 
       ValidoDal=0
    else
       ValidoDal=Cdbl(ValidoDal)
    end if 
    ValidoAl  = "0" & DataStringa(o.ValueOf("ValidoAl0"))
    if isnumeric(ValidoAl)=false or FileAssente then 
       ValidoAl=0
    else
       ValidoAl=Cdbl(ValidoAl)
    end if 
	'if ValidoAl=0 then 
	'   ValidoAl=20991231
	'end if 
	
	flagObbligatorio=1
	if o.ValueOf("FlagObbl0")<>"S" then 
	   FlagObbligatorio=0
    end if 
	
	MyQ = ""
    MyQ = MyQ & " Update AffidamentoRichiestaCompDoc set "
	MyQ = MyQ & " IdAccountDocumento = " & IdAccountDocumento
	if IsBackOffice() then 
	   MyQ = MyQ & ",FlagObbligatorio = " & flagObbligatorio
	end if 
	MyQ = MyQ & " where IdAffidamentoRichiestaComp =" & IdRichiesta
	MyQ = MyQ & " and IdDocumento=" & IdDocFir
	MyQ = MyQ & " and TipoRife='" & TipoRife & "'"
	MyQ = MyQ & " and IdRife=" & IdRife 
	ConnMsde.execute MyQ 
	
	MyQ = "" 
	MyQ = MyQ & " update Upload set "
	MyQ = MyQ & " DataUpload = " & Dtos()
	MyQ = MyQ & ",TimeUpload = " & TimeToS()
	MyQ = MyQ & ",IdTipoDocumento = " & o.ValueOf("IdTipoDocumento0")
	MyQ = MyQ & ",DescBreve='"        & apici(o.ValueOf("DescBreve0")) & "'"
	MyQ = MyQ & ",DescEstesa='"       & apici(o.ValueOf("descEstesa0")) & "'"
	MyQ = MyQ & ",ValidoDal=" & ValidoDal
	MyQ = MyQ & ",ValidoAl="  & ValidoAl
	MyQ = MyQ & " where IdUpload = " & Idupload 
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else
	    'aggiorno dettaglio 
		IdTipoValidazione = o.ValueOf("IdTipoValidazione0")
		NoteValidazione   = o.ValueOf("NoteValidazione0")
        if FileAssente and cdbl(flagObbligatorio)=1 then 
		   IdTipoValidazione="NONVAL"
		end if 
		MyQ = "" 
		MyQ = MyQ & " update AccountDocumento set "
		MyQ = MyQ & " IdTipoValidazione = '" & apici(IdTipoValidazione) & "'"
		MyQ = MyQ & ",NoteValidazione = '" & apici(NoteValidazione)     & "'"
		MyQ = MyQ & " where IdAccountDocumento = " & IdAccountDocumento 
		if Session("LoginTipoUtente")=ucase("BackO") then 
		   ConnMsde.execute MyQ 
		   response.write MyQ 
		end if 
		
		'response.end 
		
		
	    FlagUpdLista=true
		response.redirect VirtualPath & PaginaReturn
		response.end
	End If
	DescIn=""
End if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdTabella"          ,IdTabella)
  xx=setValueOfDic(Pagedic,"IdTabellaDesc"      ,IdTabellaDesc)
  xx=setValueOfDic(Pagedic,"IdTabellaKeyInt"    ,IdTabellaKeyInt)
  xx=setValueOfDic(Pagedic,"IdTabellaKeyString" ,IdTabellaKeyString)
  xx=setValueOfDic(Pagedic,"PaginaReturn"       ,PaginaReturn)
  xx=setValueOfDic(Pagedic,"IdAccount"          ,IdAccount)
  xx=setValueOfDic(Pagedic,"IdRichiesta"        ,IdRichiesta)
  xx=setValueOfDic(Pagedic,"IdDocumento"        ,IdDocumento)
  xx=setValueOfDic(Pagedic,"TipoRife"           ,TipoRife)
  xx=setValueOfDic(Pagedic,"IdRife"             ,IdRife)
  
  xx=setCurrent(NomePagina,livelloPagina) 

%>


<%   
  'xx=DumpDic(SessionDic,NomePagina)
%>
<div class="d-flex" id="wrapper">
	<%
	  TitoloNavigazione="Gestione Documento"
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
			<form name="Fdati" Action="<%=NomePagina%>" method="post" enctype="multipart/form-data">
			<div class="row">
			<%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h3><b><%=DescElenco%></b> </h3>
				</div>
			</div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
   <%
   DescCliente   = LeggiCampo("Select * from Account Where idAccount=" & IdAccount,"Nominativo" )
   IdCompagnia   = LeggiCampo("Select * From AffidamentoRichiestaComp where IdAffidamentoRichiesta =" & IdRichiesta,"IdCompagnia")
   DescCompagnia = LeggiCampo("Select * From Compagnia where IdCompagnia =" & IdCompagnia,"DescCompagnia")
   %>
            <div class="row">
               <div class="col-2">
                   <p class="font-weight-bold">Utente</p>
               </div> 
	           <div class="col-8">
	              <input type="text" readonly class="form-control" value="<%=DescCliente%>" >	  
	           </div>
	           <div class="col-2">
	               
	           </div>
            </div> 	
            <%if not isCliente() and DescCompagnia<>"" then %>
            <div class="row">
               <div class="col-2">
                   <p class="font-weight-bold">Compagnia</p>
               </div> 
	           <div class="col-8">
	              <input type="text" readonly class="form-control" value="<%=DescCompagnia%>" >	  
	           </div>
	           <div class="col-2">
	               
	           </div>
            </div> 
			<%end if %>
			
   <%
  
   LeggiDati=false
   if Cdbl(IdUpload)>0 then
      err.clear 
      LeggiDati=true
      
      MySql = "" 
      MySql = MySql & " Select * from Upload Where IdUpload = " & IdUpload

      RsRec.CursorLocation = 3
      RsRec.Open MySql, ConnMsde 

      If Err.number<>0 then	
       	 LeggiDati=false
      elseIf RsRec.EOF then	
         LeggiDati=false
		 RsRec.close 
      End if
   end if   
 
   NameLoaded= ""
   NameLoaded= NameLoaded & "IdTipoDocumento,LI" 
   NameLoaded= NameLoaded & ";DescBreve,TE"  
   NameLoaded= NameLoaded & ";DescEstesa,TE"  
   NameLoaded= NameLoaded & ";ValidoDal,DT"  
   NameLoaded= NameLoaded & ";ValidoAl,DT" 
   DescLoaded="0"
   
   l_Id = "0"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   if LeggiDati then 
       ValoC = RsRec("IdTipoDocumento")   
   else
       ValoC = IdDocFir
   end if  
   descAltro=""
   if Cdbl(ValoC)>0 then 
      descAltro = LeggiCampo("SELECT * From Documento Where IdDocumento=" & ValoC,"DescDocumento")
	  if TipoRife="COOB" and cdbl(IdRife)>0 then 
	     qs = "Select Ragsoc as descInfo from AccountCoobbligato Where IdAccountCoobbligato=" & IdRife 
		 descAltro = trim(descAltro & " " & LeggiCampo(qs,"descInfo"))
	  end if 
   end if    
   ao_lbd = "Tipo Documento"                       'descrizione label 
   ao_nid = "IdTipoDocumento" & l_Id              'nome ed id
   ao_val = ValoC 'valore di default
   
   if Cdbl(ValoC)>0 then
      ao_Tex = "SELECT * From Documento Where IdDocumento=" & ValoC
	  ao_Att = "0"
   else
      ao_Tex = "SELECT * From Documento Where IdDocumentoInterno='' order By DescDocumento"
	  ao_Att = "1"
   end if 
   
   'response.write ao_tex
   ao_ids = "IdDocumento"				  'valore della select 
   ao_des = "DescDocumento"              'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = "cambiaId()"              'azzero evento
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			      
   <%
   if LeggiDati then 
       ValoC = RsRec("DescBreve")   
   else
       ValoC = ""
   end if
   if ValoC="" then 
      ValoC=descAltro
   end if 
   ao_lbd = "Descrizione Breve"       'descrizione label 
   ao_nid = "DescBreve" & l_Id            'nome ed id
   ao_val = "|value=" & ValoC       'valore di default
   ao_Plh = "|placeholder=Descr.Breve"              'placeholder
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		

   <%if FlagDescEstesa="S" then %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			      
   <%
   if LeggiDati then 
       ValoC = RsRec("DescEstesa")   
   else
      ValoC = Request("DescEstesa" & l_Id) 
   end if
   if ValoC="" then 
      ValoC=descAltro
   end if 
   
   ao_lbd = "Descrizione estesa"       'descrizione label 
   ao_nid = "DescEstesa" & l_Id            'nome ed id
   ao_val = "|value=" & ValoC       'valore di default
   ao_Plh = "|placeholder=Descr.Estesa"              'placeholder
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		   
   <%else%>
      <input type="hidden" name="DescEstesa<%=l_Id%>" id="DescEstesa<%=l_Id%>" value="">
   <%end if %>
   
   
   <% If ShowValidoDal then%>
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ValoC=""
   if LeggiDati then 
      ValoC = RsRec("ValidoDal")   
   else
      ValoC = Request("ValidoDal" & l_Id) 
   end if   
   ao_lbd = "Valido Dal"         'descrizione label
   ao_3ls = "col-6"                       'size terzo elemento	
   ao_div = "col-4"	   
   ao_nid = "ValidoDal" & l_Id            'nome ed id
   ao_val = ""       'valore di default
   if len(ValoC)<>8 then 
      ValoC=DtoS()
   end if 
   ValoC=Stod(ValoC)
   ao_val = ValoC 
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddDate.asp"--> 
   
   
   <% end if  %>
   <% If ShowValidoAl then%>
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ValoC=""
   if LeggiDati then 
      ValoC = RsRec("ValidoAl")   
   else
      ValoC = Request("ValidoAl" & l_Id) 
   end if 
 
   ao_lbd = "Valido Al"         'descrizione label
   ao_3ls = "col-6"                       'size terzo elemento	
   ao_div = "col-4"	   
   ao_nid = "ValidoAl" & l_Id            'nome ed id
   ao_val = ""       'valore di default
   if len(ValoC)<>8 then 
      ValoC="20991231"
   end if 
   ValoC=Stod(ValoC)
   ao_val = ValoC 
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddDate.asp"--> 
   
   
   <% end if  %>
   <%if LeggiDati=true and RsRec("PathDocumento")<>""  then%>
   
   <div class="row" >
	   <div class="col-2">
		  <p class="font-weight-bold">File Caricato</p>
	   </div>
	   <div class = "col-6">
	   <input value="<%=RsRec("NomeDocumento")%>" type="text" Id="FilePrecedente0" Name="FilePrecedente0" READONLY class="form-control"  >
	   
	   
	   </div>
	   <div class="col-2">
				<%
				IdlinkForDownload=""
				Linkdocumento=RsRec("PathDocumento")%>
				<!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->
          <%if IdlinkForDownload<>"" then %>
		  <a id="<%=IdlinkForDownload & "_1"%>" href="#!" title="Cancella" onclick="localDelFile('<%=IdlinkForDownload%>');">
		  <i class="fa fa-2x fa-trash"></i></a>
		  <%end if %>
	   </div>   
   <div class="col-2">
      <p class="font-weight-bold"> </p>
   </div>  
   </div> 

   <%end if %>
   
   
<div class="row" >

   <div class="col-2">
      <p class="font-weight-bold">File Da Caricare</p>
   </div>
   <div class = "col-8">
   
  <div class="custom-file">
    <input type="file" id="FileIn0" name="FileIn0"  aria-describedby="inputGroupFileAddon01">
   </div>
     </div>
   <div class="col-2">
      <p class="font-weight-bold"> </p>
   </div>

</div> 
<% if Session("LoginTipoUtente")=ucase("BackO") then %>

<div class="row" >

   <div class="col-2">
      <p class="font-weight-bold">Documento Obbligatorio</p>
   </div>
   <%
   selezionato="" 
   if FlagRichiesto = 1 then 
      selezionato = " checked "   
   end if 
   %>
   <div class = "col-8">
   	<input id="FlagObbl<%=l_Id%>" <%=selezionato%> name="FlagObbl<%=l_Id%>" 
	type="checkbox" value = "S" class="big-checkbox">

   </div>
   <div class="col-2">
      <p class="font-weight-bold"> </p>
   </div>

</div>         

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   ValoC = IdTipoValidazione

   descAltro=""
   ao_lbd = "Stato Documento"                       'descrizione label 
   ao_nid = "IdTipoValidazione" & l_Id              'nome ed id
   ao_val = ValoC 'valore di default
   
   ao_Tex = "SELECT * From TipoValidazione order By DescTipoValidazione"
   ao_Att = "0"
      
   'response.write ao_tex
   ao_ids = "IdTipoValidazione"		  'valore della select 
   ao_des = "DescTipoValidazione"     'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""              'azzero evento
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			      
   <%
   ValoC = NoteValidazione  
   ao_lbd = "Note Documento"       'descrizione label 
   ao_nid = "NoteValidazione" & l_Id            'nome ed id
   ao_val = "|value=" & ValoC       'valore di default
   ao_Plh = "|placeholder=Note Di Validazione"              'placeholder
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->

<%end if %>
   
   <%if SoloLettura=false then%>
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;localSubmit('submit','0');S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div></div>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
		</div>
   <%end if %>
   
   
			
			
			
			<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
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
