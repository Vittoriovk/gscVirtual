<%
  NomePagina="DocumentoClienteUpload.asp"
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
FlagDataScadenza="S"
IdDocFir=0
if FirstLoad then 
   IdTabella          = Session("swap_IdTabella")
   IdTabellaDesc      = Session("swap_IdTabellaDesc")
   IdTabellaKeyInt    = cdbl("0" & Session("swap_IdTabellaKeyInt"))
   IdUpload           = cdbl("0" & Session("swap_IdUpload"))
   IdTabellaKeyString = Session("swap_IdTabellaKeyString")
   OperAmmesse        = Session("swap_OperAmmesse")
   FlagFileUpload     = Session("swap_FlagFileUpload")
   FlagDescEstesa     = Session("swap_FlagDescEstesa")
   FlagDataScadenza   = Session("swap_FlagDataScadenza") 
   IdAccount          = cdbl("0" & Session("swap_IdAccount"))
   IdAccountDocumento = cdbl("0" & Session("swap_IdAccountDocumento"))
   FunctionToCallIns  = Session("swap_FunctionToCallIns")     
   FunctionToCallUpd  = Session("swap_FunctionToCallUpd")     
   ProcedureToCall    = Session("swap_ProcedureToCall")   
   IdDocFir           = Session("swap_IdDocumentoToLoad")   
   TipoRife           = Session("swap_TipoRife")
   IdRife             = Session("swap_IdRife")
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
   if FlagFileUpload="" then
      FlagFileUpload     = getValueOfDic(Pagedic,"FlagFileUpload")
   end if 
   if FlagDescEstesa="" then 
      FlagDescEstesa     = getValueOfDic(Pagedic,"FlagDescEstesa")
   end if 
   if FlagDataScadenza="" then 
      FlagDataScadenza   = getValueOfDic(Pagedic,"FlagDataScadenza")
   end if 
else
   IdTabella          = getValueOfDic(Pagedic,"IdTabella")
   IdTabellaDesc      = getValueOfDic(Pagedic,"IdTabellaDesc")
   IdTabellaKeyInt    = cdbl("0" & getValueOfDic(Pagedic,"IdTabellaKeyInt"))
   IdUpload           = cdbl("0" & getValueOfDic(Pagedic,"IdUpload"))
   IdTabellaKeyString = getValueOfDic(Pagedic,"IdTabellaKeyString")
   OperAmmesse        = getValueOfDic(Pagedic,"OperAmmesse")
   PaginaReturn       = getValueOfDic(Pagedic,"PaginaReturn")
   FlagFileUpload     = getValueOfDic(Pagedic,"FlagFileUpload")
   FlagDescEstesa     = getValueOfDic(Pagedic,"FlagDescEstesa")
   FlagDataScadenza   = getValueOfDic(Pagedic,"FlagDataScadenza")
   IdAccount          = cdbl("0" & getValueOfDic(Pagedic,"IdAccount"))
   IdAccountDocumento = cdbl("0" & getValueOfDic(Pagedic,"IdAccountDocumento"))
   FunctionToCallIns  = getValueOfDic(Pagedic,"FunctionToCallIns")     
   FunctionToCallUpd  = getValueOfDic(Pagedic,"FunctionToCallUpd")    
   ProcedureToCall    = getValueOfDic(Pagedic,"ProcedureToCall")    
   IdDocFir           = getValueOfDic(Pagedic,"IdDocumentoToLoad")
   TipoRife           = getValueOfDic(Pagedic,"TipoRife")
   IdRife             = getValueOfDic(Pagedic,"IdRife")
   
end if 
response.write IdAccount
if cdbl(IdUpload)=0 and Cdbl(IdAccountDocumento)>0 then 
   IdUpload=LeggiCampo("Select IdUpload From AccountDocumento where IdAccountDocumento=" & IdAccountDocumento,"IdUpload")
   IdDocFir=LeggiCampo("Select IdDocumento From AccountDocumento where IdAccountDocumento=" & IdAccountDocumento,"IdDocumento")
   IdDocFir=Cdbl("0" & IdDocFir)
end if  
if FlagFileUpload="" then 
   FlagFileUpload  ="S"
end if 
if FlagDescEstesa  ="" then 
   FlagDescEstesa  ="S"
end if 
if FlagDataScadenza="" then 
   FlagDataScadenza="S"
end if 
'response.write "QUII" & IdTabella
 ' response.end 
if IdTabella="" or Cdbl(IdAccount)=0 then 
   response.redirect virtualpath & PaginaReturn
   response.end
end if 
ShowValidoDal=True
ShowValidoAl =True

OperAmmesse="IUD"
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
		if Cdbl(IdAccountDocumento=0) then 
			MyQ = "" 
			MyQ = MyQ & "insert into AccountDocumento (IdAccount,IdDocumento,IdUpload"
			MyQ = MyQ & ",NoteValidazione,IdTipoValidazione,TipoRife,IdRife) values ("
			MyQ = MyQ & " " & idAccount 
			MyQ = MyQ & "," & o.ValueOf("IdTipoDocumento0")
			MyQ = MyQ & "," & IdUpload
			MyQ = MyQ & "," & "'','NONRIC'"
			MyQ = MyQ & "," & "'" & TipoRife & "'," & NumForDb(IdRife)
			MyQ = MyQ & ")"
			
			ConnMsde.execute MyQ 
			IdAccountDocumento = GetTableIdentity("AccountDocumento")
		else 
			MyQ = "" 
			MyQ = MyQ & "update AccountDocumento Set IdUpload=" & IdUpload
			MyQ = MyQ & "where IdAccountDocumento=" & IdAccountDocumento
			ConnMsde.execute MyQ 
		end if 
		
		if FunctionToCallIns<>"" then 
		   FunctionToCallIns = replace(FunctionToCallIns,"$IdAccountDocumento$",IdAccountDocumento)
		   FunctionToCallIns = replace(FunctionToCallIns,"$IdUpload$",IdUpload)
           callP=VirtualPath & FunctionToCallIns
           Server.Execute(callP) 		   
		end if 
		if ProcedureToCall<>"" then 
		   ProcedureToCall = replace(ProcedureToCall,"$Action$","INS")
		   ProcedureToCall = replace(ProcedureToCall,"$IdAccountDocumento$",IdAccountDocumento)
		   ProcedureToCall = replace(ProcedureToCall,"$IdUpload$",IdUpload)
		   response.write proceduretoCall
		   ConnMsde.execute ProcedureToCall
		   
		end if 
	End If
end if

If FileCambiato=false and o.FileNameOf("FileIn0")<>"" then 
   FileCambiato=true
end if 

'response.write FileCambiato 

if Oper="UPD" and FileCambiato then 
   'se aggiorno il file lo metto in caricato 
   MyQ = "" 
   MyQ = MyQ & "update AccountDocumento Set IdTipoValidazione='NONRIC'"
   MyQ = MyQ & "where IdAccountDocumento=" & IdAccountDocumento
   ConnMsde.execute MyQ 
   
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


if Oper="UPD" then 
    Session("TimeStamp")=TimePage
	KK=o.ValueOf("ItemToRemove")
    ValidoDal = "0" & DataStringa(o.ValueOf("ValidoDal0"))
    if isnumeric(ValidoDal)=false then 
       ValidoDal=0
    else
       ValidoDal=Cdbl(ValidoDal)
    end if 
    ValidoAl  = "0" & DataStringa(o.ValueOf("ValidoAl0"))
    if isnumeric(ValidoAl)=false then 
       ValidoAl=20991231
    else
       ValidoAl=Cdbl(ValidoAl)
    end if 
	if ValidoAl=0 then 
	   ValidoAl=20991231
	end if 
	   
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
	'response.write MyQ
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else
	    FlagUpdLista=true
		if FunctionToCallUpd<>"" then 
		   FunctionToCallUpd = replace(FunctionToCallUpd,"$IdAccountDocumento$",IdAccountDocumento)
		   FunctionToCallUpd = replace(FunctionToCallUpd,"$IdUpload$",IdUpload)
           callP=VirtualPath & FunctionToCallUpd
           Server.Execute(callP) 		   
		end if 
		if ProcedureToCall<>"" then 
		   ProcedureToCall = replace(ProcedureToCall,"$Action$","UPD")
		   ProcedureToCall = replace(ProcedureToCall,"$IdAccountDocumento$",IdAccountDocumento)
		   ProcedureToCall = replace(ProcedureToCall,"$IdUpload$",IdUpload)
		   ConnMsde.execute ProcedureToCall
		end if 
		response.redirect VirtualPath & PaginaReturn
		response.end
	End If
	DescIn=""
End if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdTabella"          ,IdTabella)
  xx=setValueOfDic(Pagedic,"IdTabellaDesc"      ,IdTabellaDesc)
  xx=setValueOfDic(Pagedic,"IdTabellaKeyInt"    ,IdTabellaKeyInt)
  xx=setValueOfDic(Pagedic,"IdUpload"           ,IdUpload)
  xx=setValueOfDic(Pagedic,"IdTabellaKeyString" ,IdTabellaKeyString)
  xx=setValueOfDic(Pagedic,"OperAmmesse"        ,OperAmmesse)
  xx=setValueOfDic(Pagedic,"PaginaReturn"       ,PaginaReturn)
  xx=setValueOfDic(Pagedic,"FlagFileUpload"     ,FlagFileUpload)
  xx=setValueOfDic(Pagedic,"FlagDescEstesa"     ,FlagDescEstesa)
  xx=setValueOfDic(Pagedic,"FlagDataScadenza"   ,FlagDataScadenza)
  xx=setValueOfDic(Pagedic,"IdAccount"          ,IdAccount)
  xx=setValueOfDic(Pagedic,"IdAccountDocumento" ,IdAccountDocumento)
  xx=setValueOfDic(Pagedic,"FunctionToCallIns"  ,FunctionToCallIns)
  xx=setValueOfDic(Pagedic,"FunctionToCallUpd"  ,FunctionToCallUpd)
  xx=setValueOfDic(Pagedic,"ProcedureToCall"    ,ProcedureToCall)
  xx=setValueOfDic(Pagedic,"IdDocumentoToLoad"  ,IdDocFir)
  xx=setValueOfDic(Pagedic,"TipoRife"           ,TipoRife)
  xx=setValueOfDic(Pagedic,"IdRife"             ,IdRife)
  xx=setCurrent(NomePagina,livelloPagina) 

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
			<form name="Fdati" Action="<%=NomePagina%>" method="post" enctype="multipart/form-data">
			<div class="row">
			<%RiferimentoA="col-1  text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h3><b><%=DescElenco%></b> </h3>
				</div>
			</div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
   <%
  
   LeggiDati=false
   if Cdbl(IdUpload)>0 then
      err.clear 
      LeggiDati=true
      Set RsRec = Server.CreateObject("ADODB.Recordset")
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
   if FlagDescEstesa="S" then 
      NameLoaded= NameLoaded & ";DescEstesa,TE"  
   end if 
   If ShowValidoDal then 
      NameLoaded= NameLoaded & ";ValidoDal,DTO"  
   end if 
   if LeggiDati=false then 
      NameLoaded= NameLoaded & ";FileIn,TE"
   end if  
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
   if Cdbl(ValoC)=0 then 
      ValoC = cdbl("0" & Session("swap_IdDocumento"))
   end if 
   descAltro=""
   if Cdbl(ValoC)>0 then 
      descAltro = LeggiCampo("SELECT * From Documento Where IdDocumento=" & ValoC,"DescDocumento")
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
   <%if LeggiDati=true then%>
   
   <div class="row" >
	   <div class="col-2">
		  <p class="font-weight-bold">File Caricato</p>
	   </div>
	   <div class = "col-8">
	   <input value="<%=RsRec("NomeDocumento")%>" type="text" READONLY class="form-control"  >
	   
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
