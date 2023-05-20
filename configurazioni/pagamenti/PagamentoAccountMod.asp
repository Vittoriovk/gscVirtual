<%
  NomePagina="PagamentoAccountMod.asp"
  titolo="Modifica Pagamento per cliente"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/modelli/functionPagamenti.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionEvento.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionUpload.asp"-->
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

	ImpostaValoreDi("Oper",Op);
	document.Fdati.submit();
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
	if (Op=="update" || Op=="prosegui")
	   xx=ElaboraControlli();
   	
 	if (xx==false)
	   return false;
		
	ImpostaValoreDi("Oper",op);
	document.Fdati.submit(); 
}
</script>

<%

  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdAccount = 0
  IdCliente = 0
  IdMovEco  = 0
  if FirstLoad then 
     IdCliente     = cdbl("0" & getCurrentValueFor("IdCliente"))
     IdAccount     = cdbl("0" & getCurrentValueFor("IdAccount"))
	 IdMovEco      = cdbl("0" & getCurrentValueFor("IdMovEco"))
	 IdTipoCredito = getCurrentValueFor("IdTipoCredito")
     DescCliente   = getCurrentValueFor("DescCliente")
     OperTabella   = getCurrentValueFor("OperTabella")
     PaginaReturn  = getCurrentValueFor("PaginaReturn") 
  else
     IdCliente       = cdbl("0" & getValueOfDic(Pagedic,"IdCliente"))
     IdAccount       = cdbl("0" & getValueOfDic(Pagedic,"IdAccount"))
	 IdMovEco        = cdbl("0" & getValueOfDic(Pagedic,"IdMovEco"))
	 IdTipoCredito   = getValueOfDic(Pagedic,"IdTipoCredito")
	 DescTipoCredito = getValueOfDic(Pagedic,"DescTipoCredito")
     DescCliente     = getValueOfDic(Pagedic,"DescCliente")     
     OperTabella     = getValueOfDic(Pagedic,"OperTabella")
     PaginaReturn    = getValueOfDic(Pagedic,"PaginaReturn")
   end if   
  
   IdCliente = cdbl(IdCliente)
   IdAccount = cdbl(IdAccount)
   IdMovEco  = cdbl(IdMovEco)

   if DescCliente="" then 
      DescCliente=LeggiCampo("select * from Cliente Where IdCliente=" & IdCliente,"Denominazione")
   end if  
   if IdAccount=0 then 
      IdAccount  =LeggiCampo("select * from Cliente Where IdCliente=" & IdCliente,"IdAccount")   
   end if 
   
   if IdTipoCredito="" then 
      response.redirect RitornaA(PaginaReturn)
	  response.end 
   end if 
   
   if IdMovEco=0 then 
      OperAmmesse = "IU"
   end if 
   
  'inserisco account 
   flagProsegui=false 
   flagEsci    =false
   if Oper=ucase("prosegui") then 
      flagProsegui=true
      Oper=ucase("update")
   end if 
   
   if Oper=ucase("update") then 
      Ritorna=false 
	  OperAmmesse="U"
      Session("TimeStamp")=TimePage
      MsgErrore=""
      ImptMovEco=Request("ImptMovEco0")
      DescMovEco=Request("DescMovEco0")
      
	  MyQ=""
      if Cdbl(IdMovEco)=0 then 
	     IdStatoCredito="COMP"
		 if flagProsegui=true then 
		    IdStatoCredito="LAVO"
		 end if 
         MyQ = MyQ & " insert into AccountMovEco"
         MyQ = MyQ & "(IdAccount,IdAccountGestore,IdTipoCredito,IdStatoCredito,DescStatoCredito"
		 MyQ = MyQ & ",DescMovEco,DataMovEco,TimeMovEco,ImptMovEco"
		 MyQ = MyQ & ",FlagStorico,NoteStorico,IdUpload,SistemaSorgente,SegnoSistema) values "
         MyQ = MyQ & "(" & IdAccount & ",0,'" & IdTipoCredito & "','" & IdStatoCredito & "',''"
		 MyQ = MyQ & ",'" & apici(DescMovEco) & "',"  & Dtos() & "," & TimeToS() & "," & NumForDb(ImptMovEco)
         MyQ = MyQ & ",0,'',0,'RIC',1)"  
		 ConnMsde.execute MyQ 
		 'response.write MyQ & Err.description 
		 if Err.Number=0 then 
		    IdMovEco=GetTableIdentity("AccountMovEco")
			xx=UpdatePagaAccount(IdAccount,IdTipoCredito)
			if IdStatoCredito="LAVO" then 
			   IdProdotto=0
			   'response.write "ecco"
               XX=createEvento(IdTipoCredito,"CARI",Session("LoginIdAccount"),DescMovEco,"AccountMovEco","IdAccountMovEco=" & IdMovEco,true,IdProdotto)		 
            end if 
			flagEsci=true 
		 end if 
      else 
	     err.clear 
         MyQ = MyQ & " update AccountMovEco set "
         MyQ = MyQ & " DescMovEco ='" & apici(DescMovEco) & "'"  
		 IdStatoCredito=""
		 if flagProsegui=true then 
		    IdStatoCredito="LAVO"
            MyQ = MyQ & ",IdStatoCredito ='LAVO'"   
		 end if 
         MyQ = MyQ & ",ImptMovEco = " & NumForDb(ImptMovEco) 
         MyQ = MyQ & " Where IdAccountMovEco = " & IdMovEco 
         MyQ = MyQ & " and   IdAccount=" & IdAccount
		 ConnMsde.execute MyQ 
		 'response.write MyQ
		 if err.number=0 then 
		    xx=UpdatePagaAccount(IdAccount,IdTipoCredito)
			if IdStatoCredito="LAVO" then 
			   IdProdotto=0
               XX=createEvento(IdTipoCredito,"CARI",Session("LoginIdAccount"),DescMovEco,"AccountMovEco","IdAccountMovEco=" & IdMovEco,true,IdProdotto)		 
            end if 			
		    flagEsci=true
		 end if 
      end if 
	  if flagEsci = true then 
	     UpdateUploadDoc    = "S"
	     IdTabella          = "AccountMovEco"
         IdTabellaKeyInt    = IdMovEco
         IdTabellaKeyString = ""
	     %>
	     <!--#include virtual="/gscVirtual/utility/modalUpload/updateUploadDoc.asp"-->
	     <%
		 flagEsci = false 
	  end if 
	  
   end if 
   'response.write flagEsci
   if flagEsci=true then 
      response.redirect VirtualPath & PaginaReturn
	  response.end 
   end if 
   
   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
	 
   'recupero i dati 
   
   if cdbl(IdMovEco)>0 then
      OperAmmesse="U"
	  MySql = ""
	  MySql = MySql & " Select * From  AccountMovEco "
	  MySql = MySql & " Where IdAccountMovEco=" & IdMovEco
	  xx=GetInfoRecordset(DizDatabase,MySql)
	  if Getdiz(DizDatabase,"FlagValidato")="1" then 
	     OperAmmesse=""
      end if 
   else
     OperAmmesse="IU"   
   end if 
     
   DescPageOper="Aggiornamento"
   If cdbl(IdMovEco)=0 then 
      DescPageOper = "Inserimento"
   end if
   'response.write OperAmmesse
  'registro i dati della pagina 
  if descTipoCredito="" then 
     DescTipoCredito=LeggiCampo("select * from TipoCredito where IdTipoCredito='" & IdTipoCredito & "'","DescTipoCredito")
  end if 
  
  xx=setValueOfDic(Pagedic,"IdCliente"      ,IdCliente)
  xx=setValueOfDic(Pagedic,"DescCliente"    ,DescCliente)
  xx=setValueOfDic(Pagedic,"IdAccount"      ,IdAccount)
  xx=setValueOfDic(Pagedic,"IdMovEco"       ,IdMovEco)  
  xx=setValueOfDic(Pagedic,"IdTipoCredito"  ,IdTipoCredito)
  xx=setValueOfDic(Pagedic,"DescTipoCredito",DescTipoCredito)
  xx=setValueOfDic(Pagedic,"OperAmmesse"    ,OperAmmesse)
  xx=setValueOfDic(Pagedic,"PaginaReturn"   ,PaginaReturn)
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
				<div class="col-11"><h3>Movimento <%=DescTipoCredito%> per Cliente :</b> <%=DescCliente%> </h3>
				</div>
			</div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

    <%
      l_Id = "0"
	  err.clear
      ReadOnly=""
	  SoloLettura=false
      if instr(OperAmmesse,"U")=0 or (instr(OperAmmesse,"I")>0 and cdbl("0" & IdMovEco)>0) then 
         SoloLettura=true
         ReadOnly=" readonly "
      end if 
      NameLoaded= ""
	  NameLoaded= NameLoaded & "DescMovEco,TE;ImptMovEco,FLP"
   %>
   
			<div class="row">
               <div class="col-2">
               </div> 
			   <div class="col-6">
                  <div class="form-group ">
				     <%xx=ShowLabel("Descrizione Ricarica")
					   nn="DescMovEco" & l_Id
					   vv=Getdiz(DizDatabase,"DescMovEco")
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=vv%>" >
                  </div>		
			   </div>
			</div>
			<div class="row">
               <div class="col-2">
               </div> 
			
			   <div class="col-2">
                  <div class="form-group ">
				     <%xx=ShowLabel("Importo ricarica &euro;")
					 nn="ImptMovEco" & l_Id
					 vv=Getdiz(DizDatabase,"ImptMovEco")
					 %>
					 <input type="text" <%=readonly%> class="form-control" name="<%=nn%>" id="<%=nn%>" value="<%=vv%>" >
                  </div>		
			   </div>	
			</div>
			<div class="row">
               <div class="col-2">
               </div> 
			      
			      <%
				    conta=1
					InsertRow      = true 
					canUpload      = not SoloLettura
					IsFileUploadMd = "N"
					NomeFileUpload = ""
                    Linkdocumento  = ""
                    DescFileUpload = ""
					IdTableUpload  = Getdiz(DizDatabase,"IdUpload")
					IdModUpload    = "_U0"
					modalUploadFilesId = modalUploadFilesId & ";" & IdModUpload
                    if cdbl(IdTableUpload)>0 then 
                        Linkdocumento =LeggiCampo("select * from Upload Where IdUpload=" & IdTableUpload,"PathDocumento")
						NomeFileUpload=LeggiCampo("select * from Upload Where IdUpload=" & IdTableUpload,"NomeDocumento")
                        DescFileUpload=LeggiCampo("select * from Upload Where IdUpload=" & IdTableUpload,"DescBreve")
                    end if 
  
				  %>
				  <div class="col-6">
				  <!--#include virtual="/gscVirtual/utility/modalUpload/showUploadDoc.asp"-->
                  </div>
            </div>   
     <%if SoloLettura=false then%>
	    <br>
	 <div class="row">
		<div class="col-2">
		</div>	 
		<div class="col-2">
		   <button type="button" onclick="localFun('update','0')" class="btn btn-warning">Registra</button>
		</div>
	    <div class="col-2">
		   <button type="button" onclick="localFun('prosegui','0')" class="btn btn-success">&nbsp;&nbsp;Invia&nbsp;&nbsp;</button>
        </div>   
   
     </div>
     <%end if %>
   <input type="hidden" name="localVirtualPath" id="localVirtualPath" value = "<%=VirtualPath%>">
   
			<!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
			
			<input type="hidden" name="modalUploadFilesId" id="modalUploadFilesId" value="<%=modalUploadFilesId%>">
			</form>
        <!--#include virtual="/gscVirtual/utility/modalUpload/ModalUpload.asp"-->
		<!--#include virtual="/gscVirtual/utility/modalUpload/ModalUploadScript.asp"-->
		</div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

</body>

</html>
