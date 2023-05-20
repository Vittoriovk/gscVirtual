<%
  NomePagina="PagamentoAccountModBackO.asp"
  titolo="Modifica Pagamento per account"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/modelli/functionPagamenti.asp"-->
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
  IdMovEco  = 0
  if FirstLoad then 
     IdAccount       = cdbl("0" & getCurrentValueFor("IdAccount"))
     IdMovEco        = cdbl("0" & getCurrentValueFor("IdMovEco"))
     IdTipoCredito   = getCurrentValueFor("IdTipoCredito")
     TipoUtente      = getCurrentValueFor("TipoUtente")
     DescUtente      = getCurrentValueFor("DescUtente")
     OperTabella     = getCurrentValueFor("OperTabella")
     PaginaReturn    = getCurrentValueFor("PaginaReturn") 
  else
     IdAccount       = cdbl("0" & getValueOfDic(Pagedic,"IdAccount"))
     IdMovEco        = cdbl("0" & getValueOfDic(Pagedic,"IdMovEco"))
     IdTipoCredito   = getValueOfDic(Pagedic,"IdTipoCredito")
     DescTipoCredito = getValueOfDic(Pagedic,"DescTipoCredito")
     TipoUtente      = getValueOfDic(Pagedic,"TipoUtente")
     DescUtente      = getValueOfDic(Pagedic,"DescUtente")     
     OperTabella     = getValueOfDic(Pagedic,"OperTabella")
     PaginaReturn    = getValueOfDic(Pagedic,"PaginaReturn")
   end if   
  
   IdAccount       = cdbl(IdAccount)
   IdMovEco        = cdbl(IdMovEco)

   'per ora non gestisco l'inserimento ma solo la modifica 
   if cdbl(IdMovEco)=0 then 
      response.redirect RitornaA(PaginaReturn)
      response.end 
   end if 
  
   if IdMovEco=0 then 
      OperAmmesse = "IU"
   end if 
   
  'inserisco account 
   flagProsegui="" 
   tipoEvento  =""
   flagEsci    =false
   if Oper=ucase("valida") or Oper=ucase("integra") or Oper=ucase("annulla") then 
      if Oper=ucase("valida")  then 
	     flagProsegui="ATTI"
		 tipoEvento  ="ACCE"
	  end if 
	  if Oper=ucase("integra") then 
	     flagProsegui="INCO"
		 tipoEvento  ="INTE"
	  end if 
	  if Oper=ucase("annulla") then 
	     flagProsegui="ANNU"
		 tipoEvento  ="ANNU"
	  end if 
      
      Oper=ucase("update")
   end if 
   
   if Oper=ucase("update") then 
      Ritorna=false 
      OperAmmesse="U"
      Session("TimeStamp")=TimePage
      MsgErrore=""
      ImptMovEco       = Request("ImptMovEco0")
      DescMovEco       = Request("DescMovEco0")
	  DescStatoCredito = Trim(Request("DescStatoCredito0"))
      
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
            flagEsci=false 
         end if 
      else 
         err.clear 
         MyQ = MyQ & " update AccountMovEco set "
         MyQ = MyQ & " DescMovEco ='" & apici(DescMovEco) & "'"  
		 MyQ = MyQ & ",DescStatoCredito ='" & apici(DescStatoCredito) & "'"  
         if flagProsegui<>"" then 
            MyQ = MyQ & ",IdStatoCredito ='" & flagProsegui & "'"   
         end if 
         MyQ = MyQ & ",ImptMovEco = " & NumForDb(ImptMovEco) 
         MyQ = MyQ & " Where IdAccountMovEco = " & IdMovEco 
         MyQ = MyQ & " and   IdAccount=" & IdAccount
         ConnMsde.execute MyQ 
         'response.write MyQ
         if err.number=0 then 
            xx=UpdatePagaAccount(IdAccount,IdTipoCredito)
			if flagProsegui<>"" then 
               IdProdotto=0
               'response.write "ecco"
			   if DescStatoCredito="" then 
			      DescStatoCredito=DescMovEco 
			   end if 
               XX=createEvento(IdTipoCredito,tipoEvento,Session("LoginIdAccount"),DescStatoCredito,"AccountMovEco","IdAccountMovEco=" & IdMovEco,true,IdProdotto)
			end if 
            flagEsci=false
         end if 
      end if 
   end if 
   'response.write flagEsci
   if flagEsci=true then 
      response.redirect VirtualPath & PaginaReturn
      response.end 
   end if 
   
   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
   
   if cdbl(IdMovEco)>0 then
      OperAmmesse="U"
      MySql = ""
      MySql = MySql & " Select * From  AccountMovEco "
      MySql = MySql & " Where IdAccountMovEco=" & IdMovEco
      xx=GetInfoRecordset(DizDatabase,MySql)
   else
     OperAmmesse="IU"   
   end if 
   
   if IdTipoCredito="" and cdbl(IdMovEco)>0 then 
      IdTipoCredito   = GetDiz(DizDatabase,"IdTipoCredito")
      DescTipoCredito = LeggiCampoTabella("TipoCredito",IdTipoCredito) 
   end if 
   if IdAccount=0 and cdbl(IdMovEco)>0 then 
      IdAccount  = GetDiz(DizDatabase,"IdAccount")
      
      Dim DizAccount
      Set DizAccount = CreateObject("Scripting.Dictionary")
      xx=GetInfoRecordset(DizAccount,"select * from Account where IdAccount=" & IdAccount)
      TipoUtente = GetDiz(DizAccount,"IdTipoAccount")

      TipoUtente = LeggiCampoTabellaText("TipoAccount",TipoUtente)
      DescUtente = GetDiz(DizAccount,"Nominativo")
   end if 

  'registro i dati della pagina 
  if descTipoCredito="" then 
     DescTipoCredito=LeggiCampo("select * from TipoCredito where IdTipoCredito='" & IdTipoCredito & "'","DescTipoCredito")
  end if 
  
  xx=setValueOfDic(Pagedic,"DescUtente"     ,DescUtente)
  xx=setValueOfDic(Pagedic,"TipoUtente"     ,TipoUtente)
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
            <%RiferimentoA="col-1;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <div class="col-11"><h3>Gestione Movimenti Contabili</h3>
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
               <% xx=writeDiv(2,"Tipo movimento",DescTipoCredito,"","") %>
               <% xx=writeDiv(2,"Tipo Utente"   ,TipoUtente     ,"","") %>
               <% xx=writeDiv(4,"Nominativo"    ,DescUtente     ,"","") %>
            </div>
            <div class="row">
               <% 
			   DescStatoCredito=LeggiCampoTabellaText("StatoCredito",GetDiz(DizDatabase,"IdStatoCredito"))
			   xx=writeDiv(2,"Stato movimento",DescStatoCredito,"","") %>			
			</div>
            <br>
			<br>
            <div class="row">
               <% xx=writeDiv(8,"Descrizione Movimento"   ,Getdiz(DizDatabase,"DescMovEco") ,"DescMovEco" & l_Id,readonly) %>
            </div>
            <div class="row">
               <% xx=writeDiv(2,"Importo ricarica &euro;" ,Getdiz(DizDatabase,"ImptMovEco") ,"ImptMovEco" & l_Id,readonly) %>
            </div>
            <div class="row">
               <% xx=writeDiv(8,"Annotazioni"             ,Getdiz(DizDatabase,"DescStatoCredito") ,"DescStatoCredito" & l_Id,readonly) %>            
            </div>
     <%if SoloLettura=false then%>
     <div class="row">
        <div class="col-1">
        </div>     
        <div class="col-2">
           <button type="button" onclick="localFun('update','0')"   class="btn btn-warning">Registra</button>
        </div>
        <div class="col-2">
           <button type="button" onclick="localFun('valida','0')" class="btn btn-success">&nbsp;&nbsp;Valida&nbsp;&nbsp;</button>
        </div>   
        <div class="col-2">
           <button type="button" onclick="localFun('integra','0')"  class="btn btn-info">&nbsp;Integrazione&nbsp;</button>
        </div>   		
        <div class="col-2">
           <button type="button" onclick="localFun('annulla','0')"  class="btn btn-danger">&nbsp;&nbsp;Annulla&nbsp;&nbsp;</button>
        </div>    
     </div>
     <%end if %>
   
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
