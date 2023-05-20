<%
  NomePagina="ProdottiFornitoreMail.asp"
  titolo="Menu Supervisor - Dashboard"
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
  livelloPagina="00"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include virtual="/gscVirtual/include/head.asp"-->
<!-- Custom styles for this  -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">
<script>
function update()
{
	xx=ImpostaValoreDi("DescLoaded","0");
	xx=ElaboraControlli();
	
 	if (xx==false)
	   return false;
	
	ImpostaValoreDi("Oper","update");
	document.Fdati.submit();
}

</script>
</head>

<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<%
NameLoaded=NameLoaded & "MailDocumentazione,EM"

Set Rs = Server.CreateObject("ADODB.Recordset")

IdProdotto   = 0
IdFornitore  = 0
IdAccount    = 0
DescFornitore= ""
DescProdotto = ""
if FirstLoad then 
   PaginaReturn    = getCurrentValueFor("PaginaReturn")
   IdProdotto      = "0" & getCurrentValueFor("IdProdotto")
   IdFornitore     = "0" & getCurrentValueFor("IdFornitore")   
   if cdbl(IdFornitore)>0 then 
      Rs.CursorLocation = 3 
      Rs.Open "Select * from Fornitore where IdFornitore=" & IdFornitore, ConnMsde   
      IdAccount     = Rs("IdAccount")
	  DescFornitore = Rs("DescFornitore")
      Rs.close 
   end if      
   if cdbl(IdProdotto)>0 then 
      Rs.CursorLocation = 3 
      Rs.Open "Select * from Prodotto where IdProdotto=" & IdProdotto, ConnMsde   
	  DescProdotto = Rs("DescProdotto")
      Rs.close 
   end if    
else
   IdProdotto     = "0" & getValueOfDic(Pagedic,"IdProdotto")
   IdAccount      = "0" & getValueOfDic(Pagedic,"IdAccount")
   IdFornitore    = "0" & getValueOfDic(Pagedic,"IdFornitore")
   DescProdotto   = getValueOfDic(Pagedic,"DescProdotto")
   DescFornitore  = getValueOfDic(Pagedic,"DescFornitore")
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
end if 
if cdbl(IdProdotto)=0 or Cdbl(IdFornitore)=0  then 
   response.redirect virtualPath & PaginaReturn
   response.end 
end if 

IdAccount =cdbl(IdAccount)
IdProdotto=cdbl(IdProdotto)
on error resume next
 
if Oper="UPDATE" then 
    Session("TimeStamp")=TimePage
	MailDocumentazione=Request("MailDocumentazione0")
	CodiceProdotto=Request("CodiceProdotto0")
	IdProcessoElaborativo=Request("IdProcessoElaborativo0")
	if IdProcessoElaborativo="-1" then 
	   IdProcessoElaborativo=""
	end if 
	LinkWeb=Request("LinkWeb0")
	MyQ = "" 
	MyQ = MyQ & " Update AccountProdotto set "
	MyQ = MyQ & " MailDocumentazione = '" & apici(MailDocumentazione) & "'"
	MyQ = MyQ & ",CodiceProdotto = '" & apici(CodiceProdotto) & "'"
	MyQ = MyQ & ",LinkWeb = '" & apici(LinkWeb) & "'"
	MyQ = MyQ & ",IdProcessoElaborativo='" & apici(IdProcessoElaborativo) & "'"
	MyQ = MyQ & " where IdProdotto = " & IdProdotto
	MyQ = MyQ & " and   IdAccount  = " & IdAccount 
	ConnMsde.execute MyQ
	response.redirect VirtualPath & PaginaReturn
	response.end

End if 
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdProdotto"     ,IdProdotto)
  xx=setValueOfDic(Pagedic,"DescProdotto"   ,DescProdotto)
  xx=setValueOfDic(Pagedic,"IdFornitore"    ,IdFornitore)
  xx=setValueOfDic(Pagedic,"DescFornitore"  ,DescFornitore)
  xx=setValueOfDic(Pagedic,"IdAccount"      ,IdAccount)
  xx=setValueOfDic(Pagedic,"PaginaReturn"   ,PaginaReturn)
  
  xx=setCurrent(NomePagina,livelloPagina) 
  err.clear 

%>

<% 
  'xx=DumpDic(SessionDic,NomePagina)
%>
<div class="d-flex" id="wrapper">

	<%
	  TitoloNavigazione="Configurazioni"
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
			<%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h4>Configurazione Prodotto Fornitore</h4>
				</div>
			</div>
			<div class="row">
			   <div class="col-1">
			   </div>
               <div class="col-5">
                  <div class="form-group ">
                     <%xx=ShowLabel("Fornitore")%>
                     <input type="text" readonly class="form-control input-sm" value="<%=DescFornitore%>" >
                  </div>        
               </div>			
               <div class="col-5">
                  <div class="form-group ">
                     <%xx=ShowLabel("Prodotto")%>
                     <input type="text" readonly class="form-control input-sm" value="<%=DescProdotto%>" >
                  </div>        
               </div>			
			   
			</div> 			
<%

MySql = "" 
MySql = MySql & " Select *"
MySql = MySql & " from AccountProdotto"
MySql = MySql & " where IdProdotto = " & IdProdotto
MySql = MySql & " and   IdAccount = " & IdAccount

MailDocumentazione    = LeggiCampo(MySql,"MailDocumentazione")
CodiceProdotto        = LeggiCampo(MySql,"CodiceProdotto")
LinkWeb               = LeggiCampo(MySql,"LinkWeb")
IdProcessoElaborativo = LeggiCampo(MySql,"IdProcessoElaborativo")
'response.write MySql 
%>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
   <br>
			<div class="row">
			   <div class="col-1">
			   </div>
               <div class="col-10">
                  <div class="form-group ">
                     <%xx=ShowLabel("Mail Per Documentazione")%>
                     <input type="text" id="MailDocumentazione0" name="MailDocumentazione0" 
					 class="form-control input-sm" value="<%=MailDocumentazione%>" >
                  </div>        
               </div>
            </div>   
			<div class="row">
			   <div class="col-1">
			   </div>
               <div class="col-4">
                  <div class="form-group ">
                     <%xx=ShowLabel("Codice Prodotto Fornitore")%>
                     <input type="text" id="CodiceProdotto0" name="CodiceProdotto0" 
					 class="form-control input-sm" value="<%=CodiceProdotto%>" >
                  </div>        
               </div>
			   <div class="col-1">
			   </div>
               <div class="col-4">
                  <div class="form-group ">
                     <%xx=ShowLabel("Processo elaborativo")
  			           query = ""
			           query = query & " Select * from ProcessoElaborativo " 
					   query = query & " Where TipoProcesso = 'ATTIVA_PRODOTTO' " 
			 		   query = query & " order By DescProcessoElaborativo"
			           response.write ListaDbChangeCompleta(Query,"IdProcessoElaborativo0",IdProcessoElaborativo,"IdProcessoElaborativo","DescProcessoElaborativo",1,"","","","","dati assenti","class='form-control form-control-sm'")					 
					 
					 %>
                  </div>        
               </div>			   
            </div>   
			<div class="row">
			   <div class="col-1">
			   </div>
               <div class="col-10">
                  <div class="form-group ">
                     <%xx=ShowLabel("Sito web per utilizzo prodotto")%>
                     <input type="text" id="LinkWeb0" name="LinkWeb0" 
					 class="form-control input-sm" value="<%=LinkWeb%>" >
                  </div>        
               </div>
            </div>   
			
			
   <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
   
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;update();S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div></div>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
		</div>
		
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
