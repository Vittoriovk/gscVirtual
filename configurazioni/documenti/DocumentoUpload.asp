<%
  NomePagina="DocumentoUpload.asp"
  titolo="Menu - Dashboard"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
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
<!--#include file="DocumentoUploadLogica.asp"-->
<%   
  'xx=DumpDic(SessionDic,NomePagina)
%>
<div class="d-flex" id="wrapper">
	<%
	  TitoloNavigazione="Caricamento Documento"
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
            <!--#include file="DocumentoUploadElabora.asp"-->
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
