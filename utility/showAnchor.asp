<%
  NomePagina="ShowAchor.asp"
  titolo="test"
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
<!-- Custom styles for this template -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">

</head>


<%
Lista=""
Lista=lista & "plus,prev,upda,dele,save,dett,dett-t,dett-g"
Lista=lista & ",tecn,prof,clie,pdf,docu,prod,forn,perc,from,money"
Lista=lista & ",pict,hand,info,mail,lucc,puli,penn,manu,ok,ko,uplo"
Lista=lista & ",matr,cert,sele,crea,minu,card,requ,remo,erre,logi" 
Lista=lista & ",copy,effe,"

arDati=split(Lista,",")
for j=lbound(arDati) to uBound(arDati)-1 
   opt=ArDati(j)
   RiferimentoA="col-2;#;;2;" & opt & "; --" & opt & "--;;localGes();N"
   response.write "<BR>" & opt
   %>
   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
<% next %>

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
  <script>
$('.btn').onClick(function(e){
  e.preventDefault();
});  
</script>
</body>

</html>



