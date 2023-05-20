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
      $('[data-toggle="tooltip"]').tooltip();  
    });
  </script>


<%
if FirstLoad then 
	response.write "<script language=javascript>document.Fdati.submit();</script>" 
	response.end 
end if
%>
