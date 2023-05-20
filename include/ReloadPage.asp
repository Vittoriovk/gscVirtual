   <%
   'eseguo la logica solo quando ricarica 
   if FirstLoad=true  then 
      %>
	  <form name="Fdati" Action="<%=NomePagina%>" method="post">
	  <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
	  </form>
	  <%
      response.write "<script language=javascript>document.Fdati.submit();</script>" 
      response.end
   end if
   %>