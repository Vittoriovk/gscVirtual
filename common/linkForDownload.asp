<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<%
if linkDocumento<>"" then 
   IdlinkForDownload="linkForDownload_" & TimeToS()
   Estensione=ucase(right(Trim(linkDocumento),3))
   icona="fa-file-pdf-o"
   if Estensione="ZIP" or Estensione="RAR" then 
      icona="fa-file-archive-o"
   end if 
%>
   <a Id="<%=IdlinkForDownload%>" href='/gscVirtual/utility/download.asp?tt=<%=linkDocumento%>' title="Mostra Documento" target="_new">
	<i class="fa fa-2x <%=icona%>"></i></a>  
<%
end if 
%>

