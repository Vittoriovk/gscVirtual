<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<!--#include virtual="/gscVirtual/common/functionNew.asp"-->
	<%

	defColor=default_nav_defColor
	tarColor=default_nav_tarColor
	
	styleBase=default_nav_styleBase
	styleSele=default_nav_styleSele

	if opzioneSidebar="" then 
	   opzioneSidebar=Session("opzioneSidebar")
	end if	
	if opzioneSidebar="" then 
	   opzioneSidebar="dash"
	end if 
	if VirtualPath="" then 
	   VirtualPath=Session("VirtualPath")
	end if 
	
	Dim d
	Set d=Server.CreateObject("Scripting.Dictionary")
	'aggiungo tutti i default 
	d.Add "dash",defColor
	d.Add "dashS",styleBase
	
	d.Add "prof",defColor
	d.Add "profS",styleBase
	
	d.Add "paga",defColor
	d.Add "pagaS",styleBase

    d.Add "util",defColor
    d.Add "utilS",styleBase

    d.Add "coob",defColor
    d.Add "coobS",styleBase

    d.Add "atii",defColor
    d.Add "atiiS",styleBase
	

    d.Add "docu",defColor
    d.Add "docuS",styleBase

    d.Add "affi",defColor
    d.Add "affiS",styleBase

    d.Add "affa",defColor
    d.Add "affaS",styleBase

    d.Add "affc",defColor
    d.Add "affcS",styleBase

    d.Add "cauz",defColor
    d.Add "cauzS",styleBase
	
    d.Add "form",defColor
    d.Add "formS",styleBase

	   
	showCoob = false 
	showAffi = false
	showCazu = false 
	showDocu = false
	showForm = false 
	if isServizioAttivo("CAUZ_PROV") then 
	   showCoob = true 
	   showAffi = true
	   showCazu = true 
	   showDocu = true
	end if 
	
	
	

	d.remove opzioneSidebar
	d.remove opzioneSidebar & "S"
	d.Add opzioneSidebar,tarColor
	d.Add opzioneSidebar & "S",styleSele
	
	%>

    <!-- Sidebar -->
    <div class="bg-dark border-right" id="sidebar-wrapper">
      <div class="sidebar-heading">
		<a href="<%=VirtualPath%>/bar/ClieDashboard.asp">
		<span><img src="/gscVirtual/img/logo_piccolo.png"  alt="homepage" />                            
        </span>
		
		</a>
	  </div>
      <div class="list-group list-group-flush bg-dark">
        <a href="<%=VirtualPath%>bar/ClieDashboard.asp" class="list-group-item list-group-item-action text-white <%=d.item("dash")%>">
		 <i class="fa fa-2x fa-home" <%=d.item("dashS")%> ></i>&nbsp;Dashboard</a>

		<% if showCoob then %>
           <a href="<%=VirtualPath%>configurazioni/Clienti/ClienteCoobbligatiSwap.asp" class="list-group-item list-group-item-action text-white <%=d.item("coob")%>">
		   <i class="fa fa-2x fa-users"  <%=d.item("coobS")%> ></i>&nbsp;Coobbligati</a>
		   
           <a href="<%=VirtualPath%>configurazioni/Clienti/ClienteAtiSwap.asp" class="list-group-item list-group-item-action text-white <%=d.item("atii")%>">
		   <i class="fa fa-2x fa-industry"  <%=d.item("atiiS")%> ></i>&nbsp;A.T.I.</a>		
  
		   
		<% end if %>
		 
		<% if showDocu then %>
           <a href="<%=VirtualPath%>configurazioni/Clienti/DocumentiClienteSwap.asp" class="list-group-item list-group-item-action text-white <%=d.item("docu")%>">
		   <i class="fa fa-2x fa-file-text-o"  <%=d.item("docuS")%> ></i>&nbsp;Documenti</a>
		<% end if %>
		
		<% if showAffi then %>
        <a href="<%=VirtualPath%>link/ClienteAffidamento.asp" class="list-group-item list-group-item-action text-white <%=d.item("affi")%>">
		<i class="fa fa-2x fa-handshake-o"  <%=d.item("affiS")%> ></i>&nbsp;Affidamento</a>
		<% end if %>
		
		<% if showAffi then %>
        <a href="<%=VirtualPath%>link/ClienteCauzione.asp" class="list-group-item list-group-item-action text-white <%=d.item("cauz")%>">
		<i class="fa fa-2x fa-credit-card"  <%=d.item("cauzS")%> ></i>&nbsp;Cauzioni</a>
		<%end if %>

		<% if showForm then %>
        <a href="<%=VirtualPath%>link/ClienteFormazione.asp" class="list-group-item list-group-item-action text-white <%=d.item("form")%>">
		<i class="fa fa-2x fa-graduation-cap"  <%=d.item("formS")%> ></i>&nbsp;Formazione</a>
		<%end if %>
		
        <a href="<%=VirtualPath%>link/ClientePagamento.asp" class="list-group-item list-group-item-action text-white <%=d.item("paga")%>">
		<i class="fa fa-2x fa-credit-card-alt"  <%=d.item("pagaS")%> ></i>&nbsp;Pagamenti</a>
		
        <a href="<%=VirtualPath%>link/ClieUtility.asp" class="list-group-item list-group-item-action text-white <%=d.item("util")%>">
        <i class="fa fa-2x fa-cogs"  <%=d.item("utilS")%> ></i>&nbsp;Utility</a>		
		
		</div>
    </div>
    <!-- /#sidebar-wrapper -->