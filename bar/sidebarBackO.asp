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
	d.Add "prof",defColor
	d.Add "affi",defColor
	d.Add "paga",defColor
	d.Add "cauz",defColor
	d.Add "clie",defColor
	d.Add "form",defColor

	d.Add "dashS",styleBase
	d.Add "profS",styleBase
	d.Add "affiS",styleBase
	d.Add "pagaS",styleBase
	d.Add "cauzS",styleBase
	d.Add "clieS",styleBase
	d.Add "formS",styleBase

	showForm = false 
	showAffi = false
	showCauz = false 

	if isServizioAttivo("CAUZ_PROV") then 
	   showAffi = true
	   showCauz = true 
	end if 
	if isServizioAttivo("CAUZ_DEFI") then 
	   showCauz = true 
	end if 	
	if isServizioAttivo("FORMAZ") then 
       showForm = true 
	end if 
	
	d.remove opzioneSidebar
	d.remove opzioneSidebar & "S"
	d.Add opzioneSidebar,tarColor
	d.Add opzioneSidebar & "S",styleSele
	
	%>

    <!-- Sidebar -->
    <div class="bg-dark border-right" id="sidebar-wrapper">
      <div class="sidebar-heading">
		<a href="<%=VirtualPath%>/bar/BackODashboard.asp">
		<span><img src="/gscVirtual/img/logo_piccolo.png"  alt="homepage" />                            
        </span>
		</a>
	  </div>
      <div class="list-group list-group-flush bg-dark">
        <a href="<%=VirtualPath%>bar/BackODashboard.asp" class="list-group-item list-group-item-action text-white <%=d.item("dash")%>">
		 <i class="fa fa-2x fa-home" <%=d.item("dashS")%> ></i>&nbsp;Dashboard</a>
        <a href="<%=VirtualPath%>configurazioni/Collaboratori/BackOModifica.asp" class="list-group-item list-group-item-action text-white <%=d.item("prof")%>">
		 <i class="fa fa-2x fa-user" <%=d.item("profS")%> ></i>&nbsp;Profilo</a>	
        <a href="<%=VirtualPath%>link/BackOPagamento.asp" class="list-group-item list-group-item-action text-white <%=d.item("paga")%>">
		<i class="fa fa-2x fa-money"  <%=d.item("pagaS")%> ></i>&nbsp;Pagamenti</a>
        <%if showAffi then %>
        <a href="<%=VirtualPath%>link/BackOAffidamento.asp" class="list-group-item list-group-item-action text-white <%=d.item("affi")%>">
		<i class="fa fa-2x fa-handshake-o"  <%=d.item("affiS")%> ></i>&nbsp;Affidamento</a>
		<%end if %>
		<%if showCauz then %>
        <a href="<%=VirtualPath%>link/BackOCauzioneProvvisoria.asp" class="list-group-item list-group-item-action text-white <%=d.item("cauz")%>">
		<i class="fa fa-2x fa-credit-card"  <%=d.item("cauzS")%> ></i>&nbsp;Cauzioni</a>
		<%end if %>
		<%if showForm then %>
        <a href="<%=VirtualPath%>link/backOFormazione.asp" class="list-group-item list-group-item-action text-white <%=d.item("form")%>">
		<i class="fa fa-2x fa-graduation-cap"  <%=d.item("formS")%> ></i>&nbsp;Formazione</a>
		<%end if %>
		
        <a href="<%=VirtualPath%>configurazioni/Clienti/ListaClientiBack.asp" class="list-group-item list-group-item-action text-white <%=d.item("clie")%>">
		<i class="fa fa-2x fa-users"  <%=d.item("clieS")%> ></i>&nbsp;Clienti</a>
		
		</div>
    </div>
    <!-- /#sidebar-wrapper -->