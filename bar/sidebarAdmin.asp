<!--#include virtual="/gscVirtual/common/parametri.asp"-->
	<%

	defColor=default_nav_defColor
	tarColor=default_nav_tarColor
	
	styleBase=default_nav_styleBase
	styleSele=default_nav_styleSele
	
	if opzioneSidebar="" then 
	   opzioneSidebar=Session("opzioneSidebar")
	end if	
	if opzioneSidebar="" then 
	   opzioneSidebar="conf"
	end if 
	if VirtualPath="" then 
	   VirtualPath=Session("VirtualPath")
	end if 
	
	Dim d
	Set d=Server.CreateObject("Scripting.Dictionary")
	'aggiungo tutti i default 
	d.Add "dash",defColor
	d.Add "conf",defColor
	d.Add "prof",defColor
	d.Add "inte",defColor
	d.Add "oper",defColor
	d.Add "auto",defColor
	d.Add "prov",defColor

	d.Add "dashS",styleBase
	d.Add "confS",styleBase
	d.Add "profS",styleBase
	d.Add "inteS",styleBase
	d.Add "operS",styleBase
	d.Add "autoS",styleBase
	d.Add "provS",styleBase
	
	d.remove opzioneSidebar
	d.remove opzioneSidebar & "S"
	d.Add opzioneSidebar,tarColor
	d.Add opzioneSidebar & "S",styleSele
	%>

    <!-- Sidebar -->
    <div class="bg-dark border-right" id="sidebar-wrapper">
      <div class="sidebar-heading">
		<a href="<%=VirtualPath%>/bar/AdminDashboard.asp">
		<span><img src="/gscVirtual/img/logo_piccolo.png"  alt="homepage" />                            
        </span>
		
		</a>
	  </div>
      <div class="list-group list-group-flush bg-dark">
        <a href="<%=VirtualPath%>bar/AdminDashboard.asp" class="list-group-item list-group-item-action text-white <%=d.item("dash")%>">
		 <i class="fa fa-2x fa-home" <%=d.item("dashS")%> ></i>&nbsp;Dashboard</a>
        <a href="<%=VirtualPath%>configurazioni/AnagraficaSuperV.asp" class="list-group-item list-group-item-action text-white <%=d.item("prof")%>">
		 <i class="fa fa-2x fa-user" <%=d.item("profS")%> ></i>&nbsp;Profilo</a>
        <a href="<%=VirtualPath%>link/AdminConfigurazioni.asp" class="list-group-item list-group-item-action <%=d.item("conf")%>">
		<i class="fa fa-2x fa-bars"  <%=d.item("confS")%> ></i>&nbsp;Configurazioni</a>		 
        <a href="<%=VirtualPath%>link/AdminIntermediari.asp" class="list-group-item list-group-item-action text-white <%=d.item("inte")%>">
		<i class="fa fa-2x fa-users"  <%=d.item("inteS")%> ></i>&nbsp;Collaboratori</a>
        <a href="<%=VirtualPath%>link/AdminOperatori.asp" class="list-group-item list-group-item-action text-white <%=d.item("oper")%>">
		<i class="fa fa-2x fa-user"  <%=d.item("operS")%> ></i>&nbsp;Utenti</a>
        <a href="<%=VirtualPath%>link/AdminAutorizzazioni.asp" class="list-group-item list-group-item-action text-white <%=d.item("auto")%>">
		<i class="fa fa-2x fa-key"  <%=d.item("autoS")%> ></i>&nbsp;Autorizzazioni</a>
        <a href="<%=VirtualPath%>link/AdminProvvigioni.asp" class="list-group-item list-group-item-action text-white <%=d.item("prov")%>">
		<i class="fa fa-2x fa-percent"  <%=d.item("provS")%> ></i>&nbsp;Provvigioni</a>
		</div>
    </div>
    <!-- /#sidebar-wrapper -->