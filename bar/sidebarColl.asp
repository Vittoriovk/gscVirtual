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
	d.Add "inte",defColor
	d.Add "clie",defColor
	d.Add "cont",defColor
	d.Add "util",defColor
	d.Add "oper",defColor
	

	d.Add "dashS",styleBase
	d.Add "profS",styleBase
	d.Add "inteS",styleBase
	d.Add "clieS",styleBase
	d.Add "contS",styleBase
	d.Add "utilS",styleBase
	d.Add "operS",styleBase

	flagCauzProv = false 
	flagCauzDefi = false 
	flagFormaz   = false 
	
	if instr(Session("Login_servizi_attivi"),"CAUZ_PROV")>0 then 
	   flagCauzProv = true  
	end if 
	if instr(Session("Login_servizi_attivi"),"CAUZ_DEFI")>0 then 
	   flagCauzDefi = true  
	end if 
	if instr(Session("Login_servizi_attivi"),"FORMAZ")>0 then 
	   flagFormaz = true  
	end if 
	
	if flagCauzProv = true then 
	   d.Add "affi",defColor
	   d.Add "affiS",styleBase
	   d.Add "affa",defColor
	   d.Add "affaS",styleBase
	   d.Add "affc",defColor
	   d.Add "affcS",styleBase	
	end if 
	if flagCauzProv = true or flagCauzDefi=true then 
	   d.Add "cauz",defColor
	   d.Add "cauzS",styleBase
	end if 	
    if flagFormaz = true then
	   d.Add "form",defColor
	   d.Add "formS",styleBase	
	end if 
	
	d.remove opzioneSidebar
	d.remove opzioneSidebar & "S"
	d.Add opzioneSidebar,tarColor
	d.Add opzioneSidebar & "S",styleSele
	
	%>

    <!-- Sidebar -->
    <div class="bg-dark border-right" id="sidebar-wrapper">
      <div class="sidebar-heading">
		<a href="<%=VirtualPath%>/bar/CollDashboard.asp">
		<span><img src="/gscVirtual/img/logo_piccolo.png"  alt="homepage" />                            
        </span>
		
		</a>
	  </div>
      <div class="list-group list-group-flush bg-dark">
        <a href="<%=VirtualPath%>bar/CollDashboard.asp" class="list-group-item list-group-item-action text-white <%=d.item("dash")%>">
		 <i class="fa fa-2x fa-home" <%=d.item("dashs")%> ></i>&nbsp;Dashboard</a>
        <%if Session("LoginTipoCollaboratore")<> "SEGN" and session("FlagGeneraCollaboratore")="1" then %>		 
        <a href="<%=VirtualPath%>configurazioni/collaboratori/ListaCollaboratoreIntermediario.asp" class="list-group-item list-group-item-action text-white <%=d.item("inte")%>">
		<i class="fa fa-2x fa-users"  <%=d.item("inteS")%> ></i>&nbsp;Collaboratori</a>
		<%end if %>
		
        <a href="<%=VirtualPath%>configurazioni/collaboratori/ListaOperatoreColl.asp" class="list-group-item list-group-item-action text-white <%=d.item("oper")%>">
		<i class="fa fa-2x fa-user"  <%=d.item("operS")%> ></i>&nbsp;Utenti</a>
		
        <a href="<%=VirtualPath%>configurazioni/Clienti/ListaClientiColl.asp" class="list-group-item list-group-item-action text-white <%=d.item("clie")%>">
		<i class="fa fa-2x fa-industry"  <%=d.item("clieS")%> ></i>&nbsp;Aziende</a>
		<%if Session("LoginTipoCollaboratore")<> "SEGN" then %>		
           <%if flagCauzProv = true then %>
		   
             <a href="<%=VirtualPath%>link/CollAffidamento.asp" class="list-group-item list-group-item-action text-white <%=d.item("affi")%>">
		     <i class="fa fa-2x fa-handshake-o"  <%=d.item("affiS")%> ></i>&nbsp;Affidamento</a>

             <a href="<%=VirtualPath%>link/CollAffidamentoATI.asp" class="list-group-item list-group-item-action text-white <%=d.item("affa")%>">
		     <i class="fa fa-2x fa-university"  <%=d.item("affaS")%> ></i>&nbsp;ATI</a>
		   
             <a href="<%=VirtualPath%>link/CollAffidamentoCOOB.asp" class="list-group-item list-group-item-action text-white <%=d.item("affc")%>">
		     <i class="fa fa-2x fa-id-card-o"  <%=d.item("affcS")%> ></i>&nbsp;Coobbligato</a>
		   
		   <%end if %>
		   <%if flagCauzProv = true or flagCauzDefi=true then %>
             <a href="<%=VirtualPath%>link/CollCauzioni.asp" class="list-group-item list-group-item-action text-white <%=d.item("cauz")%>">
             <i class="fa fa-2x fa-shield"  <%=d.item("cauzS")%> ></i>&nbsp;Cauzioni</a> 
		   <%end if %>
		<%end if %>

		<%if flagFormaz = true then %>
           <a href="<%=VirtualPath%>link/CollFormazione.asp" class="list-group-item list-group-item-action text-white <%=d.item("form")%>">
           <i class="fa fa-2x fa-graduation-cap"  <%=d.item("formS")%> ></i>&nbsp;Formazione</a>
		<%end if %>

        <a href="<%=VirtualPath%>link/CollContabilita.asp" class="list-group-item list-group-item-action text-white <%=d.item("cont")%>">
        <i class="fa fa-2x fa-money"  <%=d.item("contS")%> ></i>&nbsp;Contabilit&agrave;</a>
		
        <a href="<%=VirtualPath%>link/CollUtility.asp" class="list-group-item list-group-item-action text-white <%=d.item("util")%>">
        <i class="fa fa-2x fa-cogs"  <%=d.item("utilS")%> ></i>&nbsp;Utility</a>
		
		</div>
    </div>
    <!-- /#sidebar-wrapper -->