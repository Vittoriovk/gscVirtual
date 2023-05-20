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

	d.Add "dashS",styleBase
	d.Add "confS",styleBase
	d.Add "profS",styleBase
	
	d.remove opzioneSidebar
	d.remove opzioneSidebar & "S"
	d.Add opzioneSidebar,tarColor
	d.Add opzioneSidebar & "S",styleSele

	%>

    <!-- Sidebar -->
    <nav class="sidebar sidebar-offcanvas" id="sidebar">
      <ul class="nav">
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>SupervisorConfigurazioni.asp" <%=d.item("dash")%>>
            <i class="icon-grid menu-icon"<%=d.item("dashS")%>></i>
            <span class="menu-title">Dashboard</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/Ramo.asp">
            <i class="mdi mdi-barley menu-icon"></i>
            <span class="menu-title">Ramo</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" data-toggle="collapse" href="#ramo" aria-expanded="false"
            aria-controls="ramo">
            <i class="mdi mdi-alert-circle menu-icon"></i>
            <span class="menu-title">Rischi</span>
            <i class="menu-arrow"></i>
          </a>
          <div class="collapse" id="ramo">
            <ul class="nav flex-column sub-menu">
              <li class="nav-item"> <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/Caratteristica.asp">Template rischi</a></li>
              <li class="nav-item"> <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/Rischio.asp">Rischio</a></li>
            </ul>
          </div>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/Elenco.asp">
            <i class="mdi mdi-barley menu-icon"></i>
            <span class="menu-title">Elenco</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/Compagnia.asp">
            <i class="mdi mdi-account menu-icon"></i>
            <span class="menu-title">Compagnie</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/Documento.asp">
            <i class="mdi mdi-account menu-icon"></i>
            <span class="menu-title">Documenti</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/DatoTecnico.asp">
            <i class="mdi mdi-wrench menu-icon"></i>
            <span class="menu-title">Dati Tecnici</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/ListaDocumento.asp">
            <i class="mdi mdi-account menu-icon"></i>
            <span class="menu-title">Liste Documento</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/TratFisc.asp">
            <i class="mdi mdi-wrench menu-icon"></i>
            <span class="menu-title">Trattamenti Fiscali</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/Certificazione.asp">
            <i class="mdi mdi-wrench menu-icon"></i>
            <span class="menu-title">Certificazione Cauzioni</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/prodotti/ProdottoTemplate.asp">
            <i class="mdi mdi-library-books menu-icon"></i>
            <span class="menu-title">Template prodotti</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/Fornitore.asp">
            <i class="mdi mdi-truck menu-icon"></i>
            <span class="menu-title">Fornitori</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/ServizioDocumentoCoobbligato.asp">
            <i class="mdi mdi-truck menu-icon"></i>
            <span class="menu-title">Coobbligati</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/Parametri.asp">
            <i class="mdi mdi-library-books menu-icon"></i>
            <span class="menu-title">Parametri</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/prodotti/GruppoProdotti.asp">
            <i class="mdi mdi-library-books menu-icon"></i>
            <span class="menu-title">Ragg. Prodotti</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/DirittiEmissione.asp">
            <i class="mdi mdi-library-books menu-icon"></i>
            <span class="menu-title">Diritti Emissioni</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/prodotti/ProdottiAttiva.asp">
            <i class="mdi mdi-library-books menu-icon"></i>
            <span class="menu-title">Prodotti Attivi</span>
          </a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="<%=VirtualPath%>/configurazioni/Tabelle/ServizioDocumentoATI.asp">
            <i class="mdi mdi-library-books menu-icon"></i>
            <span class="menu-title">Documentazione ATI</span>
          </a>
        </li>
      </ul>
    </nav>