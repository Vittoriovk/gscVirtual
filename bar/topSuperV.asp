<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<nav class="navbar col-lg-12 col-12 p-0 fixed-top d-flex flex-row">
  <div class="text-center navbar-brand-wrapper d-flex align-items-center justify-content-center">
      <a class="navbar-brand brand-logo mr-5" href="/gscVirtual/SupervisorConfigurazioni.asp"><img src="/gscVirtual/img/logo.png" class="mr-2"
              alt="logo" /></a>
      <a class="navbar-brand brand-logo-mini" href="/gscVirtual/SupervisorConfigurazioni.asp"><img src="/gscVirtual/img/logo_piccolo2.png"
              alt="logo" /></a>
  </div>
  <div class="navbar-menu-wrapper d-flex align-items-center justify-content-end">
      <button class="navbar-toggler navbar-toggler align-self-center" style="margin-left: -1.5rem" type="button" data-toggle="minimize">
          <span class="icon-menu"></span>
      </button>
      <div class="mx-xl-3">
        <span class="text-warning">&nbsp;Utente</span>
        <span class="text-primary"><%=Session("LoginNominativo")%></span>
        <br>
        <span class="text-warning">&nbsp;Azienda Corrente</span>
        <span class="text-primary"><%=GetDiz(session("AziendaWork"),"DescAzienda")%></span>
      </div>
      <ul class="navbar-nav navbar-nav-right">
          <li class="nav-item nav-profile dropdown">
              <a class="nav-link dropdown-toggle" href="#" data-toggle="dropdown" id="profileDropdown">
                  <img src="/gscVirtual/images/faces/face28.jpg" alt="profile" />
              </a>
              <div class="dropdown-menu dropdown-menu-right navbar-dropdown" aria-labelledby="profileDropdown">                        
                  <a class="dropdown-item" href="/gscVirtual/configurazioni/AnagraficaSuperV.asp">
                      <i class="icon-head text-primary"></i>Profilo
                  </a>
                  <a class="dropdown-item" href="/gscVirtual/SupervisorConfigurazioni.asp">
                      <i class="ti-settings text-primary"></i>Configurazioni
                  </a>
                  <a class="dropdown-item" href="/gscVirtual/logout.asp">
                      <i class="ti-power-off text-primary"></i>Esci
                  </a>
              </div>
          </li>
      </ul>
      <button class="navbar-toggler navbar-toggler-right d-lg-none align-self-center" type="button" data-toggle="offcanvas">
        <span class="icon-menu"></span>
      </button>
  </div>
</nav>
<%
Function GetDiz(D,K)
Dim RetVal
on error resume next 
	' on error resume next 
	RetVal=D.item(ucase(K))
	err.clear
	GetDiz=RetVal
End Function
%>