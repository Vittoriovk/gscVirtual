<!--#include virtual="/gscVirtual/common/parametri.asp"-->
      <nav class="navbar navbar-expand-lg navbar-light <%=default_nav_bar_bg_color%> " style="margin-bottom: 3px;">
        <button class="btn btn-primary" id="menu-toggle"><span class="navbar-toggler-icon"></span></button>

        <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarSupportedContent" aria-controls="navbarSupportedContent" aria-expanded="false" aria-label="Toggle navigation">
          <span class="navbar-toggler-icon"></span>
        </button>
		<div>
			<span class="text-warning" >&nbsp;Utente (Admin) </span>
			<span class="<%=default_nav_span_text_color%>" ><%=Session("LoginNominativo")%></span>
			<br>
			<span class="text-warning" >&nbsp;Azienda Corrente</span>
			<span class="<%=default_nav_span_text_color%>" ><%=GetDiz(session("AziendaWork"),"DescAzienda")%></span>
		</div>
        <div class="collapse navbar-collapse" id="navbarSupportedContent">
          <ul class="navbar-nav ml-auto mt-2 mt-lg-0">
            <li class="nav-item active">
              <a class="nav-link <%=default_nav_span_text_color%>" href="/gscVirtual/bar/AdminDashboard.asp"><i class="fa fa-home">&nbsp;</i>Home <span class="sr-only">(current)</span></a>
            </li>
            <li class="nav-item">
              <a class="nav-link <%=default_nav_span_text_color%>" href="/gscVirtual/configurazioni/AnagraficaSuperV.asp"><i class="fa fa-user">&nbsp;</i>Profilo</a>
            </li>
            <li class="nav-item">
              <a class="nav-link <%=default_nav_span_text_color%>" href="/gscVirtual/link/AdminIntermediari.asp"><i class="fa fa-users">&nbsp;</i>Intermediari</a>
            </li>			
            <li class="nav-item">
              <a class="nav-link <%=default_nav_span_text_color%>" href="/gscVirtual/link/AdminOperatori.asp"><i class="fa fa-user">&nbsp;</i>Operatori</a>
            </li>			
            <li class="nav-item">
              <a class="nav-link <%=default_nav_span_text_color%>" href="/gscVirtual/link/AdminProvvigioni.asp"><i class="fa fa-percent">&nbsp;</i>Provvigioni</a>
            </li>			
            <li class="nav-item">
              <a class="nav-link <%=default_nav_span_text_color%>" href="/gscVirtual/logout.asp"><i class="fa fa-sign-out">&nbsp;</i>Esci</a>
            </li>			
          </ul>
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