<!--#include virtual="/gscVirtual/common/function.asp"-->
<!--#include virtual="/gscVirtual/common/functionNew.asp"-->
<!--#include virtual="/gscVirtual/common/connDb.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAccessoDb.asp"-->
<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<!--#include virtual="/gscVirtual/common/FunMailWithAttach.asp"-->

<!--#include virtual="/gscVirtual/common/initSession.asp"-->

<%
user = trim(checkSQLsemplice(Request.Form("email")))
pass = Trim(checkSQLsemplice(Request.Form("password")))

Esito=""

if len(user)>0 and Request.Form("Oper")="Recupera" then 
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = ""
	sql = sql & "SELECT A.* FROM Account A,TipoAccount B "
	sql = sql & " WHERE userId = '" &  apici(user) & "'"
	sql = sql & " AND abilitato = 1"
	sql = sql & " AND FlagAttivo = 'S'"
	sql = sql & " AND A.IdTipoAccount = B.idTipoAccount"
	sql = sql & " AND A.IdAzienda in (0," & Session("IdAziendaMaster") & ")"
	
	rs.Open sql, ConnMsde, 1, 3
	
	Nome = ""
	pass = ""
	mail = ""
	If Not rs.EOF Then
		Session("LoginTipoUtente")=ucase(rs("IdTipoAccount"))
		Session("LoginIdAccount") =rs("IdAccount")
		Nome=rs("Nominativo")
		pass=decripta(rs("PassWord"))
		mail=rs("email1")
		rs.Close
	end if 
	if Nome<>"" and pass<>"" and mail<>"" then 
	   xx=SendMailMessageHTMLWithAttach("", mail, "", "Recupero Password ", "la sua password e' : " & pass , "la sua password e' : " & pass, "", false)
	end if 

   Esito = "La sua Richiesta è stata inviata : se l'utenza esiste ricevera' a breve nella mail di registrazione la password"
end if 

'controllo se disabilitato con messaggio 
if len(user)>0 and len(pass)>0 and Esito="" then 
	v_passw=cripta(pass)
	'response.write "passw:" & vpass & "<br>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = ""
	sql = sql & "SELECT A.* FROM Account A "
	sql = sql & " WHERE userId = '" &  apici(user) & "' And PassWord='" & Apici(v_passw) & "'"
	sql = sql & " AND abilitato = 0 "
	sql = sql & " AND A.IdAzienda in (0," & Session("IdAziendaMaster") & ")"
	'response.write sql 
	'response.end 
	Esito=LeggiCampo(sql,"DescBlocco")
end if 

if len(user)>0 and len(pass)>0 and Esito="" then 
	v_passw=cripta(pass)
	'response.write "passw:" & vpass & "<br>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = ""
	sql = sql & "SELECT A.* FROM Account A,TipoAccount B "
	sql = sql & " WHERE userId = '" &  apici(user) & "' And PassWord='" & Apici(v_passw) & "'"
	sql = sql & " AND (abilitato = 1 and flagAttivo='S')"
	sql = sql & " AND A.IdTipoAccount = B.idTipoAccount"
	sql = sql & " AND A.IdAzienda in (0," & Session("IdAziendaMaster") & ")"

	'response.write sql
	'response.end 
	
	rs.Open sql, ConnMsde, 1, 3
	
	If Not rs.EOF Then
		Session("IdAzienda")             = Rs("IdAzienda")
		Session("IdAziendaWork")         = Rs("IdAzienda")
		Session("IdSubAziende")          = GetSubAziende(Session("IdAzienda"))
		Session("LoginTipoUtente")       = ucase(rs("IdTipoAccount"))
		Session("IdProfiloAbilitazione") = Rs("IdProfiloAbilitazione")
		Session("LoginIdAccount")        = rs("IdAccount")
		Session("LoginNominativo")       = rs("Nominativo")
		Session("IdAziendaWork")         = rs("IdAzienda")
		rs.Close
		
		'carico i parametri di default 
		xx=ElencoParametriAttivi(session("Login_Parametri"),0)
		
		'default per tipo utente
		Session("sideBar_" & Session("LoginIdAccount")) = "sidebar" & Session("LoginTipoUtente") & ".asp"
		Session("topBar_"  & Session("LoginIdAccount")) = "Top"     & Session("LoginTipoUtente") & ".asp"		
		
	
		if Session("LoginTipoUtente")=ucase("SuperV") then
		    if default_show_superv_dashboard=true then 
			   Session("LoginHomePage") = VirtualPath & "bar/SuperVDashboard.asp"
			   Session("opzioneSidebar")= "dash"
			else
			   Session("LoginHomePage") = VirtualPath & "SupervisorConfigurazioni.asp"
			   Session("opzioneSidebar")= "conf"
			
			end if 
			
		elseif Session("LoginTipoUtente")=ucase("Admin") then
			Session("LoginHomePage") = VirtualPath & "bar/AdminDashboard.asp"
			Session("opzioneSidebar")= "dash"
			
		elseif Session("LoginTipoUtente")=ucase("Coll") then
			sql = ""
			sql = sql & "SELECT * FROM Collaboratore Where IdAccount = " & Session("LoginIdAccount")
			rs.Open sql, ConnMsde, 1, 3
			If Not rs.EOF Then
			    Session("LoginIdCollaboratore")   = rs("IdCollaboratore")
				Session("LoginIdAccountLev1")     = rs("IdAccountLivello1")
				Session("LoginIdAccountLev2")     = rs("IdAccountLivello2")
				session("LivelloAccount")         = rs("livello")
				Session("LoginTipoCollaboratore") = ucase(rs("IdTipoCollaboratore"))
				session("FlagGeneraCollaboratore") = Rs("FlagGeneraCollaboratore")

				Session("LoginRefAccountLev1") = rs("IdAccountLivello1")
                Session("LoginRefAccountLev2") = rs("IdAccountLivello2")
                Session("LoginRefAccountLev3") = "0"
				if session("LivelloAccount") = 1 then 
				   Session("LoginRefAccountLev1") = Session("LoginIdAccount")
				end if 
				if session("LivelloAccount") = 2 then 
				   Session("LoginRefAccountLev2") = Session("LoginIdAccount")
				end if 
				if session("LivelloAccount") = 3 then 
				   Session("LoginRefAccountLev3") = Session("LoginIdAccount")
				end if 
				xx=SetParametroSingolo(session("Login_Parametri"),"VAL_COB",Session("LoginIdAccount"),Session("LoginIdAccountLev1"),Session("LoginIdAccountLev2"))
				xx=SetParametroSingolo(session("Login_Parametri"),"VAL_ATI",Session("LoginIdAccount"),Session("LoginIdAccountLev1"),Session("LoginIdAccountLev2"))
				xx=SetParametroSingolo(session("Login_Parametri"),"ASS_PRO",Session("LoginIdAccount"),Session("LoginIdAccountLev1"),0)
 
			end if 
			
			rs.Close
			Session("LoginHomePage") = VirtualPath & "bar/CollDashboard.asp"
			Session("opzioneSidebar")= "dash"
			
		elseif Session("LoginTipoUtente")=ucase("BackO") then
			Session("LoginHomePage") = VirtualPath & "bar/BackODashboard.asp"
			Session("LoginIdUtente") = LeggiCampo("select * from Utente Where IdAccount=" & Session("LoginIdAccount"),"IdUtente")
			Session("opzioneSidebar")= "dash"
			Session("LoginIdAccountProdotti") = Session("LoginIdAccount")

		elseif Session("LoginTipoUtente")=ucase("Clie") then
		    set session("ClienteWork")    = getInfoClienteByAccount(Session("LoginIdAccount"))
		    Session("LoginIdCliente")     = GetDiz(session("ClienteWork"),"IdCliente")
			Session("LoginIdAccountLev1") = GetDiz(session("ClienteWork"),"IdAccountLivello1")
			Session("LoginIdAccountLev2") = GetDiz(session("ClienteWork"),"IdAccountLivello2")
			Session("LoginHomePage")  = VirtualPath & "bar/ClieDashboard.asp"
            Session("opzioneSidebar") = "dash"
			'carico i parametri necessari
			'per questi valgono quelli del primo livello
			xx=SetParametroSingolo(session("Login_Parametri"),"VAL_COB",Session("LoginIdAccountLev1"),0,0)
			xx=SetParametroSingolo(session("Login_Parametri"),"VAL_ATI",Session("LoginIdAccountLev1"),0,0)
            xx=SetParametroSingolo(session("Login_Parametri"),"ASS_PRO",Session("LoginIdAccountLev1"),0,0)
				
		end if 
		if Session("LoginTipoUtente")=ucase("Coll") or Session("LoginTipoUtente")=ucase("Clie") then
		   abilitato=GetDiz(session("Login_Parametri") ,"ASS_PRO")
		   if abilitato="N" then 
		      Session("LoginIdAccountProdotti")=Session("LoginIdAccountLev1")
		   else
		      Session("LoginIdAccountProdotti")=Session("LoginIdAccount")
		   end if
		   
		end if
		if instr("COLL_CLIE_BACKO",Session("LoginTipoUtente")) > 0 then
		   ConnMsde.execute "AssegnaProdottoSessione '" & Session.SessionID & "'," & Session("LoginIdAccount")
		   Session("Login_servizi_attivi")=ElencoServiziAttivi(Session("LoginIdAccount"))
		end if 
		'elenco dei servizi attivi
       
		Session("SessionId")      =Session.SessionID
		set session("AziendaWork")=getInfoAzienda(Session("IdAziendaWork"))
		
		Response.Redirect(Session("LoginHomePage"))		
        Response.End        		
	else
		rs.Close
		Set rs = Nothing
		Esito="Dati Non Validi"
	End If
End if 

If Esito<>"" then 
         %>
			<form name="FCambiaFunzione" Action="login.asp" method="post">
			  <input type="text" name="email" Id="email" value ="<%=request("email")%>">
			  <input type="text" name="password" Id="password" value ="<%=request("password")%>">
			  <input type="text" name="IdAzienda" Id="IdAzienda" value ="<%=request("IdAzienda")%>">
			  
			  <input type="text" name="esito" Id="esito" value ="<%=Esito%>">
			</form>
         <%

	    response.write "<script language=javascript>document.FCambiaFunzione.submit();</script>" 	
        response.end	
End if 

%>

