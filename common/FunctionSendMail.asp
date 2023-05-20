<!--#include file="FunMailWithAttach.asp"-->

<%
Function SendMailRichiesta(IdRichiesta)
Dim MySql,tmp_RS
Dim Nominativo,email1,email2,ClienteRichiesta,DescRichiesta,DescStato,NoteStato
Dim FromF,toAddress,CCAddress,Oggetto,Testo,TestoHTML,AttachList,Pec
on error resume next 
    Set tmp_RS = Server.CreateObject("ADODB.Recordset")
	MySql = ""
	MySql = MySql & " Select * From  Richiesta A, Account B, StatoRichiesta C"
	MySql = MySql & " Where A.IdRichiesta=" & IdRichiesta
	MySql = MySql & " And   A.IdStatoRichiesta = C.IdStatoRichiesta"
	MySql = MySql & " And   A.IdAccount = B.IdAccount "

	tmp_RS.CursorLocation = 3 
	tmp_RS.Open MySql, ConnMsde
	
	Nominativo        = tmp_RS("Nominativo")             '%Nominativo% 
    email1            = tmp_RS("email1")
	email2            = tmp_RS("email2")
    ClienteRichiesta  = tmp_RS("ClienteRichiesta")       '%Cliente%
    DescRichiesta     = tmp_RS("DescRichiesta")          '%Pratica% 
	DescStato         = tmp_RS("DescStatoRichiesta")     '%DescStato% 
    NoteStato         = tmp_RS("NoteStato")              '%AltroStato%

	tmp_RS.close 	
    
	FromF      = ""
	toAddress  = email1
	CCAddress  = email2
	Oggetto    = "Notifica variazione pratica : " & DescRichiesta & " (" & ClienteRichiesta & ")"
	Testo      = CaricaTemplate("Richiesta.txt")
	Testo      = replace(Testo,"%Nominativo%" , Nominativo)
	Testo      = replace(Testo,"%Cliente%"    , ClienteRichiesta)
	Testo      = replace(Testo,"%Pratica% "   , DescRichiesta)
	Testo      = replace(Testo,"%DescStato%"  , DescStato)
	Testo      = replace(Testo,"%AltroStato%" , NoteStato)
	
	TestoHTML  = CaricaTemplate("Richiesta.html")
	TestoHTML  = replace(TestoHTML,"%Nominativo%" , Nominativo)
	TestoHTML  = replace(TestoHTML,"%Cliente%"    , ClienteRichiesta)
	TestoHTML  = replace(TestoHTML,"%Pratica% "   , DescRichiesta)
	TestoHTML  = replace(TestoHTML,"%DescStato%"  , DescStato)
	TestoHTML  = replace(TestoHTML,"%AltroStato%" , NoteStato)
	
	AttachList = ""
	Pec        = false 
	
	if Testo<>"" or TestHtml<>"" then 
	   xx = SendMailMessageHTMLWithAttach(fromF, ToAddress, CCAddress, Oggetto, Testo, TestoHTML, AttachList, Pec)
	end if 
	
	set tmp_RS = nothing
	err.clear
	
End Function 

Function SendMailBackOffice(IdRichiesta)
Dim MySql,tmp_RS
Dim Nominativo,email1,email2,ClienteRichiesta,DescRichiesta,DescStato,NoteStato
Dim FromF,toAddress,CCAddress,Oggetto,Testo,TestoHTML,AttachList,Pec
on error resume next 

    Set tmp_RS = Server.CreateObject("ADODB.Recordset")
	MySql = ""
	MySql = MySql & " Select * From  Richiesta A, Account B, StatoRichiesta C"
	MySql = MySql & " Where A.IdRichiesta=" & IdRichiesta
	MySql = MySql & " And   A.IdStatoRichiesta = C.IdStatoRichiesta"
	MySql = MySql & " And   A.IdAccount = B.IdAccount "

	tmp_RS.CursorLocation = 3 
	tmp_RS.Open MySql, ConnMsde
	
	IdBackOffice      = tmp_RS("IdBackOffice")
    ClienteRichiesta  = tmp_RS("ClienteRichiesta")       '%Cliente%
    DescRichiesta     = tmp_RS("DescRichiesta")          '%Pratica% 
	DescStato         = tmp_RS("DescStatoRichiesta")     '%DescStato% 
    NoteStato         = tmp_RS("NoteStato")              '%AltroStato%

	tmp_RS.close 	
	
	'recupero i back_office a cui mandare la nota 
	MySql = ""
	MySql = MySql & " Select * From  Account Where IdTipoAccount = 'BackOffice' "
	   
	if Cdbl(IdBackOffice)>0 then 
	   MySql = MySql & " and IdAccount = "	& IdBackOffice
    end if 	
	tmp_RS.CursorLocation = 3 
	tmp_RS.Open MySql, ConnMsde
	email1=""
	email2=""
	do while not tmp_Rs.eof
		if tmp_RS("email1")<>"" then 
		   if email1<>"" then 
		      email1=email1 & ";"
		   end if 
		   email1 = email1 & tmp_RS("email1")
		end if 
		if tmp_RS("email2")<>"" then 
		   if email2<>"" then 
		      email2=email2 & ";"
		   end if 
		   email2 = email2 & tmp_RS("email2")
	    end if 
		tmp_RS.MoveNext
	Loop	

	tmp_RS.close 	
	
	FromF      = ""
	toAddress  = email1
	CCAddress  = email2
	Oggetto    = "Notifica variazione pratica : " & DescRichiesta & " (" & ClienteRichiesta & ")"
	Testo      = "la pratica in oggetto e' stata posta nello stato : " & DescStato & " con note:" & NoteStato
	
	TestoHTML  = Testo
	
	AttachList = ""
	Pec        = false 
	
	if Testo<>"" or TestHtml<>"" then 
	   xx = SendMailMessageHTMLWithAttach(fromF, ToAddress, CCAddress, Oggetto, Testo, TestoHTML, AttachList, Pec)
	end if 
	
	set tmp_RS = nothing
	err.clear
	
End Function 

Function CaricaTemplate(nomeTemplate)
Dim PathBase,FS,MF,template  
    on error resume next 
	template = ""

    PathBase=Server.MapPath(VirtualPath) & "/template/" & nomeTemplate
    Set fs = CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(PathBase) then 
	   set mf = fs.OpenTextFile(PathBase,1)
	   template = mf.readAll
	   mf.close 
	end if 
	
	set Fs = nothing 
	err.clear 

	CaricaTemplate=template 

End function 

%>