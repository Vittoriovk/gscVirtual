<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows 2000 Type Library" -->
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->

<!--#include file="FunMailWithAttachSetup.asp"-->

<%
Session.Timeout = 1000
Server.ScriptTimeout = 1000

Function SendMailMessageHTMLWithAttach(fromF, ToAddress, CCAddress, Oggetto, Testo, TestoHTML, AttachList, Pec)
Dim ObjMessage,J

    On Error Resume Next
    Pec = false 'Patch per errore pec aruba

	if fromF="" then 
		fromF=SendUserName
	end if 
	
	if CCAddress="SENDER" then 
		CCAddress = fromF
	end if 
	
    Set ObjMessage = CreateObject("CDO.Message")
    
    ObjMessage.Subject = Oggetto
    ObjMessage.To = ToAddress
	ObjMessage.Cc = CCAddress
	ObjMessage.Bcc = BccUserName
	
    ObjMessage.TextBody = Testo 
    ObjMessage.HTMLBody = TestoHTML 

	if len(Trim(AttachList)) > 0 then
		Attach=split(AttachList,";")
		For J=lbound(Attach) to Ubound(Attach)
			kk=Trim(Attach(J))
			if len(kk)>0 then 
				ObjMessage.AddAttachment kk
			end if
		Next
	end if
	if Pec = true then
		if fromF="" then 
			fromF=SendUserNamePec
		end if 	
		ObjMessage.From = fromF
		'==This section provides the configuration information for the remote SMTP server.
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		'Name or IP of Remote SMTP Server
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SmtpServerPec
		'Type of authentication, NONE, Basic (Base64 encoded), NTLM
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
		'Your UserID on the SMTP server
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusername") = SendUserNamePec
		'Your password on the SMTP server
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SendPasswordPec
		'Server port (typically 25)
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SmtpServerPortPerc
		'Use SSL for the connection (False or True)
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
		'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 240
	else
		if fromF="" then 
			fromF=SendUserName
		end if 	
		ObjMessage.From = fromF
		'==This section provides the configuration information for the remote SMTP server.
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		'Name or IP of Remote SMTP Server
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SmtpServer
		'Type of authentication, NONE, Basic (Base64 encoded), NTLM
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
		'Your UserID on the SMTP server
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusername") = SendUserName
		'Your password on the SMTP server
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SendPassword
		'Server port (typically 25)
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SmtpServerPort
		'Use SSL for the connection (False or True)
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
		'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
		ObjMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 240
	end if
    ObjMessage.Configuration.Fields.Update
    '==End remote SMTP server configuration section==
    ObjMessage.Send
	'response.write err.description
	SendMailMessageHTMLWithAttach = err.description
    Err.Clear
End Function


Function CreaTestoMail(Nome,addInfo,TextHtml,TextText)
Dim TestoHtml,Testo
	TestoHtml="<html>"
	TestoHtml= TestoHtml & "<head>"
	TestoHtml= TestoHtml & "<title>" & MittenteServizio & "</title>"
	TestoHtml= TestoHtml & "<style type=""text/css"">"
	TestoHtml= TestoHtml & "a:link {text-decoration:none;}"
	TestoHtml= TestoHtml & "a:visited {text-decoration:none;}"
	TestoHtml= TestoHtml & "a:active {text-decoration:none;}"
	TestoHtml= TestoHtml & "a:hover {text-decoration:none;}"
	TestoHtml= TestoHtml & "</style>"
	TestoHtml= TestoHtml & "</head>"
	TestoHtml= TestoHtml & "<body>"
	TestoHtml= TestoHtml & "<div style='width:780px; float:auto; font-family: Verdana, Geneva, sans-serif;font-size:10.5pt;'>"
	if Nome="" then 
		TestoHtml= TestoHtml & "Gentile Cliente,<br>"
	else
		TestoHtml= TestoHtml & "Spett. " & Nome & ",<br>"
	end if 
	
	TestoHtml= TestoHtml & "la presente comunicazione contiene informazioni di tuo interesse.<br>"
	TestoHtml= TestoHtml & "<br><br>"
	if addInfo<>"" then 
		TestoHtml= TestoHtml & addInfo & "<br><br>"
	end if 
	
	TestoHtml= TestoHtml & "Se hai ricevuto questa email per errore, ti preghiamo di ignorarla.<br><br>"
	TestoHtml= TestoHtml & "Servizio Clienti " & MittenteServizio & ".<br><br>"
	if MittenteLogo<>"" then 
	   TestoHtml= TestoHtml & "<img src='data:image/jpeg;base64," & MittenteLogo & "'>"
	end if 
	TestoHtml= TestoHtml & "</body>"
	TestoHtml= TestoHtml & "</html>"
	
	Testo = ""
	if Nome="" then 
		Testo = Testo & "Gentile Cliente," & VbCrLf
	else
		Testo = Testo & "Spett. " & Nome & "," & VbCrLf
	end if 	
	Testo = Testo & "la presente comunicazione contiene informazioni di tuo interesse." & VbCrLf  & VbCrLf 
	if addInfo<>"" then 
		Testo= Testo & addInfo & VbCrLf & VbCrLf
	end if 	
	Testo = Testo & "Se hai ricevuto questa email per errore, ti preghiamo di ignorarla."  & VbCrLf  & VbCrLf
	Testo = Testo & "Servizio Clienti " & MittenteServizio & "."  & VbCrLf  & VbCrLf 

	
	TextHtml = TestoHtml
	TextText = Testo
			
End function 

function setInfoMail(mail1,mail2,mail3,send_mail,copy_mail)
dim m1,m2,m3,toM,CcM
	m1 = trim(mail1)
	m2 = trim(mail2)
	m3 = trim(mail3)

	toM=""
	ccM=""
	if m1<>"" then 
		toM = m1
		ccM = m2 
		if ccM<>"" and m3<>"" then 
			ccM=ccM & ";" & m3
		else
			ccM=m3
		end if 
	elseif m2<>"" then 
		toM = m2
		ccM = m3
	else
		toM = m3
	end if 
	send_mail = toM
	copy_mail = ccM
	'response.write "m1=" & m1 & "-m2=" & m2 & "-m3=" & m3 & "-send_mail" & send_mail & "<br>"
	setInfoMail = ""

end function 			


%>