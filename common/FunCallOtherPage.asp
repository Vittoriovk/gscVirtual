<%
'======================================================================================================='
'== parametri in
'==   > opUrl     = url da richiamare 
'==   > opMethod  = POST o GET
'==   > opData    = dati da inviare
'==   > opType    = tipo del dato passato (non obbligatorio)
'==   > opReferer = referer (non obbligatorio)
'== parametri out
'==   > opResp    = risposta 
'==
'== la funzione ritorna vuoto se chiamata ok - altro se Ko 
'======================================================================================================='
Function CallOtherPage(opUrl,opMethod,opData,opType,opReferer,opResp)
Dim objXML,Flusso,retVal 
	on error resume next 
	Flusso = true 
	RetVal = ""
	opResp = ""
	
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	
	if Err.number<>0 then 
		err.clear
		Set objXML = Server.createObject("Microsoft.XMLHTTP")
		if Err.number<>0 then
			Flusso = false 
			RetVal = "CallOtherPage:CreateXml_" & err.description
		end if 
		err.clear
	end if 
	if Flusso = true then 
		objXML.open opMethod, opUrl , false 
		if opType="" then 
			objXML.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		else
			objXML.setRequestHeader "Content-Type",opType
		end if 
		
		if opReferer<>"" then 
			objXML.setRequestHeader "HTTP_REFERER", opReferer	
		end if 
		
		objXML.Send(opData)
		
		If objXML.Status >= 400 And objXML.Status <= 599 Then
			Flusso = false 
			RetVal = "CallOtherPage:Send_" & objXML.Status & " - " & objXML.statusText
		else
			opResp = objXML.responseText
		end if 
	end if 
	Set objXML = nothing
	err.clear

	CallOtherPage = RetVal
End Function 

Function CallOtherPageBearer(opUrl,opMethod,opData,opType,opReferer,opResp,authorization)
Dim objXML,Flusso,retVal 
	on error resume next 
	Flusso = true 
	RetVal = ""
	opResp = ""
	
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	
	if Err.number<>0 then 
		err.clear
		Set objXML = Server.createObject("Microsoft.XMLHTTP")
		if Err.number<>0 then
			Flusso = false 
			RetVal = "CallOtherPage:CreateXml_" & err.description
		end if 
		err.clear
	end if 
	if Flusso = true then 
		objXML.open opMethod, opUrl , false 
		if opType="" then 
			objXML.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		else
			objXML.setRequestHeader "Content-Type",opType
		end if 
		
		if opReferer<>"" then 
			objXML.setRequestHeader "HTTP_REFERER", opReferer
		end if 
		
		if authorization<>"" then 
		   'response.write "authorization=" & authorization & "<br> "
		   objXML.setRequestHeader "Authorization", authorization
		end if 
		
		objXML.Send(opData)
		
		If objXML.Status >= 400 And objXML.Status <= 599 Then
			Flusso = false 
			RetVal = "CallOtherPageBearer:Send_" & objXML.Status & " - " & objXML.statusText
		else
			opResp = objXML.responseText
		end if 
	end if 
	Set objXML = nothing
	err.clear

	CallOtherPageBearer = RetVal
End Function 


Function CallOtherPageB64(opUrl,opMethod,opData,opType,opReferer,opResp)
Dim objXML,Flusso,retVal 
	on error resume next 
	Flusso = true 
	RetVal = ""
	opResp = ""
	
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	if false then 
		response.write "qui:" & err.description
	end if 
	
	if Err.number<>0 then 
		err.clear
		Set objXML = Server.createObject("Microsoft.XMLHTTP")
		if Err.number<>0 then
			Flusso = false 
			RetVal = "CallOtherPage:CreateXml_" & err.description
		end if 
		err.clear
	end if 
	
	if Flusso = true then 
		objXML.open opMethod, opUrl , false 
		if opType="" then 
			objXML.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		else
			objXML.setRequestHeader "Content-Type",opType
		end if 
		
		if opReferer<>"" then 
			objXML.setRequestHeader "HTTP_REFERER", opReferer	
		end if 
		
		objXML.Send(opData)
		If objXML.Status >= 400 And objXML.Status <= 599 Then
			Flusso = false 
			RetVal = "CallOtherPageB64:Send_" & objXML.Status & " - " & objXML.statusText
		else
			opResp = encodeBase64(objXML.responseBody)
		end if 
	end if 
	Set objXML = nothing
	err.clear

	CallOtherPageB64 = RetVal
End Function 

private function encodeBase64(bytes)
  dim DM, EL
  Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  Set EL = DM.createElement("tmp")
  EL.DataType = "bin.base64"
  ' Set bytes, get encoded String
  EL.NodeTypedValue = bytes
  encodeBase64 = EL.Text
end function


function getDomain()
Dim strProtocol,servePort,strDomain
   If lcase(Request.ServerVariables("HTTPS")) = "on" Then 
      strProtocol = "https" 
   Else
      strProtocol = "http" 
   End If

  servePort = Request.ServerVariables("server_port")
  strDomain = strProtocol & "://" & Request.ServerVariables("SERVER_NAME") & ":" & servePort
  getDomain = strDomain
end function 

%>