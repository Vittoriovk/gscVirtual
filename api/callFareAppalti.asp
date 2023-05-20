<!--#include virtual="/gscVirtual/common/function.asp"-->
<!--#include virtual="/gscVirtual/common/functionNew.asp"-->
<!--#include virtual="/gscVirtual/common/connDb.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAccessoDb.asp"-->
<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<!--#include virtual="/gscVirtual/common/FunCallOtherPage.asp"-->
<!--#include virtual="/gscVirtual/api/aspjson/aspJSON1.19.asp" -->

<%
on error resume next 
'riceve in input un pagamento di un account e ne richiede il pagamento 

retV=""

token=trim(getToken())

response.write "token:" & token & "<hr>"

'aspetto 2 secondi
tempoS = second(Time())
tempoN = tempoS
do while TempoN-tempoS < 2
   tempoN = second(Time())
   if tempoN<TempoS then 
      tempoN=tempoN+60
   end if 
loop

lista=getListaBandi(token,"ferrovie")
response.write "lista 1:" & lista & "<br>"


Function getToken()
Dim opUrl,opMethod,opData,opType,opReferer,opResp
    opUrl     = "https://api.appaltando.it/API/v1/login.php"
    opMethod  = "POST"
    opData = ""
    opData = opData & "username=nextbroker"
    opData = opData & "&password=0cc80b3e231b5281dfef46b82eafbb30461cce34cc7e4c35779b1c434db2516c5de0081b77efef7691e9f2f0e952a1e4aed7e03c3b4fe066cebf3235c1049a86"

    opType    = ""
    opReferer = "http://www.nextbroker.com"
    opResp    = ""
    if AmbienteDiSviluppo="SIII" then 
       getToken = "TOKEN-SVILUPPO"
    else
        xx = CallOtherPage(opUrl,opMethod,opData,opType,opReferer,opResp)
        if xx="" then 
           if instr(ucase(opResp),ucase("Successful login"))>0 then 
              Set oJSON = New aspJSON
              oJSON.loadJSON(opResp)
              getToken = oJSON.data("jwt")
           else
              getToken = opResp
           end if 
        else
           getToken = "ERR:" & xx 
        end if 
    end if 
end function 

'restituisce un dizionario con tutte le righe 
Function getListaBandi(token,ente)
Dim opUrl,opMethod,opData,opType,opReferer,opResp,cToken
Dim DizDatabase
    Set DizDatabase = CreateObject("Scripting.Dictionary")
   
    opUrl     = "https://api.appaltando.it/API/v1/bandi.php"
    opMethod  = "POST"
    opData = ""
    opData = opData & "ente=" & ente 

    opType    = ""
    opReferer = "http://www.nextbroker.com"
    opResp    = ""
    if AmbienteDiSviluppo="SIII" then 
       getListaBandi = "BANDI-SVILUPPO"
    else
        cToken = ""
        if token<>"" then 
           cToken = "Token " & token
        end if 
        xx = CallOtherPageBearer(opUrl,opMethod,opData,opType,opReferer,opResp,cToken)
        if xx="" then 
           getListaBandi = opResp
           Set oJSON = New aspJSON
           oJSON.loadJSON(opResp)
           conta = 0 
           For Each thingy In oJSON.data("data")
               conta = conta + 1 
               Set this = oJSON.data("data").item(thingy)
               xx=SetDiz(DizDatabase,conta & "_Id",this.item("id"))
               xx=SetDiz(DizDatabase,conta & "_cig",this.item("cig"))
               xx=SetDiz(DizDatabase,conta & "_ragionesociale",this.item("ragionesociale"))               
               xx=SetDiz(DizDatabase,conta & "_provincia",this.item("provincia"))
               xx=SetDiz(DizDatabase,conta & "_datapubblicazione",this.item("datapubblicazione"))
               xx=SetDiz(DizDatabase,conta & "_oggetto",this.item("oggetto"))
               xx=SetDiz(DizDatabase,conta & "_categoria",this.item("categoria"))
               xx=SetDiz(DizDatabase,conta & "_importocomplessivo",this.item("importocomplessivo"))
               xx=SetDiz(DizDatabase,conta & "_scadenzaappalto",this.item("scadenzaappalto"))
           Next
           xx=SetDiz(DizDatabase,"ESITO","OK")
           xx=SetDiz(DizDatabase,"CONTA",conta)
        else
           xx=SetDiz(DizDatabase,"ESITO","ERR:" & xx)
        end if 
    end if 
	xx=DumpDic(DizDatabase,NomePagina)
	set getListaBandi =  DizDatabase
end function 



Function requestAction(token,action,jsonRequest)
Dim opUrl,opMethod,opData,opType,opReferer,opResp,xx
    opUrl     = getDomain() & replace(virtualPath & "/api/callActionBrokeriamo.aspx","//","/")
    opMethod  = "POST"
    opData    = "token=" & token & "&action=" & action & "&jsonRequest=" & server.urlEncode(jsonRequest)  
    opType    = ""
    opReferer = "http://www.mysite.com"
    opResp    = ""
    yy = writeTraceAttivita("CallPagBrokeriamo: send " & opUrl & "?" & opData,IdAttivita,IdNumAttivita)
    xx = CallOtherPage(opUrl,opMethod,opData,opType,opReferer,opResp)
    yy = writeTraceAttivita("CallPagBrokeriamo: rece " & xx & " - " & opResp,IdAttivita,IdNumAttivita)
    if xx="" then 
       requestAction = opResp
    else
       requestAction = "ERR:" & xx 
    end if     
end function 

%>