<%

Function getToken(u,p)
Dim opUrl,opMethod,opData,opType,opReferer,opResp
    opUrl     = "https://api.appaltando.it/API/v1/login.php"
    opMethod  = "POST"
    opData = ""
    opData = opData & "username=" & u
    opData = opData & "&password=" & p

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
Function getListaBandi(DizDatabase,token,cig,ente,oggetto,scadenzadal,scadenzaal,complessivomin,complessivomax)
Dim opUrl,opMethod,opData,opType,opReferer,opResp,cToken
on error resume next 
    conta     = 0
    opUrl     = "https://api.appaltando.it/API/v1/bandi.php"
    opMethod  = "POST"
    opData = ""
    opData = opData & "ente=" & ente 
	opData = opData & "&cig=" & cig 
	opData = opData & "&oggetto=" & oggetto
	opData = opData & "&scadenzadal=" & scadenzadal
	opData = opData & "&scadenzaal=" & scadenzaal
	opData = opData & "&complessivomin=" & complessivomin
	opData = opData & "&complessivomin=" & complessivomax	

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
	'xx=DumpDic(DizDatabase,NomePagina)
	getListaBandi = conta
end function 

'restituisce un dizionario con il bando 
Function getBando(DizDatabase,token,id)
Dim opUrl,opMethod,opData,opType,opReferer,opResp,cToken
   
    conta     = 0
    opUrl     = "https://api.appaltando.it/API/v1/bando.php"
    opMethod  = "POST"
    opData = ""
    opData = opData & "id=" & id 

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
           getBando = opResp
		   'response.write OpResp 
		   'response.write "<hr>"
           Set oJSON = New aspJSON
           oJSON.loadJSON(opResp)
		   'response.write oJSON.data("data").item("id") 

           xx=SetDiz(DizDatabase,"Id"                ,Pulisci(oJSON.data("data").item("id")))
           xx=SetDiz(DizDatabase,"cig"               ,Pulisci(oJSON.data("data").item("cig")))
           xx=SetDiz(DizDatabase,"dataimmissione"    ,Pulisci(oJSON.data("data").item("dataimmissione")))
           xx=SetDiz(DizDatabase,"ragionesociale"    ,Pulisci(oJSON.data("data").item("ragionesociale")))              
           xx=SetDiz(DizDatabase,"provincia"         ,Pulisci(oJSON.data("data").item("provincia")))
           xx=SetDiz(DizDatabase,"datapubblicazione" ,Pulisci(oJSON.data("data").item("datapubblicazione")))
           xx=SetDiz(DizDatabase,"oggetto"           ,Pulisci(oJSON.data("data").item("oggetto")))
           xx=SetDiz(DizDatabase,"categoria"         ,Pulisci(oJSON.data("data").item("categoria")))
           xx=SetDiz(DizDatabase,"importocomplessivo",Pulisci(oJSON.data("data").item("importocomplessivo")))
           xx=SetDiz(DizDatabase,"scadenzaappalto"   ,Pulisci(oJSON.data("data").item("scadenzaappalto")))
           xx=SetDiz(DizDatabase,"link"              ,Pulisci(oJSON.data("data").item("link")))
           xx=SetDiz(DizDatabase,"linkente"          ,Pulisci(oJSON.data("data").item("linkente")))
           xx=SetDiz(DizDatabase,"indirizzo"         ,Pulisci(oJSON.data("data").item("indirizzo")))
           xx=SetDiz(DizDatabase,"cap"               ,Pulisci(oJSON.data("data").item("cap")))
           xx=SetDiz(DizDatabase,"citta"             ,Pulisci(oJSON.data("data").item("citta")))
           xx=SetDiz(DizDatabase,"cod_amm"           ,Pulisci(oJSON.data("data").item("cod_amm")))
           xx=SetDiz(DizDatabase,"des_amm"           ,Pulisci(oJSON.data("data").item("des_amm")))
           xx=SetDiz(DizDatabase,"Comune"            ,Pulisci(oJSON.data("data").item("Comune")))
           xx=SetDiz(DizDatabase,"nome_resp"         ,Pulisci(oJSON.data("data").item("nome_resp")))
           xx=SetDiz(DizDatabase,"cogn_resp"         ,Pulisci(oJSON.data("data").item("cogn_resp")))
           xx=SetDiz(DizDatabase,"CapEnte"           ,Pulisci(oJSON.data("data").item("Cap")))
           xx=SetDiz(DizDatabase,"ProvinciaEnte"     ,Pulisci(oJSON.data("data").item("Provincia")))
           xx=SetDiz(DizDatabase,"Regione"           ,Pulisci(oJSON.data("data").item("Regione")))
           xx=SetDiz(DizDatabase,"sito_istituzionale",Pulisci(oJSON.data("data").item("sito_istituzionale")))
		   xx=SetDiz(DizDatabase,"mail1"             ,Pulisci(oJSON.data("data").item("mail1")))
		   xx=SetDiz(DizDatabase,"cf"                ,Pulisci(oJSON.data("data").item("Cf")))
           xx=SetDiz(DizDatabase,"ESITO","OK")
    '"Indirizzo": "Via M. Lupoli,27",
    '"titolo_resp": "DIRETTORE GENERALE",
    '"tipologia_istat": "Aziende Sanitarie Locali",
    '"tipologia_amm": "Pubbliche Amministrazioni",
    '"acronimo": null,
    '"cf_validato": "S",
    		   
        else
           xx=SetDiz(DizDatabase,"ESITO","ERR:" & xx)
        end if 
    end if 
	'xx=DumpDic(DizDatabase,"callbando")
	getBando = "OK"
end function 

%>