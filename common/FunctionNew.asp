<%
Function createDataListOld(tipo,id)
Dim Query,MySet,Inizio,CampoDb   
    Query=""
    if Tipo="COMUNE_IT" then 
       Query="Select * from Distinct ComuneIstat order By DescComune"
	   CampoDb="DescComune"
    end if 
    if Query<>"" then 
       Set MySet = Server.CreateObject("ADODB.Recordset")
       MySet.CursorLocation = 3 
       MySet.Open Query, ConnMsde
	   Do while not MySet.eof 
	      if Inizio=true then 
		     response.write "<datalist id='" & id & "'>"
		     Inizio=false 
	      end if 
		  response.write "<option>" & MySet(CampoDb) & "</option>"
	      MySet.moveNext 
       loop
	   if Inizio=false then 
	      response.write "</datalist>"
	   end if 
	   MySet.close
   end if 




end function 

Function RitornaA(pagina)
Dim rp
    rp=replace(VirtualPath & "/" & pagina ,"//","/")
	RitornaA=rp
End function 

Function isCliente()
Dim retVal
    retVal=false 
    if Session("LoginTipoUtente")=ucase("Clie") then
       retVal=true 
    end if 	
	isCliente=retVal
end function 

Function isSegnalatore()
Dim retVal
    retVal=false 
    if Session("LoginTipoUtente")=ucase("Coll") and ucase(Session("LoginTipoCollaboratore"))="SEGN" then
       retVal=true 
    end if 	
	isSegnalatore=retVal
end function 

Function isCollaboratore()
Dim retVal
    retVal=false 
    if Session("LoginTipoUtente")=ucase("Coll") then
       retVal=true 
    end if 	
	isCollaboratore=retVal
end function 

Function isBackOffice()
Dim retVal
    retVal=false 
    if Session("LoginTipoUtente")=ucase("BackO") then
       retVal=true 
    end if 	
	isBackOffice=retVal
end function

Function IsSupervisor()
Dim retVal
    retVal=false 
    if Session("LoginTipoUtente")=ucase("SuperV") then
       retVal=true 
    end if 	
	IsSupervisor=retVal
end function 

Function isAdmin()
Dim retVal
    retVal=false 
    if Session("LoginTipoUtente")=ucase("Admin") then
       retVal=true 
    end if 	
	isAdmin=retVal
end function 

function getCondForLevel(level,Idaccount)
Dim retVal
    retVal=""
    if level=1 then 
       retVal = " IdAccountLivello1 = " & idAccount
	end if 
    if level=2 then 
       retVal = " IdAccountLivello2 = " & idAccount
	end if 
    if level=3 then 
       retVal = " IdAccountLivello3 = " & idAccount
	end if 	
	getCondForLevel = retVal
end function 


Function EsisteUserId(IdAccount,UserId,Attivo)
Dim retVal,MySql,v_ret
    retVal=false 
   'controllo duplicazione 
    MySql = ""
    MySql = MySql & " select top 1 idAccount From Account "
    MySql = MySql & " where IdAccount<>" & IdAccount
    if Attivo<>"" then 
       MySql = MySql & " and   FlagAttivo='S'"  
    end if 
    MySql = MySql & " and   UserId='" & apici(UserId) & "'" 
    v_ret = LeggiCampo(MySql,"IdAccount")
	if v_ret<>"" then 
	   retVal=False 
    end if 
	EsisteUserId=retVal
end function 


Function VuotoNoLista(x)
Dim retV
	retV=x
	if x="" or X="-1" then 
	   retV=""
	end if 
	VuotoNoLista=retV
End Function

Function FormatEuro(n,d)
on error resume next 
	FormatEuro=FormatNumber(n,d,-1) & " &euro;"
	err.clear
	
End Function 

Function ListaDbChangeCompleta (Query,Name,CodValue,ColCod,ColText,FlagVuoto,Change,Campo,Larghezza,DescVuoto,DescNoData,Classe)
'Query = query da eseguire 
'Name = nome del campo da assegnare
'CodValue = valore del codice
'ColCod = colonna del codice 
'ColText = colonna della descrizione 
'Flagvuoto=1 indica che devo aggiungere la riga vuota : nel caso uso le descrizione indicate
'Change=funzione da eseguire in caso di change
'Campo=se presente crea un campo hidden 
'Larghezza imposta la size della lista

Dim MySet 
    Set MySet = Server.CreateObject("ADODB.Recordset")
    MySet.CursorLocation = 3 
    MySet.Open Query, ConnMsde

	 Stile=""
	 
	 If len(larghezza)>0 then 
	    if IsNumeric(Larghezza) then 
		    Stile= "style='width: " & Larghezza & "px'"
		 end if 
	 end if 
	 
    if MySet.EOF then
	  if FlagVuoto=1 then 
	     Response.write "<SELECT " & Stile & " " & Classe &  " name='" & Name & "'  id='" & Name & "' >"   
	  else
	     Response.write "<SELECT " & Stile & " " & Classe &  " name='" & Name & "'  id='" & Name & "' disabled>"   
	  end if 
      
	  if FlagVuoto=1 and DescNoData<>"" then 
	     Response.write "<option value=-1 selected >" & server.htmlencode(DescNoData) & "</option>"
	  else
         Response.write "<option value=-1 selected > </option>"
	  end if 
    else
	   if Change="" then 
		   Response.write "<SELECT " & Stile & " " & Classe & " name='" & Name & "'  id='" & Name & "' >"   
		else
			Response.write "<SELECT " & Stile & " " & Classe &  " onchange=""" & Change & """ name='" & Name & "' id='" & Name & "' >"   
		end if 

      if FlagVuoto=1 then 
	     if DescVuoto="" then
		    DescVuoto=" "
		 end if 
	     If Trim(CodValue)="" then	
		    Response.write "<option value=-1 selected >" & server.htmlencode(DescVuoto) & "</option>"
		 else
			Response.write "<option value=-1 >" & server.htmlencode(DescVuoto) & "</option>"
	     end if
      end if 
      OutHidden=""
	  
      WHILE NOT MySet.EOF	 	
	    Selected=""
		if ucase(trim(MySet(ColCod)))=ucase(trim(CodValue)) then 
		   Selected = " Selected "
		end if 
		if len(Campo)>1 then 
			cc="li_" & Campo & "_" & Myset(ColCod)
			OutHidden=OutHidden & "<input type=hidden name=" & cc & " Id=" & cc & " value ='" & Myset(Campo) & "'>" & chr(13)
		end if 
		
		Response.write "<option value='" & Myset(ColCod) & "' " & Selected &  " >" & server.htmlencode(Myset(ColText)) & "</option>"
        myset.MoveNext
	 WEND
    end if
    Response.write "</SELECT>"
	response.write OutHidden
    
    MySet.Close
	Set MySet = nothing
    
end function

Function OptionRangeNum(Nome,DaNum,ANum,Valore)
Dim J

	X=""
    X=X & "<select name='" & Nome & "' id='" & Nome & "'  class='form-control form-control-sm'>" 
	for j=cdbl(DaNum) to Cdbl(Anum) 
		if j=cint(Valore) then 
			Sel= " Selected "	
		else
			Sel = " "
		end if 
		X=X & "<option value=""" & J & """" & Sel & ">" & J & "</option>"
	next 
	OptionRangeNum=X & "</select> "  & chr(13)

	
End Function 

Function OptionListaValori(Nome,ListaValori,ValoreBase)
Dim J,V,T 
	on error resume next
	v=split(ListaValori,";")
	X=""
    X=X & "<select name='" & Nome & "' class='form-control form-control-sm'>" 
	t=0

	for j=0 to ubound(v)-1 step 2
	   Testo=v(j)
		Valore=V(j+1)
		if Valore=ValoreBase then 
			Sel= " Selected "	
		else
			Sel = " "
		end if 
		T=T+1
		X=X & "<option value=""" & Valore & """" & Sel & ">" & Testo & "</option>"
	next 
	OptionListaValori=X & "</select> "  & chr(13)

	err.clear
	
End Function 

Function AddHref(inRequest)
Dim arReq,arDat,i,inHref,inEvento,inParametri,inTitolo,inIcona,kk,vv
inHref     =""
inEvento   =""
inParametri=""
inTitolo   =""
inIcona    =""
arReq=split(inRequest,"|")
for i=0 to ubound(arReq)
   'response.write "iiii:" & arReq(i) & "<br>"
   arDat=split(arReq(i),"=")
   if ubound(arDat)>0 then 
      kk=ucase(arDat(0))
	  vv=arDat(1)
	  if kk="HREF" then 
	     inHref      = vv
	  elseif kk="EVENTO" then 
	     inEvento    = vv
	  elseif kk="PARAMETRI" then 
	     inParametri = vv
	  elseif kk="TITOLO" then 
	     inTitolo    = vv
	  elseif kk="ICONA" then 
	     inIcona     = vv
	  end if 
   end if 
next 

AddHref=AddHrefWithIconAndFunction(inHref,inEvento,inParametri,inTitolo,inIcona)

End Function

Function AddObject(inRequest)
Dim arReq,arDat,i,inObject,inHref,inEvento,inParametri,inTitolo,inIcona,kk,vv,InClasse
Dim inType,InEventoL,inTesto,InName,InId,inRequired,inplaceHolder,inDisabled,inHeigth,inValue
Dim outResponse,StartObj,EndObj,ClassObj,ptr,tmp
inClasse   =""
inObject   =""
inHref     =""
inEvento   =""
InEventoL  =""
inParametri=""
inTitolo   =""
inTesto    =""
inIcona    =""
inType     =""
inName     =""
inId       =""
inRequired =""
inDisabled =""
inplaceHolder=""
inHeigth   =""
inValue    =""
inTarget   =""
arReq=split(inRequest,"|")
for i=0 to ubound(arReq)
   'response.write "iiii:" & arReq(i) & "<br>"
   arDat=split(arReq(i),"=")
   if ubound(arDat)>0 then 
      kk=ucase(arDat(0))
	  vv=arDat(1)
	  if kk="OBJECT" then 
	     inObject    = vv
	  elseif kk="CLASSE" then 
	     inClasse    = vv	
	  elseif kk="TYPE"   then 
	     inType      = vv	
	  elseif kk="ID"     then 
	     inId        = vv	
	  elseif kk="NAME"   then 
	     inName      = vv	
	  elseif kk="VALUE"  then 
	     inValue      = vv	
	 elseif kk="ATTRIBUTE" then 
	     if instr(ucase(vv),"REQUIRED")>0 then 
	        inRequired  = "required"	
		 end if 
		 if instr(ucase(vv),"DISABLED")>0 or instr(ucase(vv),"READONLY")>0  then 
		    inDisabled  = "disabled"
		 end if 
	     if instr(ucase(vv),"TARGET")>0 then 
	        inTarget  = "target=""_blank"""	
		 end if 		 
		 if instr(ucase(vv),"HEIGHT")>0 then 
		    ptr=instr(ucase(vv),"HEIGHT")
			tmp=mid(vv,ptr+len("HEIGHT")+1,20)
			ptr=instr(tmp,":")
			if ptr<=0 then 
			   inHeigth = tmp
			else 
		       inHeigth = mid(tmp,1,ptr-1)
			end if   
		 end if		 
		 
	  elseif kk="HREF" then 
	     inHref      = vv
	  elseif kk="EVENTO" then 
	     inEvento    = vv
	  elseif kk="EVENTOL" then 
	     inEventoL   = vv
	  elseif kk="PARAMETRI" then 
	     inParametri = vv
	  elseif kk="TITOLO" then 
	     inTitolo    = vv
	  elseif kk="TESTO" then 
	     inTesto     = vv
  	  elseif kk="ICONA" then 
	     inIcona     = vv
  	  elseif kk="PLACEHOLDER" then 
		 inplaceHolder = vv
	  end if 
   end if 
next 
outResponse = ""
if ucase(inObject)="BUTTON" then 
   StartObj = "<button"
   ClassObj = "btn waves-effect waves-light bgcolor-white icon-site-color"
   EndObj   = "</button>"
end if 
if ucase(inObject)="A"      then 
   StartObj = "<a"
   if inTarget<>"" then 
      StartObj = StartObj & " " & inTarget
   end if 
   if inHref<>"" then 
      StartObj = StartObj & " HREF=""" & inHref & """"
   end if 
   ClassObj = ""
   EndObj   = "</a>"
end if 
if ucase(inObject)="INPUT"  then 
   StartObj = "<input"
   if inHeigth<>"" then 
      ClassObj = ClassObj & inHeigth
	  inClasse="STD"
   end if 
   if ucase(inRequired)="REQUIRED" then 
      ClassObj = ClassObj &" active validate"
      inClasse="STD"
   end if 
   EndObj   = ""
end if 

outResponse = outResponse & StartObj

if inplaceHolder<>"" then 
   outResponse = outResponse & " placeHolder=$" & inplaceHolder & "$"
end if 
if inValue<>"" then 
   outResponse = outResponse & " value=$" & inValue & "$"
end if 
if inType<>"" then 
   outResponse = outResponse & " type=$" & inType & "$"
end if 
if inName<>"" then 
   outResponse = outResponse & " name=$" & inName & "$"
end if 
if inId  <>"" then 
   outResponse = outResponse & " id=$"   & inId & "$"
end if 

if InClasse="STD" then 
   if ClassObj<>"" then 
      outResponse = outResponse & " class=""" & ClassObj & """"
   end if 
elseif InClasse<>"" then 
   outResponse = outResponse & " class=""" & InClasse & """"
end if 

outResponse = outResponse & " " & inRequired
outResponse = outResponse & " " & inDisabled 

if inEvento<>"" then 
   if ucase(inEvento)="CANCELLA" or ucase(inEvento)="DELETE" then 
	  outResponse = outResponse & "  onclick=$RemoveItem('" & inParametri & "'"
   else 
	   outResponse = outResponse & "  onclick=$AttivaFunzione('" & inEvento & "'"
	   if inParametri<>"" then 
		  outResponse = outResponse & ",'" & inParametri & "'"
	   end if 		  
   end if 
   outResponse = outResponse & ");$"  
end if 
if inEventoL<>"" then 
   outResponse = outResponse & "  onclick=$" & inEventoL & "("
   if inParametri<>"" then 
	  outResponse = outResponse & "'" & inParametri & "'"
   end if 
   outResponse = outResponse & ");$"  
end if 

if inTitolo<>"" then 
   outResponse = outResponse & "  title=$" & inTitolo & "$"
end if 
outResponse = outResponse & ">"


if inTesto<>"" then 
   outResponse = outResponse & inTesto
end if 

if inIcona<>"" then 
   outResponse = outResponse & AddImage(inIcona,inTitolo)
else 
   outResponse = outResponse & inTitolo
end if

outResponse = outResponse & EndObj

AddObject=replace(outResponse,"$","""")	

End Function 

Function AddHrefWithIconAndFunction(rif,evento,parametri,titolo,icona)
dim MyHref
    MuHref=""
    MyHRef = MyHref & "<a"
	if rif="" then 
	   MyHRef = MyHref & " href=$#$"
	else
	   MyHRef = MyHref & " href=$" & rif & "$"
	end if 
	if evento<>"" then 
	   if ucase(evento)="CANCELLA" or ucase(evento)="DELETE" then 
	      MyHRef = MyHref & "  onclick=$RemoveItem('" & parametri & "'"
	   else 
	       MyHRef = MyHref & "  onclick=$AttivaFunzione('" & evento & "'"
		   if parametri<>"" then 
			  MyHRef = MyHref & ",'" & parametri & "'"
		   end if 		  
	   end if 
	   MyHRef = MyHref & ");$"  
	end if 
	if titolo<>"" then 
	   MyHRef = MyHref & "  title=$" & titolo & "$"
	end if 
	MyHRef = MyHref & ">"
	if icona<>"" then 
	   MyHRef = MyHref & AddImage(icona,titolo)
	else 
	   MyHRef = MyHref & titolo
	end if

	MyHRef = MyHref & "</a>"
	
	AddHrefWithIconAndFunction=replace(myHref,"$","""")
End Function 
						

Function AddImage(Tipo,AltInfo)
Dim MyImg,MyIcon
    MyImg  ="<i class = ""material-icons icon-site-color "">device_unknow</i>"
	MyIcon = "device_unknow"
	Tipo=Ucase(Tipo)
	if Tipo=ucase("account") then 
	   MyIcon="account_box"
	elseif Tipo=ucase("servizi") then 
	   MyIcon  ="business_center"
	elseif Tipo=ucase("lock") then 
	   MyIcon  ="lock"
	elseif Tipo=ucase("unlock") or Tipo=ucase("lock_open") then 
	   MyIcon  ="lock_open"
    elseif Tipo=ucase("delete") then 
	   MyIcon  ="delete"	
    elseif Tipo=ucase("previous") then 
	   MyIcon  ="arrow_left"	   
    elseif Tipo=ucase("salva") or Tipo=ucase("save") then 
	   MyIcon  ="save"	   
    elseif Tipo=ucase("pdf") then 
	   MyIcon  ="picture_as_pdf"		   
    elseif Tipo=ucase("documento") then 
	   MyIcon  ="folder_shared"	   
    elseif Tipo=ucase("estratto") then 
	   MyIcon  ="description"	
    elseif Tipo=ucase("storico") or Tipo=ucase("history") then 
	   MyIcon  ="history"	 
	elseif Tipo=ucase("upload") or Tipo=ucase("cloud_upload") then 
	   MyIcon  ="cloud_upload"	   
	elseif Tipo=ucase("dettaglio") then 
	   MyIcon  ="reorder"	
    end if 
	AddImage=replace(myImg,"device_unknow",MyIcon)
end function


Function GetInfoRecordset(Dizionario,MySql)
Dim ReadRs,Esito,Campo,Valore
	on error resume next
	Esito=false

	set ReadRs = ConnMsde.execute(MySql)
	if err.number=0 then 
		if not ReadRs.eof then
		   Esito=true
		   For Each objField In ReadRs.Fields
		       xx=SetDiz(Dizionario,objField.name,ReadRs(objField.name))
			   'response.write objField.name & ":" & GetDiz(Dizionario,objField.name) &  "<br>"
		   Next

		end if
	end if 
	ReadRs.close
	err.clear

	GetInfoRecordset=Esito
	
End Function

Function SetDiz(D,K,V)
on error resume next 
	D.item(ucase(K))=V
	response.write err.description
	err.clear
End Function 

Function GetDiz(D,K)
Dim RetVal
on error resume next 
	RetVal=D.item(ucase(K))
	err.clear
	GetDiz=RetVal
End Function 

Function GetDizAsValue(D,K)
Dim RetVal
on error resume next 
	RetVal=D.item(ucase(K))
	err.clear
	GetDizAsValue=GetAsValue(RetVal)
End Function 

Function GetAsValue(T)
Dim RetVal
Dim doppio 
on error resume next 
   doppio = Server.HtmlEncode("""") 
   GetAsValue=replace(T,"""",doppio)
End Function 


function DumpDic(d,dn)
dim k
on error resume next 

	response.write "start dump == " & dn & "  ========================================" & "<br>"
	for each k in d
		response.write "     " & k & " ==>> "  & d(k) & "<br>"
	next
	response.write " end  dump == " & dn & "  ========================================" & "<br>"

end function 

function getValueOfDic(Dic,K)
Dim V,kd
	on error resume next 
	
	kd = ucase(k)
	If dic.exists(kd) then
		V = dic(kd)
	else
		V = ""
	end if
	err.clear
	getValueOfDic = V
End function 

function setValueOfDic(Dic,K,V)
	on error resume next 
	dic(ucase(k)) = V
	err.clear
	
End function


function getCurrentValueFor(id)
Dim V
    V = getValueOfDic(Pagedic,id)
	if V = "" then 
	   V = Session("swap_" & id)
	end if 
	getCurrentValueFor = V
End function 

Function AddAudit(IdRichiesta,IdAccount,IdBackOffice,Descrizione)
Dim Q 
    Q = Q & " INSERT INTO LogAudit "
    Q = Q & " (IdRichiesta , IdAccount , IdBackOffice , Descrizione ,DataAudit ,TimeAudit)"
    Q = Q & " VALUES "
    Q = Q & " (" & IdRichiesta
	Q = Q & " , " & IdAccount
    Q = Q & " , " & IdBackOffice
	Q = Q & " ,'" & Apici(Descrizione) & "'"
	Q = Q & " , " & Dtos()
	Q = Q & " , " & TimeToS()
	Q = Q & " )"

	ConnMsde.execute Q
	
	AddAudit=Err.description

End function 

Function CryptAction(action)
Dim cr 
    cr = "CR_" & CryptWithKey(action,Session("cryptPass"))
    CryptAction = cr 
end function 

Function DecryptAction(action)
Dim cr 
    cr = action 
	if mid(cr,1,3) = "CR_" then 
	   cr = DecryptWithKey(mid(action,4),Session("cryptPass"))
	end if 
    DecryptAction = cr 
end function 


Function CryptWithKey(str,chiave)
on error resume next 
Dim stringacript,i,charcript
    stringacript=""
    for i=1 to len(str)
        caratteri=Asc(mid(chiave,i,1))
        stringa=Asc(mid(str,i,1))
        charcript=caratteri Xor stringa
		if charcript<16 then 
           stringacript=stringacript & "0" & hex(charcript)
		else 
		   stringacript=stringacript & hex(charcript)
		end if 
    next
    CryptWithKey=stringacript 
	err.clear 
end function 

Function DecryptWithKey(str,chiave)
Dim i,j,stringadecript,caratteri,xx 
    stringadecript=""
	j=0
    for i=1 to Len(str) step 2
	    j=j+1
        caratteri=(Asc(mid(chiave,j,1)))
        xx=mid(str,i,2)
		yy=Chr("&H" & xx )
		'response.write xx & "::"
		'response.write yy & "::"
		'response.write Asc(Chr("&H" & xx ))
        'response.write "<br>"
        stringa=Asc(Chr("&H" & xx ))
        chardecript=caratteri Xor stringa
        stringadecript=stringadecript & Chr(chardecript)
    next
    DecryptWithKey=stringadecript
End Function

Function DecodTipoColl(TipoColl)
Dim retVal
   retVal="Agente-Broker"
   if TipoColl="IN" then 
      retVal="Intermediario E"
   elseif TipoColl="SE" then 
      retVal="Segnalatore"
   end if 
   DecodTipoColl=retVal
End function 

Function ShowLabel(T)
   response.write "<label class='form-check-label font-weight-bold'  style='font-size:11px; margin-top:0px; margin-bottom:0px;'   >" & t & "</label>"
End Function 


Function getShowInfo(I)
Dim info
   info = ""
   info = info & "<i style='color:#f8b739' class='fa fa-info-circle' data-toggle='tooltip' data-placement='top'"
   info = info & " title='" & I & "' aria-hidden='true'></i>"
   getShowInfo = info 
   
end function 

Function getShowInfo2X(I)
Dim info
   info = ""
   info = info & "<i style='color:#f8b739' class='fa fa-2x fa-info-circle' data-toggle='tooltip' data-placement='top'"
   info = info & " title='" & I & "' aria-hidden='true'></i>"
   getShowInfo2X = info 
   
end function 


Function getShowAlert(I)
Dim info
   info = ""
   info = info & "<i style='color:#f8b739' class='fa fa-info-circle' "
   info = info & " title='" & I & "' ></i>"
   getShowAlert = info 
   
end function 


Function getShowAlert2X(I)
Dim info
   info = ""
   info = info & "<i style='color:#f8b739' class='fa fa-2x fa-info-circle' "
   info = info & " title='" & I & "' ></i>"
   getShowAlert2X = info 
   
end function 

Function ShowLabelInfo(T,I)
Dim testo

   testo = ""
   testo = testo & "<label class='form-check-label font-weight-bold' style='font-size:11px; "
   testo = testo & " margin-top:0px; margin-bottom:0px;'   >" 
   if T="" then 
      testo = testo & t & " " & getShowInfo2X(I) 
      testo = testo & "</label>"
   else 
      testo = testo & t & " " & getShowInfo(I) 
      testo = testo & "</label>"
   end if 
   response.write testo
End Function 

Function ShowLabelAlert(T,I)
Dim testo

   testo = ""
   testo = testo & "<label class='form-check-label font-weight-bold' style='font-size:11px; "
   testo = testo & " margin-top:0px; margin-bottom:0px;'   >" 
   if T="" then 
      testo = testo & t & " " & getShowAlert2X(I) 
      testo = testo & "</label>"
   else 
      testo = testo & t & " " & getShowAlert(I) 
      testo = testo & "</label>"
   end if 
   response.write testo
End Function 





Function NextStatoAffidamento(act)
dim retVal
  
  if act="COMP" then 
     retVal=",'ANNU','DOCU','RIFI'"
  end if 
  if act="DOCU" then 
     retVal=",'ANNU','LAVO','COMP'"
  end if 
  if act="LAVO" then 
     retVal=",'ANNU','DOCU','COMP'"
  end if 
  if act="DOCF" then 
     retVal=",'ANNU','LAVO','COMP','DOCU'"
  end if   
  retVal="('" & act & "'" & retVal & ")"
  NextStatoAffidamento = retVal
  
  'IdStatoAffidamento	DescStatoAffidamento	FlagStatoFinale
  'AFFI	Affidato	1
  'ANNU	Annullato	1
  'COMP	Lavorazione Compagnia	0
  'DOCU	Integrazione Documentazione	0
  'DOCF Integrazione Documentazione Fornitore 0
  'LAVO	Lavorazione	0
  'RICH	Richiesta	0
  'RIFI	Rifiutato	1
  'SCAD	Scaduto	1  
End Function 

function writeDiv(nCol,label,testo,nome,attributi)
Dim attr
attr=attributi 
if nome="" then 
   attr=attr & " readonly "
else
   attr=attr & "name=""" & nome &  """ Id=""" & nome & """" & " "
end if 
response.write vbNewline
response.write "<div class=""col-" & nCol & """>" & vbNewline
response.write "   <div class=""form-group"">" & vbNewline
if label<>"" then 
response.write "       <label class=""form-check-label font-weight-bold"">" & label & "</label>" & vbNewline
end if 
response.write "       <input type=""text"" " & attr & " class=""form-control"" value=""" & testo & """ >" & vbNewline
response.write "   </div>" & vbNewline
response.write "</div>" & vbNewline
end function 


function esisteCoobbligatoxAccount (idAccountcliente,IdRichiestaAffComp)
dim Q,trov,retVal,qNot   
   q = ""
   q = q & " select top 1 * from AccountCoobbligato "
   q = q & " where IdAccount=" & IdAccountCliente 
   Trovato = "0" & LeggiCampo(q,"IdAccount")
   retVal = (Cdbl(Trovato)>0) 
   if retVal = true and cdbl(IdRichiestaAffComp)>0 then 
      qNot = ""
      qNot = qNot & " select IdAccountCoobbligato "
      qNot = qNot & " from AffidamentoRichiestaCompCoob "
      qNot = qNot & " Where IdAffidamentoRichiestaComp = " & IdRichiestaAffComp
      q = ""
      q = q & " select top 1 * from AccountCoobbligato "
      q = q & " where IdAccount=" & IdAccountCliente 
      q = q & " and IdAccountCoobbligato not in (" & qNot & ") Order By RagSoc"
      Trovato = "0" & LeggiCampo(q,"IdAccountCoobbligato")
      retVal = (Cdbl(Trovato)>0)    
   end if 
   
   esisteCoobbligatoxAccount = retVal 
end function 

function leggiNominativoAccount(IdAccount)
Dim Desc 
 Desc = LeggiCampo("Select * from Account Where IdAccount=" & IdAccount ,"Nominativo")
 leggiNominativoAccount = Desc
end function 

function LeggiAccount(tab,Id)
dim IdAcc,q
q = "select IdAccount from " & tab & " where Id" & tab & "=" & id
IdAcc = cdbl("0" & LeggiCampo(q,"IdAccount"))
LeggiAccount = idAcc
end function 


function ElencoServiziAttivi(idAcc)
Dim retVal,MySet,Query
   
   on error resume next 
   retVal = ""
   Set MySet = Server.CreateObject("ADODB.Recordset")
   Query = ""
   Query = Query & " select distinct IdAnagServizio "
   Query = Query & " from ProdottoSessione"
   Query = Query & " where IdAccount = " & idAcc
   Query = Query & "   And IdSessione = '" & Session.sessionId & "'"
   
   MySet.CursorLocation = 3 
   MySet.Open Query, ConnMsde
   
   if err.number = 0 then 
      WHILE NOT MySet.EOF
	    retVal = retVal & mySet("IdAnagServizio") & ";"
        myset.MoveNext
	 WEND
      
   end if 
   MySet.close 
   err.clear 
   ElencoServiziAttivi = retVal 
end function

function ElencoParametriAttivi(Dic,idAcc)
Dim retVal,tmpVal,MySet,Query,xx

   on error resume next 
   retVal = ""
   Set MySet = Server.CreateObject("ADODB.Recordset")
   Query = ""
   Query = Query & " select a.*, isnull(B.ValoreParametro,'') as ValPar "
   Query = Query & " from TipoParametro A left join AccountTipoParametro B"
   Query = Query & " on  B.IdAccount = " & idAcc
   Query = Query & " And A.IdTipoParametro = B.IdTipoParametro "
   MySet.CursorLocation = 3 
   MySet.Open Query, ConnMsde
   
   if err.number = 0 then 
      WHILE NOT MySet.EOF
	    retVal = mySet("IdTipoParametro")
		tmpVal = trim(mySet("ValPar"))
		if tmpVal="" then 
		   tmpVal=mySet("ValoreDefault")
		end if 
		repsonse.writ
		xx=SetDiz(Dic,retVal,tmpVal)
        myset.MoveNext
	 WEND
      
   end if 
   MySet.close 
   
   err.clear 
   ElencoParametriAttivi = "" 
end function

function SetParametroSingolo(Dic,IdTipoParametro,idAccount,IdAccountLivello1,IdAccountLivello2)
Dim retVal,tmpAcc,MySet,Query,xx

   on error resume next 
   'per questi devo prendere se esiste il valore assegnato al livello 1
   if IdTipoParametro="VAL_COB" or IdTipoParametro="VAL_ATI" or IdTipoParametro="ASS_PRO" then 
      TmpAcc = IdAccount
	  if Cdbl(IdAccountLivello1)>0 then 
	     TmpAcc = IdAccountLivello1
	  end if 
      Query = ""
      Query = Query & " select ValoreParametro "
      Query = Query & " from AccountTipoParametro "
      Query = Query & " Where IdAccount = " & TmpAcc
      Query = Query & " And IdTipoParametro = '" & IdTipoParametro & "'"
	  retVal = LeggiCampo(Query,"ValoreParametro")
	  if retVal<>"" then 
		 xx=SetDiz(Dic,IdTipoParametro,retVal)
      end if       
   end if 
  
   err.clear 
   SetParametroSingolo = "" 
end function


function isServizioAttivo(servizio)
dim retVal
   retVal=false  
   if instr(ucase(Session("Login_servizi_attivi")),ucase(servizio))>0 then
      retVal=true 
   end if 
   isServizioAttivo = retVal 
end function 


function getInfoAccount(id,campo)
dim retVal,q  
   q = "select * from Account where idAccount=" & Id
   retVal=LeggiCampo(q,campo)
   getInfoAccount=retVal
end function 

function getInfoFormazione(id,campo)
dim retVal,q  
   q = "select * from Formazione where idFormazione=" & Id
   retVal=LeggiCampo(q,campo)
   getInfoFormazione=retVal
end function 

function getInfoProdotto(IdProdotto,campo)
dim retVal,q 
   q = "" 
   q = q & " select * from Prodotto "
   q = q & " where IdProdotto=" & IdProdotto
   retVal=LeggiCampo(q,campo)

   getInfoProdotto=retVal

end function 

function getInfoStatoServizio(IdStatoServizio,campo)
dim retVal,q 
   q = "" 
   q = q & " select * from StatoServizio "
   q = q & " where IdStatoServizio='" & IdStatoServizio & "'"
   retVal=LeggiCampo(q,campo)

   getInfoStatoServizio=retVal

end function 

function getInfoProdottoFornitore(IdProdotto,IdAccountFornitore,campo)
dim retVal,q 
   q = "" 
   q = q & " select * from AccountProdotto "
   q = q & " where idAccount=" & IdAccountFornitore
   q = q & " and IdProdotto=" & IdProdotto
   retVal=LeggiCampo(q,campo)
   if ucase(campo) = ucase("CodiceProdotto") and retVal = "" then 
      retVal=getInfoProdotto(IdProdotto,campo)
   end if 
   
   getInfoProdottoFornitore=retVal

end function 

'dal più importante al meno importante 
Function getPeso(x0,x1,x2,x3,x4,x5,x6,x7,x8,x9)
Dim retVal,t
   retVal = ""
   retVal = retVal & GetPesoSingolo(x0)
   retVal = retVal & GetPesoSingolo(x1) 
   retVal = retVal & GetPesoSingolo(x2)  
   retVal = retVal & GetPesoSingolo(x3)  
   retVal = retVal & GetPesoSingolo(x4)  
   retVal = retVal & GetPesoSingolo(x5)  
   retVal = retVal & GetPesoSingolo(x6)  
   retVal = retVal & GetPesoSingolo(x7)  
   retVal = retVal & GetPesoSingolo(x8)  
   retVal = retVal & GetPesoSingolo(x9)  
   getPeso = retVal 
End function 

Function getPesoSingolo(t)
Dim retVal 
   if t="" or t="-1" or t="0" then 
      retVal="0"
   else
      retVal="1"
   end if 
   getPesoSingolo = retVal 
End function 

function CalcolaImportoPerc(ImptBase,Fisso,PercBase,ImptMini,Giorni)
Dim RappGior,ImptCalA,ImptCalc
   
   RappGior = cdbl(Giorni)/365 
   ImptCalA = Cdbl(Fisso) + (cdbl(ImptBase) * cdbl(percBase)/100)
   'response.write "<br> ImptCalcA + " & ImptCalcA & " rapp " & RappGior
   
   ImptCalc = ImptCalA * RappGior
   'response.write "<br>" & ImptCalc 
   if cdbl(ImptMini) > cdbl(ImptCalc) then 
      ImptCalc = ImptMini
   end if 
   'response.write ImptCalc 
   CalcolaImportoPerc = round(ImptCalc + 0.49,0)
end function 

function getRequestAsNum(nomeRequest)
Dim req,retV 
   req = Request(nomeRequest) 
   if trim(req) ="-1" then 
      retV = 0
   else
      retV = cdbl("0" & req)
   end if 
   getRequestAsNum = cdbl(retV) 
end function 


function getProdottiAccount()
Dim MySqlIn
   MySqlIn = ""
   MySqlIn = MySqlIn & "select IdProdotto,IdAccountFornitore"
   MySqlIn = MySqlIn & " From  ProdottoSessione "
   MySqlIn = MySqlIn & " where IdSessione = '" & Session("SessionId") & "'"
   getProdottiAccount = MySqlIn
end function 
function getProdottiAccountAll()
Dim MySqlIn
   MySqlIn = ""
   MySqlIn = MySqlIn & "select * "
   MySqlIn = MySqlIn & " From  ProdottoSessione "
   MySqlIn = MySqlIn & " where IdSessione = '" & Session("SessionId") & "'"
   getProdottiAccountAll = MySqlIn
end function 

Function PasquaGregoriana(anno)
 Dim a, b, c, p, q, r
 a = anno Mod 19: b = anno \ 100: c = anno Mod 100
 p = (19 * a + b - (b \ 4) - ((b - ((b + 8) \ 25) + 1) \ 3) + 15) Mod 30
 q = (32 + 2 * ((b Mod 4) + (c \ 4)) - p - (c Mod 4)) Mod 7
 r = (p + q - 7 * ((a + 11 * p + 22 * q) \ 451) + 114)
 PasquaGregoriana = DateSerial(anno, r \ 31, (r Mod 31) + 1)
End Function

Function GiornoFestivo(dataIn)
'si aspetta la data in formato aaaammgg o gg/mm/aaaammgg
Dim r,dt,gg,mm,aa,serialDate,dayType,pasqua 
   r=false  
   if len(dataIn)=10 then 
      gg = mid(dataIn,1,2)
	  mm = mid(dataIn,4,2)
	  aa = mid(dataIn,7,4)
   else
      gg = mid(dataIn,7,2)
	  mm = mid(dataIn,5,2)
	  aa = mid(dataIn,1,4)
   end if 
   'controllo giorni feriali 
   serialDate = DateSerial(aa,mm,gg) 
   dayType    = WEEKDAY(serialDate)
   'response.write "qq=" & dayType & " "
   'e' un giorno fra lunedi e venerdi 
   if dayType > 1 and dayType < 7 then 
      if r = false and cdbl(gg) = 1  and cdbl(mm)=  1 then 
	     r = true 
	  end if 
      if r = false and cdbl(gg) = 6  and cdbl(mm)=  1 then 
	     r = true 
	  end if 
      if r = false and cdbl(gg) = 25 and cdbl(mm)=  4 then 
	     r = true 
	  end if 
      if r = false and cdbl(gg) = 1  and cdbl(mm)=  5 then 
	     r = true 
	  end if 
      if r = false and cdbl(gg) = 2  and cdbl(mm)=  6 then 
	     r = true 
	  end if 
      if r = false and cdbl(gg) = 15 and cdbl(mm)=  8 then 
	     r = true 
	  end if 
      if r = false and cdbl(gg) = 1  and cdbl(mm)= 11 then 
	     r = true 
	  end if 
      if r = false and cdbl(gg) = 25 and cdbl(mm)= 12 then 
	     r = true 
	  end if 
      if r = false and cdbl(gg) = 26 and cdbl(mm)= 12 then 
	     r = true 
	  end if 
   else 
      r = true 
   end if
   'controllo pasqua 
   if r = false and cdbl(mm)<5 then 
      
      pasqua = PasquaGregoriana(aa)
	  
	  pasquetta = DateAdd("d", 1, pasqua)
	  'response.write "== pasqua:" & pasqua & " pasquetta:" & pasquetta & " == " & serialDate & " ## "
	  if pasquetta=serialDate then 
	     r=true
	  end if 
   end if 
   
   GiornoFestivo=r
End Function 

Function getGiornoFeriale(dataIn,giorniPrec)
'si aspetta la data in formato aaaammgg o gg/mm/aaaammgg
Dim r,dt,gg,mm,aa,serialDate,i,contaFer
   r=0 
   if len(dataIn)=10 then 
      gg = mid(dataIn,1,2)
	  mm = mid(dataIn,4,2)
	  aa = mid(dataIn,7,4)
   else
      gg = mid(dataIn,7,2)
	  mm = mid(dataIn,5,2)
	  aa = mid(dataIn,1,4)
   end if   
   serialDate = DateSerial(aa,mm,gg) 
   
   'sottraggo giorni fino a trovare un giorno feriale 
   contaFer = 0
   for i=1 to giorniPrec + 7
      serialDate = DateAdd("d", -1, serialDate)
      if GiornoFestivo(serialDate)=false then 
         contaFer = contaFer + 1
      end if   
	  'superato : cerco giorno feriale 
	  if i >= giorniPrec then 
		 if contaFer=giorniPrec then 			
	        r = year(serialDate) & right("0" & month(SerialDate),2) & right("0" & day(SerialDate),2)
		    exit for 
		 end if 
	  end if 
   
   next 
   
   getGiornoFeriale=r
   
End Function 

function getStatoFlusso()
dim q 
  q = ""
  q = q & " (select * from StatoFlusso "
  q = q & "  where IdTipoUtente='*' or IdTipoUtente like '%" & Session("LoginTipoUtente") & "%') StFl" 
  
  getStatoFlusso = q
  
End Function 
  
%>
