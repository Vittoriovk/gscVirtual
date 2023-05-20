<%
Dim SessionDic 
Dim Pagedic 
Dim fromInput

on error resume next 
fromInput="R"
if TypeName(o)="ClsUpload" then 
   fromInput="B"
end if 

if fromInput="R" then 
   FirstLoad=(Request("CallingPage")<>NomePagina)
else
   FirstLoad=(o.ValueOf("CallingPage")<>NomePagina)
end if 


if IsObject(Session("PERCORSO"))=false then 
	xx = DestroyCurrent()
end if 

'dizionario della pagina 
Set SessionDic  = Session("PERCORSO")
Set Pagedic     = Server.CreateObject("Scripting.Dictionary")

'dizionario dei dati della pagina : 
Set PageDatadic = Server.CreateObject("Scripting.Dictionary")

xx= GetCurrent(NomePagina)

v_tipoRicerca = getValueOfDic(Pagedic,"TipoRicerca")
v_cercatesto  = getValueOfDic(Pagedic,"cerca_testo")
v_inizia_per  = getValueOfDic(Pagedic,"inizia_per")

'scrivo i dati letti 
if fromInput="R" then 
   xx=setValueOfDic(Pagedic,"TipoRicerca",Request("TipoRicerca"))
   xx=setValueOfDic(Pagedic,"cerca_testo",Request("cerca_testo"))
   xx=setValueOfDic(Pagedic,"inizia_per" ,Request("inizia_per"))
else
   xx=setValueOfDic(Pagedic,"TipoRicerca",o.ValueOf("TipoRicerca"))
   xx=setValueOfDic(Pagedic,"cerca_testo",o.ValueOf("cerca_testo"))
   xx=setValueOfDic(Pagedic,"inizia_per" ,o.ValueOf("inizia_per"))
end if 
err.clear 

if FirstLoad = false then 
   if fromInput="R" then
      v_tipoRicerca = Request("TipoRicerca")
      v_cercatesto  = Request("cerca_testo")
      v_inizia_per  = Request("inizia_per")
   else 
      v_tipoRicerca = o.ValueOf("TipoRicerca")
      v_cercatesto  = o.ValueOf("cerca_testo")
      v_inizia_per  = o.ValueOf("inizia_per")
   end if 
end if 

Function DestroyCurrent()
	on error resume next 
	Session.Contents.Remove("PERCORSO")
	Set Session("PERCORSO") = nothing
	err.clear 
	
	Set Session("PERCORSO") = Server.CreateObject("Scripting.Dictionary")
	
End function 

'recupero dal dizionario i dati della pagina corrente se esistono e li metto nel dizionario di pagina 
Function GetCurrent(page)
Dim V,Ar,k,p, kd, vd , pagePercorco,kdp
	on error resume next 
	
	'controllo il cammino fatto recuperando la chiave "PERCORSO"
	If SessionDic.exists("PERCORSO") then
		V=SessionDic("PERCORSO")
	else
		V=""
	end if
	Ar=split(V,"||")
	p=0
	livello="00"
	for k=lbound(Ar) to uBound(Ar)	
		'pagina successiva le devo cancellare 
		kd  = Ar(k)
		kdp = mid(kd,1,len(kd)-3)
		livelloKD=Right(kd,2)
		if p=1 then 
			if livelloKD>livello then 
			   sessionDic(kdp) = ""
			end if 
		else
			livello=livelloKD
			if ucase(kdp) = ucase(page) then 
				p=1
			end if 
		end if 
	next 
	' aggiorno il percorso 
	if p=1 then 
		p=instr(V,ucase(page))
		if p>0 then 
			SessionDic("PERCORSO") = mid(SessionDic("PERCORSO"),1,p-1) & ucase(page) & "_" & livello
		end if 
	end if 
	
	If SessionDic.exists(ucase(page)) then
		v=SessionDic(ucase(page))
	else
		v=""
	end if
	
	Ar=split(V,"||")
	for k=lbound(Ar) to uBound(Ar) 
		p=instr(Ar(k),"==")
		if p>0 then 
			kd = mid(Ar(k),1,p-1)
			vd = mid(Ar(k),p+2)
			Pagedic(kd)=vd
		end if 
	next 
	
End function 

'accoda al pagina al percorso : il parametro livello se è 00 viene 
'recuperato dal processo ed incrementato di 1 
Function SetCurrent(page,livello)
Dim V 
	V=""
	for each key in Pagedic
		if V<>"" then 
			V = V & "||"
		end if 
		V = V & key & "==" & Pagedic(key) 
	next
	SessionDic(ucase(page))=V

	If SessionDic.exists("PERCORSO") then
		v=SessionDic("PERCORSO")
		'esiste il percorso prendo l'ultimo livello e lo incremento di uno 
		if livello="00" then 
			livello = right(v,2)
			'incremento di uno e aggiungo zero iniziale 
			livello = right("0" & cdbl(TestNumeroPos(livello))+1,2)
		end if 
	else
		v=""
		livello="01"
	end if
	
	if instr(v,ucase(page))=0 then 
		if v<>"" then 
			SessionDic("PERCORSO") = SessionDic("PERCORSO") & "||"
		end if 
		SessionDic("PERCORSO") = SessionDic("PERCORSO") & ucase(page) & "_" & livello
	end if 
	
	SetCurrent=""
End function 

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

Function GetValueOfPageDic(K)
	GetValueOfPageDic = getValueOfDic(PageDataDic,K)
end function 
'scarica un recorset nel dizionario della pagina 
Function RecordSetToPageDic(RecSet)
	xx = RecordSetToDic(RecSet,PageDataDic)
end function 

'scarica un recordset in un dizionario
Function RecordSetToDic(RecSet,Dic)
Dim K,V,F,XX
on error resume next 
	for each F in RecSet.fields
		K=F.Name
		V=RecSet(K)
		xx=setValueOfDic(Dic,K,V)
	next
	err.clear
	
end function 

Function RemoveSwap()
dim counter

For Each counter in Session.Contents 
	if mid(ucase(counter),1,4)="SWAP" then 
	   session(counter)=""
	end if 
Next
end function 
%>