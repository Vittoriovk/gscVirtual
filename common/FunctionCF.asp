<%
function getMeseDaLettera(lettera)
dim retVal
   lettera=ucase(trim(lettera))
   retVal="00"
   if lettera="A" then 
      retVal="01"
   end if 
   if lettera="B" then 
      retVal="02"
   end if 
   if lettera="C" then 
      retVal="03"
   end if 
   if lettera="D" then 
      retVal="04"
   end if 
   if lettera="E" then 
      retVal="05"
   end if 
   if lettera="H" then 
      retVal="06"
   end if 
   if lettera="L" then 
      retVal="07"
   end if 
   if lettera="M" then 
      retVal="08"
   end if 
   if lettera="P" then 
      retVal="09"
   end if 
   if lettera="R" then 
      retVal="10"
   end if 
   if lettera="S" then 
      retVal="11"
   end if 
   if lettera="T" then 
      retVal="12"
   end if 
   
   getMeseDaLettera=retVal
   
end function 

function getLetteraDaMese(Mese)
dim retVal,strMesi
   strMesi="ABCDEHLMPRST"
   Mese=TestNumeroPos(trim(lettera))
   if Cdbl(Mese)>1 and Cdbl(Mese)<=12 then 
   else
      retVal=""
   end if 
   getLetteraDaMese=retVal
   
end function 
'123456789012
'gtancl65L19h703L
function getGiornoDaCF(cf)
dim retVal,strTmp,intTmp
   strTmp = mid(cf,10,2)
   retVal = 0
   if IsNumeric(strTmp) then 
      intTmp = cdbl(strTmp)
	  if cdbl(intTmp)>40 then 
	     intTmp = cdbl(intTmp) - 40
	  end if 
	  if cdbl(intTmp)>0 and cdbl(intTmp)<31 then 
	     retVal = intTmp
	  end if 
   end if 
   getGiornoDaCF=retVal
  
end function 

function getMeseDaCF(cf)
dim retVal,strTmp,intTmp
   strTmp = mid(cf,9,1)
   retVal = 0
   retVal = getMeseDaLettera(strTmp)
   getMeseDaCF=retVal
end function

'limite e' l'anno di riferimento per calcolare il secolo
'se >= mette 19 altrimenti 20
'se limite = 0 assume anno corrente 
function getAnnoDaCF(cf,limite)
dim retVal,strTmp,intTmp
   strTmp = mid(cf,7,2)
   retVal = 0
   if cdbl(limite)=0 then 
      limite=year(date()) mod 100
   end if 
   if IsNumeric(strTmp) then 
      intTmp = cdbl(strTmp)
	  if cdbl(intTmp)>=0 then 
	     if intTmp >= limite then 
		    intTmp = intTmp + 1900
		 else
		    intTmp = intTmp + 2000
		 end if 
	     retVal = intTmp
	  end if 
   end if 
   getAnnoDaCF=retVal
end function 

function getDataDaCf(cf,limite)
   dim retVal,gg,mm,aa
   retVal=""
   gg = getGiornoDaCF(cf)
   mm = getMeseDaCF(cf)
   aa = getAnnoDaCF(cf,limite)
   if cdbl(gg)>0 and cdbl(mm)>0 and cdbl(aa)>0 then 
      retVal=right("0" & gg,2) & "/" & right("0" & mm,2) & "/" & aa
   end if 
   
   getDataDaCf = retVal
end function 

function getSessoDaCF(cf)
dim retVal,strTmp,intTmp
   strTmp = mid(cf,10,2)
   retVal = ""
   if IsNumeric(strTmp) then 
      intTmp = cdbl(strTmp)
	  if cdbl(intTmp)>0 and cdbl(intTmp)<32 then 
	     retVal="M"
	  elseif cdbl(intTmp)>40 and cdbl(intTmp)<72 then 
	     retVal="F"
	  end if 
   end if 
   getSessoDaCF=retVal
  
end function 

function getDescComuneDaCF(cf)
dim retVal,q,strTmp
   strTmp = trim(mid(cf,12,4))
   retVal=""
   if strTmp<>"" then 
      q="select * from ComuneIstat Where CodiceCatasto='" & apici(strTmp) & "'"
      retVal=LeggiCampo(q,"DescComune")
   end if 
   getDescComuneDaCF=retVal
  
end function 

function getDescProvinciaDaCF(cf)
dim retVal,q,strTmp
   strTmp = trim(mid(cf,12,4))
   retVal=""
   if strTmp<>"" then 
      q="select * from ComuneIstat Where CodiceCatasto='" & apici(strTmp) & "'"
      strTmp=LeggiCampo(q,"CodiceProvincia")
	  if strTmp<>"" then 
	     q = "select * from ComuneIstat Where CodiceProvincia='" & apici(strTmp) & "' and IsCapoluogo='1'"
	     retVal = leggiCampo(q,"DescComune")
	  end if 
   end if 
   getDescProvinciaDaCF=retVal
  
end function 

function getSiglaProvinciaDaCF(cf)
dim retVal,q,strTmp
   strTmp = trim(mid(cf,12,4))
   retVal=""
   if strTmp<>"" then 
      q="select * from ComuneIstat Where CodiceCatasto='" & apici(strTmp) & "'"
      strTmp=LeggiCampo(q,"CodiceProvincia")
	  if strTmp<>"" then 
	     q = "select * from ComuneIstat Where CodiceProvincia='" & apici(strTmp) & "' and IsCapoluogo='1'"
	     retVal = leggiCampo(q,"SiglaProvincia")
	  end if 
   end if 
   getSiglaProvinciaDaCF=retVal
  
end function 
function getSiglaProvinciaDaProvincia(prov)
dim retVal,q,strTmp
   strTmp = trim(prov)
   retVal=""
   if strTmp<>"" then 
      q = ""
      q = q & " select * from Provincia Where "
	  q = q & "    IdProvincia ='" & apici(strTmp) & "'"
	  q = q & " or DescProvincia ='" & apici(strTmp) & "'"
	  q = q & " or codiceCatasto ='" & apici(strTmp) & "'" 
      retVal=LeggiCampo(q,"IdProvincia")
   end if 
   getSiglaProvinciaDaProvincia=retVal
  
end function 



function getCodeProvinciaDaCF(cf)
dim retVal,q,strTmp
   strTmp = trim(mid(cf,12,4))
   retVal=""
   if strTmp<>"" then 
      q="select * from ComuneIstat Where CodiceCatasto='" & apici(strTmp) & "'"
      strTmp=LeggiCampo(q,"CodiceProvincia")
	  if strTmp<>"" then 
	     q = "select * from ComuneIstat Where CodiceProvincia='" & apici(strTmp) & "' and IsCapoluogo='1'"
	     retVal = leggiCampo(q,"CodiceCatasto")
	  end if 
   end if 
   getCodeProvinciaDaCF=retVal
  
end function

function getIdStatoDaCF(cf)
dim retVal,q,strTmp
   strTmp = trim(mid(cf,12,4))
   retVal=""
   if strTmp<>"" then 
      if mid(ucase(StrTmp),1,1)="Z" then 
	     q="select * from Stato Where CodiceCatasto='" & apici(strTmp) & "'" 
		 retVal=LeggiCampo(q,"IdStato")
	  else 
	     retVal="IT"
      end if   
   end if 
   getIdStatoDaCF=retVal
  
end function

function getDescStatoDaCf(cf)
dim retVal,q
   retVal=""
   strTmp = getIdStatoDaCF(cf)
   if strTmp<>"" then 
      retVal=getDescStatoDaId(strTmp)
   end if 
   getDescStatoDaCf=retVal
  
end function

function getDescStatoDaId(idStato)
dim retVal,q,strTmp
   retVal=""
   q="select * from Stato Where IdStato='" & apici(idStato) & "'"
   retVal=LeggiCampo(q,"DescStato")
   getDescStatoDaId=retVal
  
end function

function getDescStatoDaCodice(codiceIstat)
dim retVal,q,strTmp
   retVal=""
   q="select * from Stato Where CodiceCatasto='" & apici(codiceIstat) & "'"
   retVal=LeggiCampo(q,"DescStato")
   getDescStatoDaCodice=retVal
  
end function

function getIdStatoDaCodice(codiceIstat)
dim retVal,q,strTmp
   retVal=""
   q="select * from Stato Where CodiceCatasto='" & apici(codiceIstat) & "'"
   retVal=LeggiCampo(q,"IdStato")
   getIdStatoDaCodice=retVal
  
end function

function getCodiceCatasto(stato,provincia,comune)
dim retVal,q,strTmp,com,pro
  strTmp=ucase(trim(Stato))
  if trim(strTmp)<>"" then 
     com = trim(comune)
	 pro = trim(provincia)  
     if strTmp="IT" or strTmp="ITA" or strTmp="ITALIA" then 
	    retVal = getCodiceCatastoComune("IT",pro,com)
	 elseif com<>"" and pro<>"" then
	    retVal = getCodiceCatastoStato(strTmp)
	 end if 
  
  end if 
  getCodiceCatasto = retVal
end function 

function getCodiceCatastoStato(stato)
dim retVal,q,strTmp
  strTmp=ucase(trim(Stato))
  if trim(strTmp)<>"" then
     q = ""   
     q = q & " select * from Stato where "
	 q = q & "    IdStato='"       & apici(strTmp) & "'"
     q = q & " or IdStatoEsteso='" & apici(strTmp) & "'"
     q = q & " or DescStato='"     & apici(strTmp) & "'"
     retVal=LeggiCampo(q,"CodiceCatasto")
  end if 
  
  getCodiceCatastoStato = retVal
end function 

function getCodiceCatastoComune(stato,provincia,comune)
dim retVal,q,strTmp,prov
  strTmp=ucase(trim(provincia))
  if trim(strTmp)<>"" then
     prov=getSiglaProvincia(stato,strTmp)
     if prov<>"" then 
        q = ""   
        q = q & " select * from ComuneIstat where "
	    q = q & "    SiglaProvincia='" & apici(prov) & "'"
        q = q & " and DescComune='"   & apici(comune) & "'"
        retVal=LeggiCampo(q,"CodiceCatasto")
     end if 
  end if 
  getCodiceCatastoComune = retVal
end function 

function getSiglaProvincia(stato,provincia)
dim retVal,q,strTmp,prov
  strTmp=ucase(trim(provincia))
  if trim(strTmp)<>"" then
     q = ""   
     q = q & " select * from Provincia where "
	 q = q & "    IdStato = '" & apici(stato) & "' and ("
	 q = q & "    DescProvincia='" & apici(strTmp) & "'"
     q = q & " or IdProvincia='"   & apici(strTmp) & "'"
     q = q & " or CodiceCatasto='" & apici(strTmp) & "'"
	 q = q & " )"
     retVal=LeggiCampo(q,"IdProvincia")
  end if 
  getSiglaProvincia = retVal
end function 



%>