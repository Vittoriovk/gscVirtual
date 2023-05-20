<%

function TrasformaInLettere(byval n)
	dim i, dec, lun, lett
	dim tri(5)
	n=cdbl(n)'Il tipo double è il più capiente
	dec=Right(formatnumber(n,2),2)'estrae i due decimali
	num=cstr(fix(n))

	Str=cstr("000000000000000000000"&num)
	lun=len(Str)
	for i=1 to 5
		'seziona la stringa 
		tri(i)=mid(Str,lun-(i*3)+1,3)
		'Il numero ora è composto da terzine tri(5)& tri(4)& tri(3)& tri(2)& tri(1).
		if tri(i)>0 then lett= TrasformaTerzina(tri(i),i) & lett
	next 
	if lett="" then lett="zero"
	TrasformaInLettere= lett&"/"&dec
end function

function TrasformaTerzina(tri,t)
'tri è la terzina "nnn" da convertire.
't è il numero della terzina nell'ordine:
't=1 va da 1 a 999
't=2 va da 1000 a 999000, quindi sono le migliaia
't=3 sono i milioni, ecc
	dim ultimeduecifre,ris,decine,unita
	dim n(30) 'tutte le cifre in lettere
	n(0)=""
	n(1)="uno":n(2)="due":n(3)="tre":n(4)="quattro":n(5)="cinque"
	n(6)="sei":n(7)="sette":n(8)="otto":n(9)="nove":n(10)="dieci"
	n(11)="undici":n(12)="dodici":n(13)="tredici":n(14)="quattordici":n(15)="quindici"
	n(16)="sedici":n(17)="diciassette":n(18)="diciotto":n(19)="diciannove"
	n(22)="venti":n(23)="trenta":n(24)="quaranta":n(25)="cinquanta"
	n(26)="sessanta":n(27)="settanta":n(28)="ottanta":n(29)="novanta"
	
	'gestisce le centinaia
	If left(tri,1)=1 then ris="cento" 'se è 1xx allora è cento
      if left(tri,1)>1 then ris=n(cint(left(tri,1)))&"cento"
      'se è nxx con n tra 2 e 9 allora usa n+"cento"
	
	'gestisce le decine: da 1 a 19 e da venti in poi
	Ultimeduecifre=cint(right(tri,2))
	if Ultimeduecifre < 20 then 
		ris=ris&n(Ultimeduecifre)  'da uno a diciannove: prende i numeri già pronti.
	else
		'gestisce separatamente decine ed unità per valori da venti in poi
		Decine=mid(tri,2,1)
		Unita=right(tri,1)
		ris=ris & n(20+decine) 'da n(22) è venti, n(23) è trenta,ecc.
		if Unita=1 or Unita=8 then ris=left(ris,len(ris)-1)
              'toglie l'ultima lettera per evitare "trentAuno"
		Ris=ris & n(Unita) 'aggiunge uno, due , tre
	end if	
	
	'gestisce la posizione della terzina nel numero (centinaia, migliaia, ecc.
	if t=2 then
		ris=ris&"mila"
		if tri=1 then ris="mille"
	end if
	if t=3 then
		ris=ris&"milioni"
		if tri=1 then ris="unmilione"
	end if
	if t=4 then
		ris=ris&"miliardi"
		if tri=1 then ris="unmiliardo"
	end if
	if t=5 then
		ris=ris&"milamiliardi"
		if tri=1 then ris="millemiliardi"
	end if
	
	TrasformaTerzina=Ris
end function
%>