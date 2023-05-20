<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
flagDebug=true
Oper = ""
IdCompagnia   = cdbl("0" & Request("IdCompagnia"))
idAccountCli  = cdbl("0" & Request("IdAccountCliente"))
IdFornitore   = cdbl("0" & Request("IdFornitore"))
DataRichiesta = trim(Request("DataRichiesta"))
if len(DataRichiesta)=10 then 
   DataRichiesta=DataStringa(DataRichiesta)
end if 
if IsNumeric(DataRichiesta) then 
   DataRichiesta=Cdbl(DataRichiesta)
else
   DataRichiesta=DtoS()
end if 

Giorni        = cdbl("0" & Request("Giorni")) 
Importo       = cdbl("0" & Request("Importo"))
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
   flagDebug = false 
end if 
'recupero il fornitore dall'affidamento
if Cdbl(IdFornitore=0) and cdbl(idAccountCli)>0 and cdbl(IdCompagnia)>0 and cdbl(DataRichiesta)>0 then 
   qForn = ""
   qForn = qForn & " Select * from AccountCreditoAffi"
   qForn = qForn & " where IdAccount=" & idAccountCli
   qForn = qForn & " and IdCompagnia=" & IdCompagnia
   qForn = qForn & " and IdTipoCredito='AFFI'"
   qForn = qForn & " and ValidoDal<="  & DataRichiesta
   qForn = qForn & " and ValidoAl>="   & DataRichiesta
   IdFornitore = cdbl("0" & LeggiCampo(qForn,"IdFornitore"))
   if flagDebug=true then 
      response.write qForn & "<br>"
   end if    
end if 


costoCauzione = 0
if Cdbl(IdCompagnia)>0 and Cdbl(IdFornitore)>0 and cdbl(DataRichiesta)>0 and Cdbl(Importo)>0 then
   IdAccForn  = Cdbl(LeggiCampo("select * from Fornitore Where IdFornitore=" & IdFornitore,"IdAccount")) 
   qProd      = "select * from Prodotto where IdCompagnia=" & IdCompagnia & " and IdAnagServizio='CAUZ_PROV'"
   idProdotto = Cdbl(LeggiCampo(qProd,"IdProdotto"))
   qFascia    = qFascia
   qFascia    = qFascia & " select * from AccountProdottoFascia"
   qFascia    = qFascia & " Where IdAccount=" & idAccForn
   qFascia    = qFascia & " and IdProdotto=" & idProdotto
   qFascia    = qFascia & " and IdFascia>=" & numForDb(Importo)
   qFascia    = qFascia & " order by IdFascia"
   if flagDebug=true then 
      response.write qFascia & "<br>"
   end if    
   
   Set Rs = Server.CreateObject("ADODB.Recordset")
'response.write MyContQ
   Rs.CursorLocation = 3
   Rs.Open qFascia, ConnMsde 
   if Rs.eof=false then 
      ImptBase = cdbl(Rs("CostoFisso"))
	  percBase = cdbl(Rs("percentuale"))/100 
	  ImptMini = cdbl(Rs("Minimo"))
	  RappGior = Giorni/365 
	  ImptCalA = Cdbl(ImptBase) + cdbl(Importo * percBase)
	  ImptCalc = ImptCalA * RappGior
      if flagDebug=true then
	     response.write "<br>ImptBase:" & ImptBase
		 response.write "<br>PercBase:" & percBase
		 response.write "<br>ImptMini:" & ImptMini
		 response.write "<br>RappGior:" & RappGior
		 response.write "<br>Importo :" & Importo
		 response.write "<br>ImptCalA:" & ImptCalA
		 response.write "<br>ImptCalc:" & ImptCalc
      end if 
	  if ImptMini > ImptCalc then 
	     ImptCalc = ImptMini
      end if 
	  costoCauzione = round(ImptCalc + 0.49,0)
      if flagDebug=true then
	     response.write "<br>costoCau:" & costoCauzione & "<br>"
      end if	  
   end if 
   rs.close 
end if 

response.write costoCauzione

%>