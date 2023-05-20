<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/common/functionCF.asp"-->
<%

action  = ucase(trim(request("action"))) 
cf      = trim(request("cf"))
id      = trim(request("id"))
limite  = testNumeroPos(Request("limite"))
retVal  = ""

if action<>"" and cf<>"" then 
   if action=ucase("getDataDaCf") then 
      retVal = getDataDaCf(cf,limite)
   end if 
   if action=ucase("getDescComuneDaCF") then 
      retVal = getDescComuneDaCF(cf)
   end if 
   if action=ucase("getDescProvinciaDaCF") then 
      retVal = getDescProvinciaDaCF(cf)
   end if   
   if action=ucase("getSiglaProvinciaDaCF")    then 
      retVal = getSiglaProvinciaDaCF(cf)
   end if   
   if action=ucase("getIdStatoDaCF")    then 
      retVal = getIdStatoDaCF(cf)
   end if      
   if action=ucase("getDescStatoDaCF")    then 
      retVal = getDescStatoDaCF(cf)
   end if  
end if 
if action<>"" then   
   if action=ucase("getDescStatoDaId")    then 
      retVal = getDescStatoDaId(id)
   end if     
   if action=ucase("getDescStatoDaCodice")    then 
      retVal = getDescStatoDaCodice(id)
   end if  
   if action=ucase("getSiglaProvinciaDaProvincia")    then 
      retVal = getSiglaProvinciaDaProvincia(id)
   end if  
   
   
   if action=ucase("getCodiceCatasto") then 
      stato     = request("Stato")
	  Provincia = request("Provincia")
	  Comune    = request("Comune")
	  'response.write Stato & Provincia & Comune
      retVal = getCodiceCatasto(stato,provincia,Comune)
   end if    
end if 

response.write retVal
response.end 

%>
