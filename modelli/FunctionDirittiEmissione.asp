<%

'ricerca i diritti per servizio. Se flagGetdiritti=true deve ricaricare i valori
'
function getDirittiEmissioneServizio(Dizionario,IdServizioRichiesto,flagGetDiritti)
Dim xx, QSel ,MyRs , Esito
Dim IdTabella, IdTabellaKey 

 on error resume next 


   Set MyRs = Server.CreateObject("ADODB.Recordset")
   MyRs.CursorLocation = 3 
   
   'bisogna utilizzare i diritti definiti 
   if flagGetDiritti=true then 
      connMsde.execute "calcolaDirittiEmissione " & IdServizioRichiesto 
   end if 
   
   'devo restituire i dati richiesti 
   Esito = false 
   qSel = ""
   qSel = qSel & " select A.* from DirittiEmissioneServizio A,ServizioRichiesto B "
   qSel = qSel & " Where B.IdServizioRichiesto = " & IdServizioRichiesto 
   qSel = qSel & " And   A.IdTabella = b.IdAttivita"
   qSel = qSel & " and   A.IdTabellaKey=b.IdNumAttivita"

   MyRs.Open qSel, ConnMsde
   'response.write qsel & err.description
   'response.write MyRs.eof      
   if MyRs.eof = false then
      Esito = true 
      For Each objField In MyRs.Fields
        xx=SetDiz(Dizionario,objField.name,MyRs(objField.name))
      Next
   end if 
   MyRs.close 
   err.clear
   getDirittiEmissioneServizio = Esito
   
end function 

function elabDirittiEmissioneServizio(IdTabella,IdTabellaKey,Dizionario,importo)
'calcolo i diritti di sistema 
Dim tipoCalcolo,imptCalcolo,fix,perc,min
Dim tmpImpt,qSel  
Dim AggiMin,AggiMax,AggiDef,AggiSel

   'diritti di emissione 
   
   tipoCalcolo = GetDiz(Dizionario,"EmisTipoCalc")
   'response.write Importo & tipoCalcolo
   fix   = cdbl("0" & GetDiz(Dizionario,"EmisSysFix"))
   perc  = cdbl("0" & GetDiz(Dizionario,"EmisSysPerc"))
   min   = cdbl("0" & GetDiz(Dizionario,"EmisSysMin"))
   imptCalcolo = getImportoCalcolo(importo,tipoCalcolo,fix,perc,min)
   xx=SetDiz(Dizionario,"EmisSysImpt",imptCalcolo)
   
   fix  = cdbl("0" & GetDiz(Dizionario,"EmisReteFix"))
   perc = cdbl("0" & GetDiz(Dizionario,"EmisRetePerc"))
   min  = cdbl("0" & GetDiz(Dizionario,"EmisReteMin"))
   imptCalcolo = getImportoCalcolo(importo,tipoCalcolo,fix,perc,min)
   xx=SetDiz(Dizionario,"EmisReteImpt",imptCalcolo)
   
   'intermediazione 
   tipoCalcolo = GetDiz(Dizionario,"InteTipoCalc")
   fix   = cdbl("0" & GetDiz(Dizionario,"InteSysFix"))
   perc  = cdbl("0" & GetDiz(Dizionario,"InteSysPerc"))
   min   = cdbl("0" & GetDiz(Dizionario,"InteSysMin"))
   imptCalcolo = getImportoCalcolo(importo,tipoCalcolo,fix,perc,min)
   xx=SetDiz(Dizionario,"InteSysImpt",imptCalcolo)

   fix   = cdbl("0" & GetDiz(Dizionario,"InteReteFix"))
   perc  = cdbl("0" & GetDiz(Dizionario,"InteRetePerc"))
   min   = cdbl("0" & GetDiz(Dizionario,"InteReteMin"))
   imptCalcolo = getImportoCalcolo(importo,tipoCalcolo,fix,perc,min)
   xx=SetDiz(Dizionario,"InteReteImpt",imptCalcolo)
   
   'intermediazione aggiuntiva 
   tipoCalcolo = GetDiz(Dizionario,"AggiTipoCalc")
   'valore selezionato 
   AggiSel = cdbl("0" & GetDiz(Dizionario,"AggiSel"))
   'valore default 
   AggiDef = cdbl("0" & GetDiz(Dizionario,"AggiDef"))
   AggiMin = cdbl("0" & GetDiz(Dizionario,"AggiMin"))
   AggiMax = cdbl("0" & GetDiz(Dizionario,"AggiMax"))
   'controllo valore selezionato 
   
   
   if ucase(tipoCalcolo)=ucase("fix") then 
      if cdbl(aggiSel)<Cdbl(AggiMin) or Cdbl(aggiSel)>cdbl(aggiMax) then 
         aggiSel=aggiDef 
      end if 	  
   else
      AggiMin = round((Cdbl(importo) * Cdbl(AggiMin) / 100) + 0.5,0)
      AggiMax = round((Cdbl(importo) * Cdbl(AggiMax) / 100) + 0.5,0)   
	  aggiDef = round((Cdbl(importo) * Cdbl(aggiDef) / 100) + 0.5,0)   
      if cdbl(aggiSel)<Cdbl(AggiMin) or Cdbl(aggiSel)>cdbl(aggiMax) then 
         aggiSel=aggiDef 
      end if 	  
   end if 
   'response.write "eccomi:" & TipoCalcolo & AggiSel & " " & AggiMin & " " & aggiMax
   AggiImpt = aggiSel

   'calcolo percentuale per sistema 
   perc  = cdbl("0" & GetDiz(Dizionario,"AggiPercSys"))    
   imptCalcolo = getImportoCalcolo(AggiImpt,"PERC",0,perc,0)
   
   xx=SetDiz(Dizionario,"AggiSel" ,aggiSel)
   xx=SetDiz(Dizionario,"AggiImpt",AggiImpt)
   xx=SetDiz(Dizionario,"AggiImptSys",imptCalcolo)
   'se passato chiave aggiorno i dati calcolati
   if IdTabella<>"" then 
      qSel = ""
      qSel = qSel & " update DirittiEmissioneServizio set "
      qSel = qSel & " EmisSysImpt="   & numForDb("0" & getValueOfDic(Dizionario,"EmisSysImpt" ))
      qSel = qSel & ",EmisReteImpt="  & numForDb("0" & getValueOfDic(Dizionario,"EmisReteImpt"))
      qSel = qSel & ",InteSysImpt="   & numForDb("0" & getValueOfDic(Dizionario,"InteSysImpt"))
      qSel = qSel & ",InteReteImpt="  & numForDb("0" & getValueOfDic(Dizionario,"InteReteImpt"))
      qSel = qSel & ",AggiSel="       & numForDb("0" & getValueOfDic(Dizionario,"AggiSel"))
      qSel = qSel & ",AggiImpt="      & numForDb("0" & getValueOfDic(Dizionario,"AggiImpt"))
	  qSel = qSel & ",AggiImptSys="   & numForDb("0" & getValueOfDic(Dizionario,"AggiImptSys"))
      qSel = qSel & " Where IdTabella='"  & apici(IdTabella)    & "'"
      qSel = qSel & " and IdTabellaKey='" & apici(IdTabellaKey) & "'"   
      'response.write "<br>" & qSel		 
      connMsde.execute qSel    
   end if 
   

end function 

function getImportoCalcolo(importo,tipoCalcolo,fisso,percentuale,minimo)
dim retImpt,tmpImpt  
   retImpt=fisso
   if Cdbl(percentuale)>0 then 
      
      'calcolo l'importo intero della percentuale 
      tmpImpt = round((Cdbl(importo) * Cdbl(percentuale) / 100) + 0.5,0)
	  'response.write "<br>Importo:" & Importo & " percentuale=" & percentuale & " calcolato=" & tmpImpt 
      'nel caso si devono sommare 
      if ucase(tipoCalcolo)=ucase("FixAndPerc") then 
         retImpt = cdbl(retImpt) + cdbl(tmpImpt)
      else 
      'caso or : prendo il piu' altro fra fisso e percentuale 
         if cdbl(retImpt) < cdbl(tmpImpt) then 
            retImpt = tmpImpt
         end if 
      end if 
   end if 
   if Cdbl(retImpt)<Cdbl(minimo) then 
      retImpt = Cdbl(minimo)
   end if    
   getImportoCalcolo = retImpt 
   
end function 

function DirittiAggiornaImptAggiSel(IdTabella,IdTabellaKey,importo)
Dim qSel 
   qSel = ""
   qSel = qSel & " update DirittiEmissioneServizio set "
   qSel = qSel & " AggiSel="       & numForDb(importo)
   qSel = qSel & ",AggiImpt="      & numForDb(importo)
   qSel = qSel & " Where IdTabella='"  & apici(IdTabella)    & "'"
   qSel = qSel & " and IdTabellaKey='" & apici(IdTabellaKey) & "'" 
   'response.write qSel 
   ConnMsde.execute QSel 
   
end function 
function DirittiGetSumDiritti(IdTabella,IdTabellaKey)
Dim qSel,retV  
   qSel = ""
   qSel = qSel & " select (EmisSysImpt + InteSysImpt + AggiImpt) as Tot from  DirittiEmissioneServizio "
   qSel = qSel & " Where IdTabella='"  & apici(IdTabella)    & "'"
   qSel = qSel & " and IdTabellaKey='" & apici(IdTabellaKey) & "'" 
   retV = cdbl("0" & LeggiCampo(qSel,"Tot"))
   DirittiGetSumDiritti = retV 
   
end function 

%>