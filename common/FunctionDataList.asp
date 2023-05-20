<%
Function createDataList(tipo,id,attr)
Dim Query,MySet,Inizio,CampoDb,tmpQ   
    Query=""

    if Tipo="STATO" then 
       Query="Select * from Stato order By DescStato"
	   CampoDb="DescStato"
    end if 	
    if Tipo="COMUNE_IT" then 
       Query="Select Distinct DescComune from ComuneIstat order By DescComune"
	   CampoDb="DescComune"
    end if 
    if Tipo="COMUNE_BYSIGLAPROV_IT" then 
       Query="Select Distinct DescComune from ComuneIstat where SiglaProvincia='" & attr & "' order By DescComune"
	   CampoDb="DescComune"
    end if 	
    if Tipo="PROVINCIA_IT" then 
       Query="Select DescProvincia from Provincia where IdStato='IT' order By DescProvincia"
	   CampoDb="DescProvincia"
    end if 	
    if Tipo="COMUNE_BYPROVINCIA_IT" then 
	   prov = ""
       prov = prov & " select IdProvincia from provincia"
	   prov = prov & " where IdProvincia='" & apici(attr) & "'" 
	   prov = prov & " or DescProvincia='" & apici(attr) & "'"
	   prov = prov & " or codiceCatasto='" & apici(attr) & "'"

       Query="Select Distinct DescComune from ComuneIstat where SiglaProvincia in (" & prov & ") order By DescComune"
	   CampoDb="DescComune"
    end if 	
	'cerco la configurazione se presente 
	if Query="" then 
	   tmpQ = "select * from DataList Where IdDataList='" & Tipo & "'"
	   'response.write tmpQ
	   Query = LeggiCampo(tmpQ,"query")
	   if Query<>"" then 
	      CampoDb = LeggiCampo("select * from DataList Where IdDataList='" & Tipo & "'","campo")
	   end if 
	end if 
	'response.write "query:" & Query
    if Query<>"" then 
       Set MySet = Server.CreateObject("ADODB.Recordset")
       MySet.CursorLocation = 3 
       MySet.Open Query, ConnMsde
	   Inizio=true
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
%>