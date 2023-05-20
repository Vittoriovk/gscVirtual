<%
Function GetListaTecn(DizDatabase,IdProdotto,IdAccountFornitore)
Dim v_ret,xx,MySql
on error resume next 
   v_ret = ""
   MySql = GetQueryGene("TECN",IdProdotto,IdAccountFornitore) 
   xx=GetOnlyOpzione(DizDatabase,MySql)
   err.clear 
   GetListaTecn = v_ret 
End function 

Function GetListaTecnTemplate(DizDatabase,IdProdottoTemplate)
Dim v_ret,xx,MySql
on error resume next 
   v_ret = ""
   MySql = GetQueryGeneTemplate("TECN",IdProdottoTemplate) 
   xx=GetOnlyOpzione(DizDatabase,MySql)
   err.clear 
   GetListaTecnTemplate = v_ret 
End function 


Function GetListaOpzi(DizDatabase,IdProdotto,IdAccountFornitore)
Dim v_ret,xx,MySql
on error resume next 
   v_ret = ""
   MySql = GetQueryGene("OPZI",IdProdotto,IdAccountFornitore) 
   xx=GetOnlyOpzione(DizDatabase,MySql)
   err.clear 
   GetListaOpzi = v_ret 
End function 

Function GetQueryProdOpzione(IdProdotto,IdAccountFornitore,IdOpzione)
Dim q
  q = ""
  q = q & " select * from ProdottoOpzione A, Opzione B " 
  q = q & " Where A.IdProdotto = " & IdProdotto
  q = q & " and   A.IdAccountFornitore in (0," & IdAccountFornitore & ")"
  q = q & " and   A.IdOpzione = B.IdOpzione "
  q = q & " and   A.IdOpzione = '" & IdOpzione & "'"
  GetQueryProdOpzione = q
end function 


Function GetQueryGene(IdTipo,IdProdotto,IdAccountFornitore)
Dim q
  q = ""
  q = q & " select * from ProdottoOpzione A, Opzione B " 
  q = q & " Where A.IdProdotto = " & IdProdotto
  q = q & " and   A.IdAccountFornitore in (0," & IdAccountFornitore & ")"
  q = q & " and   A.IdOpzione = B.IdOpzione "
  q = q & " and   B.IdTipoOpzione = '" & IdTipo & "'"
  q = q & " order by A.Rigo,A.Ordine"
  GetQueryGene = q
end function 

Function GetQueryGeneTemplate(IdTipo,IdProdottoTemplate)
Dim q
  q = ""
  q = q & " select * from ProdottoTemplateOpzione A, Opzione B " 
  q = q & " Where A.IdProdottoTemplate = " & IdProdottoTemplate
  q = q & " and   A.IdOpzione = B.IdOpzione "
  q = q & " and   B.IdTipoOpzione = '" & IdTipo & "'"
  q = q & " order by A.Rigo,A.Ordine"
  GetQueryGeneTemplate = q
end function 

Function GetOnlyOpzione(Dizionario,MySql)
Dim ReadRs,Esito,Campo,Valore
   on error resume next
   Esito=false

   set ReadRs = ConnMsde.execute(MySql)
   if err.number=0 then 
      Do while not ReadRs.eof 
         xx=SetDiz(Dizionario,ReadRs("IdOpzione"),ReadRs("DescWeb"))
         ReadRs.MoveNext 
      loop
   end if 
   ReadRs.close
   err.clear
   GetOnlyOpzione=Esito

End Function

function showOpzioneDato(idOpzione,id,prog,valore,locked,richiesto)
Dim ReadRs,IdTipoOpzione,IdAnagServizio,DescLista,formato,readonly,lista
   on error resume next 
   set ReadRs = ConnMsde.execute("select * from Opzione Where Idopzione='" & idOpzione & "'")
   DescLista=""
   formato  =""
    
   if err.number=0 then 
      if ReadRs.eof = false then 
         IdTipoOpzione   = ReadRs("IdTipoOpzione")
         IdAnagServizio  = ReadRs("IdAnagServizio")
         DescLista       = ReadRs("DescLista") 
         formato         = ReadRs("Formato") 
      end if 
      ReadRs.close 
   end if    
   set ReadRs = nothing 
   if ucase(formato)="TESTO" then 
      if cdbl("0" & richiesto)=1 then 
         NameLoaded= NameLoaded & ";" & id & ",TE"
      end if    
      
      if descLista="" then 
         lista=""
      else
         lista=" list='" & descLista & "' "
      end if 
      if locked then 
         readonly = " readonly "
      else
         readonly = ""
      end if 

      response.write vbNewLine
      response.write "<input type=""text"" " & readonly & lista & " name=""" & id & prog & """ id=""" & id & prog & """ class='form-control' value=""" & valore & """ >"
      if descLista <> "" then 
         xx=createDataList(descLista,descLista,"")
      end if 
      
   end if 
   if ucase(formato)="NUMERO" then 
      if cdbl("0" & richiesto)=1 then 
         NameLoaded= NameLoaded & ";" & id & ",FLO"
      else 
	     NameLoaded= NameLoaded & ";" & id & ",FL"
      end if    
      
      if descLista="" then 
         lista=""
      else
         lista=" list='" & descLista & "' "
      end if 
      if locked then 
         readonly = " readonly "
      else
         readonly = ""
      end if 

	  if valore="" then 
	     valore="0"
	  end if 
      response.write vbNewLine
      response.write "<input type=""text"" " & readonly & lista & " name=""" & id & prog & """ id=""" & id & prog & """ class='form-control' value=""" & valore & """ >"
      if descLista <> "" then 
         xx=createDataList(descLista,descLista,"")
      end if 
      
   end if   
   if ucase(formato)="PERC" then 
      NameLoaded= NameLoaded & ";" & id & ",FLO"
      if cdbl("0" & richiesto)=1 then 
         NameLoaded= NameLoaded & ";" & id & ",FLQ"
      else 
	     NameLoaded= NameLoaded & ";" & id & ",FLZ"
      end if    
      
      if descLista="" then 
         lista=""
      else
         lista=" list='" & descLista & "' "
      end if 
      if locked then 
         readonly = " readonly "
      else
         readonly = ""
      end if 
	  if valore="" then 
	     valore="0"
	  end if 

      response.write vbNewLine
      response.write "<input type=""text"" " & readonly & lista & " name=""" & id & prog & """ id=""" & id & prog & """ class='form-control' value=""" & valore & """ >"
      if descLista <> "" then 
         xx=createDataList(descLista,descLista,"")
      end if 
      
   end if   

end function 

function getCostoOpzione(IdProdotto,IdAccountFornitore,idOpzione,PrezzoRif)
Dim ReadRs,q,prezzoNum,RetNum,tmpPrezzo 

   on error resume next 
   q = "" 
   q = q & " select * from ProdottoOpzione "
   q = q & " where IdProdotto = " & IdProdotto
   q = q & " and IdAccountFornitore in (0," & IdAccountFornitore & ")"
   q = q & " and IdOpzione = '" & IdOpzione & "'"
   q = q & " order by IdAccountFornitore desc"
   
   set ReadRs = ConnMsde.execute(q)
   
   RetNum=0
   if err.number=0 then 
      if ReadRs.eof = false then 
	     RetNum = ReadRs("CostoFisso")
		 Perc   = ReadRs("PercSuAcquisto")
		 Minimo = ReadRs("CostoMinimoSuPerc")  
		 if Cdbl(Perc>0) and cdbl(PrezzoRif)>0 then 
		    tmpPrezzo = cdbl(PrezzoRif) * cdbl(Perc) / 100
			if Cdbl(tmpPrezzo) < Cdbl(Minimo) then 
			   tmpPrezzo = Minimo 
			end if 
			tmpPrezzo = round(tmpPrezzo,0)
			Retnum = cdbl(RetNum) + Cdbl(tmpPrezzo)
		 end if 
		 
      end if 
      ReadRs.close 
   end if    
   set ReadRs = nothing 
   getCostoOpzione = cdbl("0" & retNum)
   
end function    

function showOpzioneOpzi(idOpzione,id,valore,locked,PrezzoOpzione)
Dim DescLista,formato,readonly,FlagAttivo,infoPrezzo
   on error resume next 

   FlagAttivo = ""
   readonly   = ""
   if locked then 
      readonly = " disabled "
   end if 
   if valore<>"" then 
      FlagAttivo = " checked "
   end if 
   infoPrezzo = ""
   if Cdbl(PrezzoOpzione)>0 then 
      infoPrezzo = " al costo di " & insertPoint(PrezzoOpzione,2) & " &euro;"
   end if 
   response.write vbNewLine
   response.write "<input type='checkbox' " & readonly & FlagAttivo & " value='S' class='big-checkbox' "
   response.write " Id='" & id & "' name = '" & id & "' >"
   response.write "<span class='font-weight-bold'>&nbsp;Richiedi" & infoPrezzo & "</span> "
   
   response.write "<input type='hidden' value='" & PrezzoOpzione & "' Id='prezzo_" & id & "' name = 'prezzo_" & id & "' >"
   
end function 

function getValoreOpzione(IdAttivita,NumAttivita,IdOpzione,campo)
Dim retVal,q 
   RetVal=""
   on error resume next 
   
   q = ""
   q = q & " select * from AttivitaOpzione "
   q = q & " where IdAttivita='" & IdAttivita &  "'"
   q = q & " and IdNumAttivita=" & NumAttivita
   q = q & " and IdOpzione='" & IdOpzione &  "'"
   'response.write q
   retVal=Leggicampo(q,campo)
   err.clear 
   getValoreOpzione = retVal
end function 

function setValoreOpzione(IdAttivita,NumAttivita,IdOpzione,Valore,Costo)
Dim retVal,q,qu
   RetVal=""
   on error resume next 
   
   q = ""
   q = q & " select * from AttivitaOpzione "
   q = q & " where IdAttivita='" & IdAttivita &  "'"
   q = q & " and IdNumAttivita=" & NumAttivita
   q = q & " and IdOpzione='" & IdOpzione &  "'"
   
   retVal=Leggicampo(q,"IdAttivita")
   qu=""
   
   if retVal="" then 
      if Valore<>"" then 
	     qu = qu & " insert into AttivitaOpzione (IdAttivita,IdNumAttivita,IdOpzione,ValoreOpzione,CostoOpzione) "
		 qu = qu & " values ("
		 qu = qu & " '" & IdAttivita & "'" 
		 qu = qu & ", " & NumForDb(NumAttivita) 
		 qu = qu & ",'" & apici(idOpzione) & "'" 
		 qu = qu & ",'" & apici(valore) & "'" 
		 qu = qu & ", " & NumForDb(Costo) 
		 qu = qu & ")"
		 'response.write qu
		 ConnMsde.execute qu 
	  end if 
   else
      if Valore<>"" then 
	     qu = qu & " update AttivitaOpzione set "
		 qu = qu & " ValoreOpzione='" & apici(Valore) & "'" 
		 qu = qu & ",CostoOpzione = " & NumForDb(Costo) 
	  else
	     qu = qu & " delete from AttivitaOpzione "
	  end if 
      qu = qu & " where IdAttivita='" & IdAttivita &  "'"
      qu = qu & " and IdNumAttivita=" & NumAttivita
      qu = qu & " and IdOpzione='" & IdOpzione &  "'"
	  'response.write qu
	  connMsde.execute qu 
   end if 
   err.clear 
l
end function 

function RimuoviOpzioniAttivita(IdAttivita,NumAttivita)
dim qu 
    on error resume next 
    qu = qu & " delete from AttivitaOpzione "
    qu = qu & " where IdAttivita='" & IdAttivita &  "'"
    qu = qu & " and NumAttivita=" & NumAttivita
    connMsde.execute qu 
    err.clear 
end function 
%>