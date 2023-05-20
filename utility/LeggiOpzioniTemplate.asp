		    <%
			'deve ricevere
			'   optIdProdotto         = 0 per accedere per template 
			'   optIdProdottoTemplate =   per accedere per template 
			'   IdAttivita            attività di riferimento
			'   NumAttivita           
			%>
			
			<%
			locked = false
			if Readonly<>"" then 
			   locked = true
			end if 

			if cdbl(optIdProdotto)>0 then 
			   MySql=GetQueryGene("TECN",optIdProdotto,optIdAccountFornitore)
			else 
			   MySql=GetQueryGeneTemplate("TECN",optIdProdottoTemplate)
			end if 
			
            Dim ReadRs
            set ReadRs = ConnMsde.execute(MySql)
            if err.number=0 then 
			   if not ReadRs.eof then 
			      Rigo=ReadRs("Rigo")
				  response.write "<div class='row'>"  
				  optMaxCol=12 
                  Do while not ReadRs.eof 
				     ListaDati=ListaDati & ReadRs("IdOpzione") & ";" 
				     'cambio rigo
				     if Rigo<>ReadRs("Rigo") then 
					    Rigo=ReadRs("Rigo")
					    response.write "</DIV>"
						response.write "<div class='row'>"  
					 end if 
					 sizeCol=6
					 if ReadRs("Formato")="PERC" or  ReadRs("Formato")="NUMERO" then 
					    sizeCol=2
					 end if 
					 if ReadRs("Formato")="TESTO" And cdbl(ReadRs("maxLen"))<51 then 
					    sizeCol=2
					 end if 					 
					 'cambio rigo se non ho piu' colonne 
					 if cdbl(optMaxCol) < cdbl(sizeCol) then 
					    response.write "</DIV>"
						response.write "<div class='row'>"  					 
					    optMaxCol=12
					 end if 
					 optMaxCol = cdbl(optMaxCol) - cdbl(sizeCol)
					 %>
			          <div class="col-<%=sizeCol%>">
                          <div class="form-group "><%=ShowLabel(ReadRs("DescWeb"))%>
						  <%
						  nome   = "Dat_" & ReadRs("IdOpzione") 
						  valore    = getValoreOpzione(IdAttivita,NumAttivita,ReadRs("IdOpzione"),"ValoreOpzione")
						  richiesto = ReadRs("FlagObbligatorio")
						  xx=showOpzioneDato(ReadRs("IdOpzione"),nome,"0",valore,locked,richiesto)
						  %>
					      </div>
					  </div>
                     <%  
                     ReadRs.MoveNext 
                  loop
				  response.write "</div>"  
			   end if 
            end if 

			%>