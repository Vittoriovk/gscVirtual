				   <% 
				   ao_row = default_row_class          'se impostato mette una row con classe 
				   ao_lbd = ""                         'descrizione label 
                   ao_lbs = "col-2"                    'size della label	
                   ao_3ld = " "                        'descrizione terzo elemento
                   ao_3ls = "col-2"                    'size terzo elemento				   
				   ao_div = "col-8"                    'se impostato crea il div con la classe 
				   ao_Obj = "Object=input|type=text"   'descrizione oggetto
				   ao_nid = ""                         'nome ed id
				   ao_val = "|value="                  'valore di default
				   ao_Att = "|attribute="              'attributi - per la lista indica il flag di mostrare o no il campo vuoto
				   if SoloLettura then
				      ao_Att = "|attribute=readonly"
				   end if 
				   ao_Plh = "|placeholder="            'placeholder - per la lista la descrizione in caso di vuoto
				   ao_Cla = "|classe=form-control"     'classe 
				   ao_Lev = "|eventoL="                'evento locale 
				   ao_Eve = "|evento="                 'evento Attiva funzione 
				   ao_Par = "|parametri="              'parametri dell'evento 
				   ao_Tex = "|Testo="                  'testo del tag 
				   ao_Tit = "|Titolo="                 'titolo
				   ao_ico = "|icona="                  'icona da aggiungere 
                   ao_ids = ""			               'valore della select 
				   ao_des = ""                         'valore del testo da mostrare 		
                   ao_NoD = ""                         'descrizione per noData per lista 	
                   ao_cal = ""                         'aggiunge calendario 				   
				   %>