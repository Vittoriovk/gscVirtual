<%
Row=277
pdf.SetFont "courier","",6 
pdf.SetXY  49, Row
pdf.Cell   0,0,"EIG ltd. e' autorizzata dalla      di MALTA ad operare nel settore assicurativo nei rami:"

pdf.SetFont "courier","B",6 
pdf.SetXY  49, Row
pdf.Cell   0,0,"EIG"

pdf.SetXY  87, Row
pdf.Cell   0,0,"MFSA"

Row=Row+2
pdf.SetFont "courier","B",6 
pdf.SetXY  9, Row
pdf.Cell   0,0,"1- Infortuni; 3 – Corpi di veicoli terrestri esclusi quelli ferroviari; 10 – Responsabilità civile autoveicoli terrestri; 14 – Credito; 15 – Cauzioni."

Row=Row+2
pdf.SetFont "courier","",6 
pdf.SetXY  9, Row
pdf.Cell   0,0,"        EIG ltd.  e' ammessa all’esercizio dell’attività assicurativa  in  ITALIA nei suindicati rami in libertà di prestazione di servizi."

pdf.SetFont "courier","B",6 
pdf.SetXY  9, Row
pdf.Cell   0,0,"        EIG"
Row=Row+2
pdf.SetFont "courier","B",6 
pdf.SetXY  9, Row
pdf.Cell   0,0,"                                Codice Impresa ISVAP 40165 e  n° 00.883 Albo Imprese ISVAP -  Elenco II – P.IVA /C.F. 10176581006"
Row=Row+2
pdf.SetFont "courier","B",6 
pdf.SetXY  9, Row
pdf.Cell   0,0,"                                     Codice Fiscale del rappresentante fiscale per  l’Italia: GRR MRA  53A21 I273C"

%>