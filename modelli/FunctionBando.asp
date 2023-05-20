<%
Function GetDizBando(IdBando,cig)
   Dim MyRs,MySql,IdStato,DescStato,K,nome
   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
   xx=SetDiz(DizDatabase,"N_IdBando",0)
   xx=SetDiz(DizDatabase,"S_CIG","")
   xx=SetDiz(DizDatabase,"S_CUG","")
   xx=SetDiz(DizDatabase,"S_CPV","")
   xx=SetDiz(DizDatabase,"S_Beneficiario","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioCF","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioPI","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioSede","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioIndirizzo","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioProvincia","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioCap","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioPec","")
   xx=SetDiz(DizDatabase,"S_BeneficiarioContatti","")
   xx=SetDiz(DizDatabase,"S_OggettoAppalto","")
   xx=SetDiz(DizDatabase,"S_TipologiaAppalto","")
   xx=SetDiz(DizDatabase,"S_Categoria","")
   xx=SetDiz(DizDatabase,"S_Classe","")
   xx=SetDiz(DizDatabase,"S_Responsabile","")
   xx=SetDiz(DizDatabase,"S_Settore","")
   xx=SetDiz(DizDatabase,"S_ModalitaRealizzazione","")
   xx=SetDiz(DizDatabase,"S_LuogoEsecuzione","")
   xx=SetDiz(DizDatabase,"S_Procedura","")
   xx=SetDiz(DizDatabase,"S_CriterioAggiudicazione","")
   xx=SetDiz(DizDatabase,"N_ImportoLotto",0)
   xx=SetDiz(DizDatabase,"N_ImportoSicurezza",0)
   xx=SetDiz(DizDatabase,"N_DataPubblicazione",0)
   xx=SetDiz(DizDatabase,"N_DataScadenza",0)
   xx=SetDiz(DizDatabase,"N_DataApertura",0)
   Set MyRs = Server.CreateObject("ADODB.Recordset")
   MySql = ""
   if Cdbl(IdBando)>0 then 
      MySql = "select * from Bando Where IdBando=" & IdBando  
   elseif CIG<>"" then 
      MySql = "select * from Bando Where CIG='" & apici(cig) & "'"  
   end if 
   if mySql <>"" then 
      MyRs.CursorLocation = 3 
      MyRs.Open MySql, ConnMsde      
      if MyRs.eof = false then 
         For Each K In DizDatabase
             nome = mid(k,3,99)
             xx=SetDiz(DizDatabase,k,MyRs(nome)) 
         next 
	  end if 
	  MyRs.close 
   end if 
   'response.write "Eccomi" & err.description  & "ddd" & IdBando

   set GetDizBando = DizDatabase
End function 

Function GetNewBando(DizDatabase)
Dim MySql,v_id,Campi,Valori
   v_id=0
   MySql = ""
   Campi = ""
   Valori= ""
   'leggo tutti i parametri ad accezione di IdBando
   For Each K In DizDatabase
       
      Valo = GetDiz(DizDatabase,K)
      Tipo = mid(k,1,2)
      nome = mid(k,3,99)
      'response.write nome 
	  if nome<>ucase("IdBando") then 
	     if Campi <> "" then 
	        Campi  = Campi & ","
	        Valori = Valori & ","
	     end if 
	     Campi  = Campi & nome 

	     if Tipo="N_" then 
		    Valori = Valori & NumForDb(valo)
	     else
		    Valori = Valori & "'" & apici(valo) & "'"
		 end if 
	  end if 		
   Next
   MySql = MySql & " Insert into Bando (" & Campi &  ") Values (" & Valori & ")"
   'response.write MySql 
   
   ConnMsde.execute MySql 
   if err.number=0 then 
      v_id=GetTableIdentity("Bando")   
   end if 
   GetNewBando=v_id
end function 


Function UpdateBando(DizDatabase)
Dim v_id , MySql , Esito , v_ret,Nome,IdAz,IdBando
on error resume next 
   Esito = ""
  MySql = ""
   
   'leggo tutti i parametri ad accezione di IdBando
   IdBando=0
   For Each K In DizDatabase
        Valo = DizDatabase.item(ucase(K))
        Tipo = mid(k,1,2)
		nome = mid(k,3,99)
		'response.write nome & "=" & valo & " " & K
		if nome<>ucase("IdBando") then 
		   if MySql <> "" then 
		      MySql = MySql & ","
		   end if 
           MySql = MySql & nome & "="
		   if Tipo="N_" then 
		      MySql = MySql & NumForDb(valo)
		   else
		      MySql = MySql & "'" & apici(valo) & "'"
		   end if 		   
		else
		   IdBando=Cdbl(Valo)
		end if 
		
   Next
   MySql = " Update Bando Set " & MySql 
   MySql = MySql & " Where IdBando = " & IdBando
   ConnMsde.Execute MySql 
   'response.write MySql 
  
   UpdateBando=Esito 
End function 

Function GetDizBandoCauzione(cig)
   Dim MyRs,MySql,K,nome
   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
   on error resume next 
   xx=SetDiz(DizDatabase,"S_TROVATO","N")
   Set MyRs = Server.CreateObject("ADODB.Recordset")
   MySql = ""
   if CIG<>"" then 
      MySql = "select top 1 * from Cauzione Where CIG='" & apici(cig) & "' order By IdCauzione Desc"  
   end if 
   if mySql <>"" then 
      MyRs.CursorLocation = 3 
      MyRs.Open MySql, ConnMsde      
      if MyRs.eof = false then 
         xx=SetDiz(DizDatabase,"S_TROVATO","S")
		 xx=RecordSetToDic(MyRs,DizDatabase)
	  end if 
	  MyRs.close 
   end if 
   set GetDizBandoCauzione = DizDatabase
End function 


%>