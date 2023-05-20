<%
Function GetDizCredito(IdAccount)
   Dim MyRs,MySql,IdStato,DescStato,K,nome
   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
   xx=SetDiz(DizDatabase,"N_IdAccount",0)
   xx=SetDiz(DizDatabase,"S_IdTipoCredito","")
   xx=SetDiz(DizDatabase,"S_DescTipoCredito","")
   xx=SetDiz(DizDatabase,"N_ImptTipoCredito",0)
   Set MyRs = Server.CreateObject("ADODB.Recordset")
   MySql = "Select * from TipoCredito"
   MyRs.CursorLocation = 3 
   MyRs.Open MySql, ConnMsde      
   if MyRs.eof = false then 
      Do while not MyRs.eof 
	     tc=MyRs("IdTipoCredito")
         xx=SetDiz(DizDatabase,"S_IdTipoCredito_"   & tc ,"")
         xx=SetDiz(DizDatabase,"S_DescTipoCredito_" & tc ,"")
         xx=SetDiz(DizDatabase,"N_ImptTipoCredito_" & tc , 0)
	  
	     MyRs.moveNext 
	  loop
      For Each K In DizDatabase
          nome = mid(k,3,99)
          xx=SetDiz(DizDatabase,k,MyRs(nome)) 
      next 
   end if 
	  MyRs.close 
   
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

%>