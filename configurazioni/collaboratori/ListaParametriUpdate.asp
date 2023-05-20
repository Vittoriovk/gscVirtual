
<%

campo_VAL_COB = trim(Request("campo_VAL_COB"))
campo_VAL_ATI = trim(Request("campo_VAL_ATI"))
campo_ASS_PRO = trim(Request("campo_ASS_PRO"))
IdAccountParametro = cdbl("0" & IdAccountParametro)
if Cdbl(IdAccountParametro)>0 then 
   if campo_VAL_COB<>"" then
      tmpVal = Request(campo_VAL_COB)
      xx=AggiornaAccountParametro(IdAccountParametro,"VAL_COB",tmpVal)
   end if 
   if campo_VAL_ATI<>"" then 
      tmpVal = Request(campo_VAL_ATI)
      xx=AggiornaAccountParametro(IdAccountParametro,"VAL_ATI",tmpVal)
   end if  
   if campo_ASS_PRO<>"" then 
      tmpVal = Request(campo_ASS_PRO)
      xx=AggiornaAccountParametro(IdAccountParametro,"ASS_PRO",tmpVal)
   end if     
end if 

function AggiornaAccountParametro(IdAccount,IdTipoParametro,Valore)
Dim myQ,where,tmp  

   MyQ = ""
   where = ""
   where = where & " where IdAccount = " & IdAccount 
   where = where & " and IdTipoParametro = '" & IdTipoParametro & "'" 

	  
   if Valore="" then 
      MyQ = MyQ & " delete from AccountTipoParametro "
	  MyQ = MyQ & where
	  ConnMsde.execute MyQ
   else 
      tmp=LeggiCampo("select * from AccountTipoParametro " & where,"IdTipoParametro")
	  if tmp="" then 
         MyQ = MyQ & " insert into AccountTipoParametro(IdAccount,IdTipoParametro,ValoreParametro) "
		 MyQ = MyQ & " values(" & IdAccount & ",'" & IdTipoParametro & "','" & Valore & "')"	  
	  else
         MyQ = MyQ & " update AccountTipoParametro "
		 MyQ = MyQ & " set ValoreParametro='" & Valore & "'"
	     MyQ = MyQ & where     
	  end if 
	  ConnMsde.execute MyQ
   end if 
end function 

%>
 


