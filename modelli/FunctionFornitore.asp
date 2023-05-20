<%
Function UpdateFornitore()
Dim v_id , MySql , Esito , v_ret
on error resume next 
   Esito = ""
   v_id = cdbl("0" & GetDiz(DizDatabase,"IdFornitore"))
   if cdbl(v_id)>0 then 
      MySql = ""
      MySql = MySql & " update Fornitore set "
      MySql = MySql & " DescFornitore='"   & apici(GetDiz(DizDatabase,"DescFornitore")) & "'"
      MySql = MySql & ",DescCognome ='"    & apici(GetDiz(DizDatabase,"DescCognome"))   & "'"
      MySql = MySql & ",DescNome ='"       & apici(GetDiz(DizDatabase,"DescNome"))      & "'"
      MySql = MySql & ",IdTipoDitta = '"   & apici(GetDiz(DizDatabase,"IdTipoDitta"))   & "'"
      MySql = MySql & ",IdTipoSocieta = '" & apici(GetDiz(DizDatabase,"IdTipoSocieta")) & "'"
	  MySql = MySql & ",IdTipoMandato = '" & apici(GetDiz(DizDatabase,"IdTipoMandato")) & "'"
	  MySql = MySql & ",IdTipoIncasso = '" & apici(GetDiz(DizDatabase,"IdTipoIncasso")) & "'"
      MySql = MySql & ",PartitaIva   ='"   & apici(GetDiz(DizDatabase,"PartitaIva"))    & "'"
      MySql = MySql & ",CodiceFiscale='"   & apici(GetDiz(DizDatabase,"CodiceFiscale")) & "'"
	  MySql = MySql & ",IdSezioneRui='"    & apici(GetDiz(DizDatabase,"IdSezioneRui"))  & "'"
 	  MySql = MySql & ",NumeroRui='"       & apici(GetDiz(DizDatabase,"NumeroRui"))     & "'"
	  MySql = MySql & ",DescRuolo='"       & apici(GetDiz(DizDatabase,"DescRuolo"))     & "'"
	  MySql = MySql & ",DataIscrizioneRui=" &     GetDiz(DizDatabase,"DataIscrizioneRui")
      MySql = MySql & " where IdFornitore = " & v_id
	  'response.write MySql
      ConnMsde.execute MySql 
      if err.number<>0 then 
         Esito = Err.description 
	  else
		MySql = "" 
		MySql = MySql & " update Account "
		MySql = MySql & " set Nominativo = '"  & apici(GetDiz(DizDatabase,"DescFornitore")) & "'"
		MySql = MySql & " where IdAccount = " & GetDiz(DizDatabase,"IdAccount")
		ConnMsde.execute MySql
	  
      end if 		 
   end if 
 
   UpdateFornitore=Esito 
End function 
%>