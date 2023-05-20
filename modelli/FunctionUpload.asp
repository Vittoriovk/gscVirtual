<%
Function InizializeUpload(DizDatabase,IdUpload)
Dim xx
on error resume next 
   if isObject(DizDatabase)=false then 
      Set DizDatabase = CreateObject("Scripting.Dictionary")
   else 
      if ucase(typename(DizDatabase))<> ucase("Dictionary") then 
         Set DizDatabase = CreateObject("Scripting.Dictionary")
      end if 
   end if 

   xx=SetDiz(DizDatabase,"IdUpload"          , 0)
   xx=SetDiz(DizDatabase,"IdTabella"         ,"")
   xx=SetDiz(DizDatabase,"IdTabellaKeyInt"   , 0)
   xx=SetDiz(DizDatabase,"IdTabellaKeyString","")
   xx=SetDiz(DizDatabase,"DataUpload"        ,Dtos())
   xx=SetDiz(DizDatabase,"TimeUpload"        ,TimeToS())
   xx=SetDiz(DizDatabase,"IdTipoDocumento"   , 0)
   xx=SetDiz(DizDatabase,"DescBreve"         ,"")
   xx=SetDiz(DizDatabase,"DescEstesa"        ,"")
   xx=SetDiz(DizDatabase,"NomeDocumento"     ,"")
   xx=SetDiz(DizDatabase,"PathDocumento"     ,"")
   xx=SetDiz(DizDatabase,"ValidoDal"         , 0)
   xx=SetDiz(DizDatabase,"ValidoAl"          , 0)

   if Cdbl(IdUpload)>0 then 
      xx=GetInfoRecordset(Dizionario,"select + from Upload Where IdUpload=" & IdUpload)
   end if 
   
End function 

Function UpdateUpload(DizDatabase)
Dim v_id , MySql , Esito , v_ret
on error resume next 
   Esito = ""
   v_id = cdbl("0" & GetDiz(DizDatabase,"IdUpload"))
   response.write "ecco1:" & v_id
   if cdbl(v_id)>0 then 
      MySql = ""
      MySql = MySql & " update Upload set "
      MySql = MySql & " IdTabella='"           & apici(GetDiz(DizDatabase,"IdTabella"))               & "'"
      MySql = MySql & ",IdTabellaKeyInt ="     & TestNumeroPos(GetDiz(DizDatabase,"IdTabellaKeyInt"))
      MySql = MySql & ",IdTabellaKeyString ='" & apici(GetDiz(DizDatabase,"IdTabellaKeyString"))      & "'"
      MySql = MySql & ",DataUpload ="          & TestNumeroPos(GetDiz(DizDatabase,"DataUpload"))
      MySql = MySql & ",TimeUpload ="          & TestNumeroPos(GetDiz(DizDatabase,"TimeUpload"))
	  MySql = MySql & ",IdTipoDocumento ="     & TestNumeroPos(GetDiz(DizDatabase,"IdTipoDocumento"))
	  MySql = MySql & ",DescBreve = '"         & apici(GetDiz(DizDatabase,"DescBreve"))               & "'"
	  MySql = MySql & ",DescEstesa = '"        & apici(GetDiz(DizDatabase,"DescEstesa"))              & "'"
      MySql = MySql & ",NomeDocumento   ='"    & apici(GetDiz(DizDatabase,"NomeDocumento"))           & "'"
      MySql = MySql & ",PathDocumento='"       & apici(GetDiz(DizDatabase,"PathDocumento"))           & "'"
	  MySql = MySql & ",ValidoDal ="           & TestNumeroPos(GetDiz(DizDatabase,"ValidoDal"))
	  MySql = MySql & ",ValidoAl ="            & TestNumeroPos(GetDiz(DizDatabase,"ValidoAl"))
      MySql = MySql & " where IdUpload = " & v_id
	  response.write MySql & Err.description
      ConnMsde.execute MySql 
      if err.number<>0 then 
         Esito = Err.description 
      else
         Esito = v_id
      end if 
   else 
      MySql  = "" 
      MySql  = MySql  & " insert into Upload ("
	  MySql  = MySql  & " IdTabella"
	  MySql  = MySql  & ",IdTabellaKeyInt"
	  MySql  = MySql  & ",IdTabellaKeyString"
      MySql  = MySql  & ",DataUpload"
	  MySql  = MySql  & ",TimeUpload"
	  MySql  = MySql  & ",IdTipoDocumento"
	  MySql  = MySql  & ",DescBreve"
	  MySql  = MySql  & ",DescEstesa"
	  MySql  = MySql  & ",NomeDocumento"
	  MySql  = MySql  & ",PathDocumento"
	  MySql  = MySql  & ",ValidoDal"
	  MySql  = MySql  & ",ValidoAl "
      MySql  = MySql  & " ) values ("
      MySql  = MySql  & " '" & Apici(GetDiz(DizDatabase,"IdTabella"))               & "'"
      MySql  = MySql  & ", " & TestNumeroPos(GetDiz(DizDatabase,"IdTabellaKeyInt"))
      MySql  = MySql  & ",'" & Apici(GetDiz(DizDatabase,"IdTabellaKeyString"))      & "'"
	  MySql  = MySql  & ", " & TestNumeroPos(GetDiz(DizDatabase,"DataUpload"))
	  MySql  = MySql  & ", " & TestNumeroPos(GetDiz(DizDatabase,"TimeUpload"))
	  MySql  = MySql  & ", " & TestNumeroPos(GetDiz(DizDatabase,"IdTipoDocumento"))
	  MySql  = MySql  & ",'" & Apici(GetDiz(DizDatabase,"DescBreve"))               & "'"
	  MySql  = MySql  & ",'" & Apici(GetDiz(DizDatabase,"DescEstesa"))              & "'"
	  MySql  = MySql  & ",'" & Apici(GetDiz(DizDatabase,"NomeDocumento"))           & "'"
	  MySql  = MySql  & ",'" & Apici(GetDiz(DizDatabase,"PathDocumento"))           & "'"
	  MySql  = MySql  & ", " & TestNumeroPos(GetDiz(DizDatabase,"ValidoDal"))
	  MySql  = MySql  & ", " & TestNumeroPos(GetDiz(DizDatabase,"ValidoAl"))	  
      MySql  = MySql  & ")" 
      ConnMsde.execute MySql 
      if err.number<>0 then 
         Esito = Err.description 
      else
         Esito = GetTableIdentity("Upload")
      end if 
   end if 
 response.write "uu:" & MySql
   UpdateUpload=Esito 
End function 
%>