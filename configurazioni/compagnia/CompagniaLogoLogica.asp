<%
IdCompagnia=""
IdUpload=0
if FirstLoad then 
   IdCompagnia    = cdbl("0" & Session("swap_IdCompagnia"))
   OperAmmesse    = Session("swap_OperAmmesse")
   if Cdbl(IdCompagnia)=0 then 
      IdCompagnia = cdbl("0" & getValueOfDic(Pagedic,"IdCompagnia"))
   end if 
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
else
   IdCompagnia    = cdbl("0" & getValueOfDic(Pagedic,"IdCompagnia"))
   OperAmmesse    = getValueOfDic(Pagedic,"OperAmmesse")
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
end if 
if Cdbl(IdCompagnia)=0 then 
   response.redirect PaginaReturn
   response.end
end if 

FileCambiato=false
If FileCambiato=false and Oper="UPDATE1" and o.FileNameOf("FileP0")<>"" then 
   FileCambiato=true
end if 
If FileCambiato=false and Oper="UPDATE2" and o.FileNameOf("FileG0")<>"" then 
   FileCambiato=true
end if 

if Oper="UPDATE1" and FileCambiato then 
   NomeFilFull = o.FileNameOf("FileP0")
   sFileSplit = split(NomeFilFull, "\")
   sFile = sFileSplit(Ubound(sFileSplit))

   sFileWrite = "IMP" & IdCompagnia & "_" & Year(Now()) & Month(Now()) & Day(Now()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now()) &  "_" & sFile
  
   o.FileInputName = "FileP0"
   o.FileFullPath = PathBaseUpload  & sFileWrite
   o.save
   if o.Error <> ""  then
       MsgErrore= "Caricamento Fallito: " & o.Error & o.FileFullPath
   elseif err.number<>0 then
       MsgErrore= "Caricamento Errore : " & Err.Description
   else
       ConnMsde.execute "Update Compagnia Set LogoPiccolo='" & apici(sFileWrite) & "' where IdCompagnia=" & IdCompagnia
   end if 
end if 
if Oper="UPDATE2" and FileCambiato then 
   NomeFilFull = o.FileNameOf("FileG0")
   sFileSplit = split(NomeFilFull, "\")
   sFile = sFileSplit(Ubound(sFileSplit))

   sFileWrite = "IMP" & IdCompagnia & "_" & Year(Now()) & Month(Now()) & Day(Now()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now()) &  "_" & sFile
  
   o.FileInputName = "FileG0"
   o.FileFullPath = PathBaseUpload  & sFileWrite
   o.save
   if o.Error <> ""  then
       MsgErrore= "Caricamento Fallito: " & o.Error & o.FileFullPath
   elseif err.number<>0 then
       MsgErrore= "Caricamento Errore : " & Err.Description
   else
       ConnMsde.execute "Update Compagnia Set LogoGrande='" & apici(sFileWrite) & "' where IdCompagnia=" & IdCompagnia
   end if 
end if 
Dim Rs 
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.CursorLocation = 3 
Rs.Open "Select * from Compagnia Where IdCompagnia=" & IdCompagnia, ConnMsde

DescCompagnia=Rs("DescCompagnia")
LogoPiccolo = Rs("LogoPiccolo")
LogoGrande  = Rs("LogoGrande")
Rs.close 

  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdCompagnia"        ,IdCompagnia)
  xx=setValueOfDic(Pagedic,"IdCompagniaDesc"    ,IdCompagniaDesc)
  xx=setValueOfDic(Pagedic,"OperAmmesse"        ,OperAmmesse)
  xx=setValueOfDic(Pagedic,"PaginaReturn"       ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

%>