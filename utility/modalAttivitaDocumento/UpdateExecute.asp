<!--#include virtual="/gscVirtual/common/clsupload.asp"-->
<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<!--#include virtual="/gscVirtual/common/function.asp"-->
<!--#include virtual="/gscVirtual/common/functionNew.asp"-->
<!--#include virtual="/gscVirtual/common/connDb.asp"-->
<%
   on error resume next 
   set o = new clsUpload
   FlagDebug = false
   MsgErrore = ""
   
   PathFile            = o.valueOf("MAD_PathFile") 
   IdAttivitaDocumento = o.valueOf("MAD_IdAttivitaDocumento")
   IdAttivitaDocumento = Cdbl("0" & IdAttivitaDocumento)
   if Cdbl(IdAttivitaDocumento)=0 then 
      response.write "ERR: id non valorizzato"
      response.end 
   end if
   IdUpload = o.valueOf("MAD_IdUpload")
   IdUpload = Cdbl("0" & IdUpload)

   if FlagDebug = true then
      response.write "leggo Attivitadocumento:" & err.description & "<br>"
   end if 
   IdTabella          = LeggiCampo("select * from AttivitaDocumento Where IdAttivitaDocumento=" & IdAttivitaDocumento,"IdAttivita")
   IdTabellaKeyInt    = IdAttivitaDocumento
   IdTabellaKeyString = ""
   
   
   NomeFilFull = o.FileNameOf("MAD_FileToUpload0")
   'file obbligatorio per caricare 
   if trim(NomeFilFull)="" and cdbl(idUpload)=0 then 
      response.write "ERR: file non valorizzato"
      response.end 
   end if 
   ValidoDal = "0" & DataStringa(o.ValueOf("MAD_ValidoDal0"))
   if isnumeric(ValidoDal)=false then 
      ValidoDal=0
   else
      ValidoDal=Cdbl(ValidoDal)
   end if 
   ValidoAl  = "0" & DataStringa(o.ValueOf("MAD_ValidoAl0"))
   if isnumeric(ValidoAl)=false then 
      ValidoAl=20991231
   else
      ValidoAl=Cdbl(ValidoAl)
   end if 
   if ValidoAl=0 then 
      ValidoAl=20991231
   end if 
   if FlagDebug = true then
      response.write "ValidoDal:" & ValidoDal & " ValidoAl:" & ValidoAl & " err=" & err.description & "<br>"
   end if 
   
   DescDocumento   = o.ValueOf("MAD_DescDocumento0")
   IdTipoDocumento = o.ValueOf("MAD_IdTipoDocumento")
   'inizio aggiornamento upload
   if FlagDebug = true then
      response.write "IdTipoDocumento:" & IdTipoDocumento & " DescDocumento:" & ValidoAl & " err=" & err.description & "<br>"
   end if 
   NomeFilFull = o.FileNameOf("MAD_FileToUpload0")
   sFileSplit = split(NomeFilFull, "\")
   if FlagDebug = true then
      response.write "NomeFilFull:" & NomeFilFull & " ubound:" & Ubound(sFileSplit) & " err=" & err.description & "<br>"
   end if 
   
   sFile = sFileSplit(Ubound(sFileSplit))
   sFileWrite = o.valueOf("MAD_PathFile")
   if FlagDebug = true then
      response.write "sFile:" & sFile & " sFileWrite:" & sFileWrite & " err=" & err.description & "<br>"
   end if 
  
   q = ""
   flagUpdateAtt=false 
   if cdbl(IdUpload)=0 then 
      
      q = q & " insert into Upload (IdTabella,IdTabellaKeyInt,IdTabellaKeyString,DataUpload"
      q = q & " ,TimeUpload,IdTipoDocumento"
	  q = q & " ,DescBreve,DescEstesa,NomeDocumento,PathDocumento"
	  q = q & " ,ValidoDal,ValidoAl) "
      q = q & " values ("
      q = q & " '" & Apici(IdTabella) & "'"
      q = q & ", " & IdTabellaKeyInt
      q = q & ",'" & Apici(IdTabellaKeyString) & "'"
	  
	  q = q & ", " &  Dtos()
      q = q & ", " & TimeToS() & "," & NumForDb(IdTipoDocumento)
	  q = q & ",'" & apici(DescDocumento) & "'"
	  q = q & ",'" & apici(DescDocumento) & "'"
	  q = q & ",'" & apici(sFile) & "'"
	  q = q & ",'" & apici(sFileWrite) & "'"
      q = q & ", " & ValidoDal
	  q = q & ", " & ValidoAl & ")"   
      if FlagDebug = true then
         response.write "q:" & q & " err=" & err.description & "<br>"
      end if 	  
      xx=writeTrace(q)  
	  ConnMsde.execute q 
	  if Err.Number<>0 then 
	     MsgErrore="ERR:" & Err.description
	  else 
	     IdUpload=GetTableIdentity("Upload")
		 flagUpdateAtt=true 
	  end if 
   else 
      q = q & "update Upload set "
      q = q & " DataUpload = " & Dtos()
      q = q & ",TimeUpload = " & TimeToS()
      q = q & ",ValidoDal=" & ValidoDal
      q = q & ",ValidoAl="  & ValidoAl
      q = q & ",IdTipoDocumento = " & NumForDb(IdTipoDocumento)
      q = q & ",DescBreve='"        & apici(DescDocumento) & "'"
      q = q & ",DescEstesa='"       & apici(DescDocumento) & "'"	  
      if PathFile<>"" then
         q = q & ",NomeDocumento='" & apici(sFile) & "'"
         q = q & ",PathDocumento='" & apici(sFileWrite) & "'"
      end if   
	  q = q & "where IdUpload=" & IdUpload 
	  xx=writeTrace(q)  
	  ConnMsde.execute q 
	  if Err.Number<>0 then 
	     MsgErrore="ERR:" & Err.description
	  end if 
   end if 
   if MsgErrore<>"" then
      response.write MsgErrore
	  response.end
   end if 

   'aggiornamento tabella
   if flagUpdateAtt=true and err.number=0 then 
      ConnMsde.execute "Update AttivitaDocumento Set IdUpload=" & IdUpload & " Where IdAttivitaDocumento=" & IdAttivitaDocumento
   end if    
   response.write Err.description 
%>
