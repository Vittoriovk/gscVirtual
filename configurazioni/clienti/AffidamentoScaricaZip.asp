<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/common/FunCallOtherPage.asp"-->
<%
FlagDebug=false 
IdAffidamentoRichiestaComp=cdbl("0" & Request("IdAffidamentoRichiestaComp"))

if cdbl(IdAffidamentoRichiestaComp)>0 then 
   Set Rs = Server.CreateObject("ADODB.Recordset")
   MySql = ""
   MySql = MySql & " select A.*,B.*,D.Denominazione as DescCliente   "
   MySql = MySql & " from AffidamentoRichiestaComp A, Fornitore B, AffidamentoRichiesta C , Cliente D"
   MySql = MySql & " Where A.IdFornitore = b.Idfornitore "
   MySql = MySql & " And A.IdAffidamentoRichiesta = C.IdAffidamentoRichiesta "
   MySql = MySql & " And D.IdAccount = C.IdAccountCliente "
   MySql = MySql & " And A.IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
   if FlagDebug=true then 
      response.write MySql & "<br>"
   end if
   DescCliente = LeggiCampo(MySql,"DescCliente")

   MySql = ""
   MySql = MySql & " select DescEstesa,PathDocumento "
   MySql = MySql & " from AffidamentoRichiestaCompDoc A, AccountDocumento B, Upload C "
   MySql = MySql & " where A.IdAccountDocumento = B.IdAccountDocumento "
   MySql = MySql & " and B.IdUpload = C.IdUpload "
   MySql = MySql & " and Pathdocumento<>'' "
   MySql = MySql & " And A.IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
   
   if FlagDebug=true then 
      response.write MySql & "<br>"
   end if
   
   Rs.CursorLocation = 3
   Rs.Open MySql, ConnMsde

   'ci sono files da scaricare 
   if Rs.Eof = false then    
      filName = DescCliente & "_" & Year(Now()) & Month(Now()) & Day(Now()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now()) 
      DirPath = replace(VirtualPath & DirectoryUpload,"//","/")
 
      zipPath = replace(DirPath & "/tempZip/" & filName ,"//","/")
   
      iniPath = Server.MapPath(dirPath)
      fulPath = Server.MapPath(zipPath)
      
      Set fs = CreateObject("Scripting.FileSystemObject")
      if FlagDebug=true then 
         response.write fulPath & "<br>"
      end if
      
      
      fs.CreateFolder fulPath 
   
      if Right(iniPath,1)<>"\" and Right(iniPath,1)<>"/" then 
         iniPath=iniPath & "/"
      end if 
      if Right(fulPath,1)<>"\" and Right(fulPath,1)<>"/" then 
         fulPath=fulPath & "/"
      end if 
      if FlagDebug=true then
         response.write "inP=" & iniPath  & "<br>"
         response.write "deP=" & fulPath  & "<br>"
      end if 
      conta = 0 
      do while not Rs.eof
         DescEstesa    = Rs("DescEstesa")
         PathDocumento = Rs("PathDocumento")
         ptr = InStrRev(PathDocumento,".")
         Estensione = ""
          if ptr>0 then 
            Estensione = mid(PathDocumento,ptr) 
          end if 

          conta=conta+1
          PathFisicoIn  = iniPath & PathDocumento
          PathfisicoOut = fulPath & DescEstesa & Estensione 
          if FlagDebug=true then
             response.write "in =" & conta & " " & PathfisicoIn  & "<br>"
             response.write "out=" & conta & " " & PathfisicoOut & "<br>"
          end if 
          fs.CopyFile PathFisicoIn, PathfisicoOut,true 
          Rs.Movenext 
      loop
      Rs.close 
   
      'aggiungo eventuali coobligati
      err.clear
      MySql = ""
      MySql = MySql & " SELECT a.RagSoc,a.IdAccountCoobbligato,b.IdAccount"
      MySql = MySql & ",d.DescDocumento,e.PathDocumento"
      MySql = MySql & " FROM AffidamentoRichiestaCompCoob A, AccountCoobbligato b "
      MySql = MySql & ",AccountDocumento C , Documento D , Upload E"
      MySql = MySql & " where A.IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
      MySql = MySql & " and A.IdAccountCoobbligato = b.IdAccountCoobbligato"
      MySql = MySql & " and b.IdAccount = c.IdAccount"
      MySql = MySql & " and c.tipoRife = 'COOB'"
      MySql = MySql & " and A.IdAccountCoobbligato=c.Idrife"
      MySql = MySql & " and C.IdDocumento = d.IdDocumento"
      MySql = MySql & " and c.IdUpload = E.IdUpload "
      MySql = MySql & " and E.Pathdocumento<>''"
      Rs.Open MySql, ConnMsde
      if Err.number = 0 then 
         do while not Rs.eof
            DescEstesa    = Rs("RagSoc") & "_" & Rs("DescDocumento")
            PathDocumento = Rs("PathDocumento")
            ptr = InStrRev(PathDocumento,".")
            Estensione = ""
            if ptr>0 then 
               Estensione = mid(PathDocumento,ptr) 
            end if 
            conta=conta+1
            PathFisicoIn  = iniPath & PathDocumento
            PathfisicoOut = fulPath & DescEstesa & Estensione 
            if FlagDebug=true then
               response.write "in =" & conta & " " & PathfisicoIn  & "<br>"
               response.write "out=" & conta & " " & PathfisicoOut & "<br>"
            end if 
            fs.CopyFile PathFisicoIn, PathfisicoOut,true 
            Rs.Movenext 
         loop
         Rs.close 
  
      end if 
   
      ' chiamo zip
      zipCreate = Mid(fulPath,1,Len(fulPath)-1) & ".zip"
      opUrl     = getDomain() & "/gscvirtual/api/zipDirectory.aspx"
      opMethod  = "POST"
      opData    = "sourceDir=" & Server.URLEncode(fulPath)
      opData    = opData & "&sourceZip=" & Server.URLEncode(zipCreate)
      if FlagDebug=true then
         response.write opUrl & "?" & OpData & "<br>"
      end if 
      opType    = ""
      opReferer = "http://www.mysite.com"
      opResp    = ""
      xx = CallOtherPage(opUrl,opMethod,opData,opType,opReferer,opResp)      
      if FlagDebug=true then 
         response.write "esito  = " & xx & "<br>"
         response.write "opResp = " & OpResp 
      end if
      if opResp = "OK" then 
         qUpd = ""
         qUpd = qUpd & " update AffidamentoRichiestaComp "
         qUpd = qUpd & " Set PathDocumentoZip='tempZip/" & filName & ".zip'"
         qUpd = qUpd & " where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
         if FlagDebug=true then 
            response.write qUpd 
         end if         
         ConnMsde.execute qUpd
      end if 
      
   End if 
   Set Rs = nothing 
end if 

response.end 

%>