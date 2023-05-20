<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/common/FunCallOtherPage.asp"-->
<%
FlagDebug=true 
IdAttivita=Request("IdAttivita")
IdNumAttivita=cdbl("0" & Request("IdNumAttivita"))

if IdAttivita<>"" and cdbl(IdNumAttivita)>0 then 
   Set Rs = Server.CreateObject("ADODB.Recordset")
   DescRequest = IdAttivita & "_" & IdNumAttivita

   MySql = ""
   MySql = MySql & " select DescEstesa,PathDocumento "
   MySql = MySql & " from AttivitaDocumento A, Upload B "
   MySql = MySql & " where A.IdAttivita = '" & IdAttivita & "' "
   MySql = MySql & " and A.IdNumAttivita = " & IdNumAttivita
   MySql = MySql & " and A.IdUpload = B.IdUpload "
   MySql = MySql & " And B.PathDocumento<>''"
   
   Rs.CursorLocation = 3
   Rs.Open MySql, ConnMsde

   'ci sono files da scaricare 
   if Rs.Eof = false then    
      filName = DescRequest & "_" & Year(Now()) & Month(Now()) & Day(Now()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now()) 
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
	     if IdAttivita="CAUZ_DEFI" then 
	        qUpd = ""
		    qUpd = qUpd & " update CauzioneDef "
		    qUpd = qUpd & " Set PathDocumentoZip='tempZip/" & filName & ".zip'"
		    qUpd = qUpd & " where IdCauzioneDef = " & IdNumAttivita
	        if FlagDebug=true then 
	           response.write qUpd 
	        end if		 
	        ConnMsde.execute qUpd
		 end if 
	  end if 
	  
   End if 
   Set Rs = nothing 
end if 

response.end 

%>