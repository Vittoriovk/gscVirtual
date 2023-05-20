<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<!--#include virtual="/gscVirtual/common/FunCallOtherPage.asp"-->
<%
FlagDebug=true 
IdAttivita=Request("IdAttivita")
IdNumAttivita=cdbl("0" & Request("IdNumAttivita"))

if IdAttivita<>"" and cdbl(IdNumAttivita)>0 then 

   if IdAttivita="CAUZ_DEFI" then 
      opUrl     = getDomain() & "/gscvirtual/pdf/PdfCauzioneDefinitivaBozza.asp"
      opMethod  = "POST"
      opData    = "IdCauzione=" & IdNumAttivita
	  opData    = opData & "&SendBrowser=N"
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
      if mid(opResp,1,3) = "OK:" then 
	     filName = mid(OpResp,4)
         qUpd = ""
         qUpd = qUpd & " update CauzioneDef "
         qUpd = qUpd & " Set PathDocumentoRichiesta='" & filName & "'"
         qUpd = qUpd & " where IdCauzioneDef = " & IdNumAttivita
         if FlagDebug=true then 
            response.write qUpd 
         end if		 
         ConnMsde.execute qUpd
      end if 
   end if 
end if 

response.end 

%>