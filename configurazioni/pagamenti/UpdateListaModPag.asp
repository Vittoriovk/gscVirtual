<%

'IdAccountModPag        = Account 
err.clear

MySqlDoc = ""
MySqlDoc = MySqlDoc & "select * from AccountModPag Where IdAccount = " & IdAccountModPag
  
Set RsDoc = Server.CreateObject("ADODB.Recordset")

RsDoc.CursorLocation = 3 
RsDoc.Open MySqlDoc, ConnMsde
if err.number=0 then
   if RsDoc.EOF then 
      FlagBorsellino = 0
      ImptBorsellino = 0
      ImptBorsellinoImpe = 0
      ImptBorsellinoUtil = 0
      ImptBorsellinoDisp = 0
      ImptBorsellinoValo = 0
      FlagFido = 0
      ImptFido = 0
      ImptFidoImpe = 0
      ImptFidoUtil = 0
      ImptFidoDisp = 0
      ImptFidoValo = 0
      FlagEstratto = 0
      ImptEstratto = 0
      ImptEstrattoImpe = 0
      ImptEstrattoUtil = 0
      ImptEstrattoDisp = 0
      ImptEstrattoValo = 0
   else
      FlagBorsellino = RsDoc("FlagBorsellino")
      ImptBorsellino = RsDoc("ImptBorsellino")
      ImptBorsellinoImpe = RsDoc("ImptBorsellinoImpe")
	  ImptBorsellinoUtil = RsDoc("ImptBorsellinoUtil")
      ImptBorsellinoDisp = RsDoc("ImptBorsellinoDisp")
      ImptBorsellinoValo = RsDoc("ImptBorsellinoValo")
      FlagFido = RsDoc("FlagFido")
      ImptFido = RsDoc("ImptFido")
      ImptFidoImpe = RsDoc("ImptFidoImpe")
      ImptFidoUtil = RsDoc("ImptFidoUtil")
      ImptFidoDisp = RsDoc("ImptFidoDisp")
      ImptFidoValo = RsDoc("ImptFidoValo")
      FlagEstratto = RsDoc("FlagEstratto")
      ImptEstratto = RsDoc("ImptEstratto")
      ImptEstrattoImpe = RsDoc("ImptEstrattoImpe")
      ImptEstrattoUtil = RsDoc("ImptEstrattoUtil")
      ImptEstrattoDisp = RsDoc("ImptEstrattoDisp")
      ImptEstrattoValo = RsDoc("ImptEstrattoValo")
   end if 
   RsDoc.close 
end if 

if Request("LMP_checkBorsellino0")="S" then 
   FlagBorsellino=1
else
   FlagBorsellino=0
end if 
if Request("LMP_checkFido0")="S" then 
   FlagFido=1
else
   FlagFido=0
end if 
if Request("LMP_checkEstratto0")="S" then 
   FlagEstratto=1
else
   FlagEstratto=0
end if 

ImptFidoNew     = TestNumeroPos(Request("LMP_ImptFido0"))
ImptEstrattoNew = TestNumeroPos(Request("LMP_ImptEstratto0"))
'response.write ImptFidoNew & " " & ImptEstrattoNew 
ImptFidoUsed = cdbl(ImptFidoImpe)     + cdbl(ImptFidoUtil)
ImptEstrUsed = cdbl(ImptEstrattoImpe) + cdbl(ImptEstrattoUtil)
If cdbl(ImptFidoNew) >= cdbl(ImptFidoUsed) and cdbl(ImptEstrattoNew)>=cdbl(ImptEstrUsed) then  
   ImptFido         = cdbl(ImptFidoNew)
   ImptFidoDisp     = Cdbl(ImptFido)        - cdbl(ImptFidoUsed)
   ImptEstratto     = cdbl(ImptEstrattoNew)
   ImptEstrattoDisp = Cdbl(ImptEstratto)    - cdbl(ImptEstrUsed)
   myUpd = ""
   myUpd = myUpd & " update AccountModPag set "
   myUpd = myUpd & " FlagBorsellino = "   & FlagBorsellino
   myUpd = myUpd & ",FlagFido = "         & FlagFido
   myUpd = myUpd & ",FlagEstratto = "     & FlagEstratto
   myUpd = myUpd & ",ImptFido = "         & NumForDb(ImptFido)
   myUpd = myUpd & ",ImptFidoDisp = "     & NumForDb(ImptFidoDisp)
   myUpd = myUpd & ",ImptEstratto = "     & NumForDb(ImptEstratto)   
   myUpd = myUpd & ",ImptEstrattoDisp = " & NumForDb(ImptEstrattoDisp)      
   myUpd = myUpd & " Where IdAccount = "  & IdAccountModPag 
   
   if Request("LMP_UPDATE")="" then 
      ConnMsde.execute myUpd 
      'response.write myUpd
   end if 
   
end if 

           
%>
