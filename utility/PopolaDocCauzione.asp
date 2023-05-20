<!--#include virtual="/gscVirtual/include/includeStd.asp"-->


<%

IdTipoCauzione = Request("IdTipoCauzione")
IdCauzione     = Cdbl("0" & Request("IdCauzione"))
%>

   <div class="row">
       <div class="col-8">
            <div class="form-group "><b>Documento</b></div>        
       </div>
       <div class="col-2">
            <div class="form-group "><b>Leggi</b></div>        
       </div>
   </div>
   
      <%
	  Set RsCauzione = Server.CreateObject("ADODB.Recordset")
	  MyContQ = ""
	  MyContQ = MyContQ & " select * from Upload "
	  MyContQ = MyContQ & " where IdTabella = '"     & apici(IdTipoCauzione) & "'"
      MyContQ = MyContQ & " and   IdTabellaKeyInt =" & IdCauzione

      RsCauzione.CursorLocation = 3
      RsCauzione.Open MyContQ, ConnMsde 
      LeggiCauzione=true 
	  Conta=0
      If Err.number<>0 then	
       	 LeggiCauzione=false
      elseIf RsCauzione.EOF then	
         LeggiCauzione=false
		 RsCauzione.close 
      End if
	  if LeggiCauzione then 
	     
		 flagC=true 
	     do while not RsCauzione.eof 
		    conta=conta+1
			color=" "
			if flagC then 
			   color="bg-light"
			   flagC = false 
			else 
			   flagC=true 
			end if 
		 %>
   <div class="row <%=color%>">
       <div class="col-8">
            <div class="form-group "><%=RsCauzione("DescBreve")%></div>        
       </div>
       <div class="col-2">
            <div class="form-group ">
			<%Linkdocumento=RsCauzione("PathDocumento")%>
			<!--#include virtual="/gscVirtual/common/linkForDownload.asp"-->
			</div>        
       </div>
   </div>

		 <%
		    RsCauzione.moveNext 
		 loop  
		 RsCauzione.close
	  end if
  
	  if conta=0 then 
	     response.write "<h2>Nessun documento in archivio</h2> "
	  end if 
      
	  %>
