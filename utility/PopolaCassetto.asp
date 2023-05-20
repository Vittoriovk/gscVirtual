<!--#include virtual="/gscVirtual/include/includeStd.asp"-->


<%

IdAccount  =Request("IdAccount")
IdDocumento=Request("IdDocumento")
TipoRife   =Request("TipoRife")
IdRife     =Cdbl("0" & Request("idRife"))
%>

        <div class="table-responsive"><table class="table"><tbody>
        <thead>
        <tr>
            <th scope="col">Selezione</th>
            <th scope="col">documento</th>
			<th scope="col">Stato</th>
            <th scope="col">Valido Dal</th>
			<th scope="col">Valido Al</th>
        </tr>
        </thead>

      <%
	  Oggi=Dtos()
	  Set RsCassetto = Server.CreateObject("ADODB.Recordset")
	  MyContQ = ""
	  MyContQ = MyContQ & " select * from AccountDocumento A, Upload B,Documento C "
	  MyContQ = MyContQ & " where A.IdAccount = " & IdAccount
      MyContQ = MyContQ & " and   A.IdDocumento =" & IdDocumento
      MyContQ = MyContQ & " and   A.TipoRife='" & Apici(TipoRife) & "'"
	  MyContQ = MyContQ & " and   A.IdRife=" & IdRife
	  MyContQ = MyContQ & " and   C.IdDocumento =" & IdDocumento
	  MyContQ = MyContQ & " and   A.IdUpload = b.IdUpload"
	  'MyContQ = MyContQ & " and   B.ValidoDal <=" & Oggi
	  MyContQ = MyContQ & " and   B.ValidoAl >=" & Oggi
      MyContQ = MyContQ & " order By ValidoDal "  

      RsCassetto.CursorLocation = 3
      RsCassetto.Open MyContQ, ConnMsde 
      LeggiContatti=true 
	  Conta=0
      If Err.number<>0 then	
       	 LeggiContatti=false
      elseIf RsCassetto.EOF then	
         LeggiContatti=false
		 RsCassetto.close 
      End if
	  if LeggiContatti then 
	     
	     do while not RsCassetto.eof 
		    conta=conta+1
			checked=""
			if conta=1 then 
			   checked=" checked "
			end if 
		 %>
        <tr>
            <td scope="col">
		        <div class="form-check">
			      <input name="CassettoCampoNome" type="radio" id="radio<%=conta%>" 
				  value="<%=RsCassetto("IdAccountDocumento")%>" <%=checked%>>
		  </div>
  		    </td>
			<%
			DescDocumento=RsCassetto("DescBreve") 
			if DescDocumento="" then 
			   DescDocumento=RsCassetto("DescDocumento")
			end if 			
			%>
            <td scope="col"><%=DescDocumento%></td>
			<%
			IdTipoValidazione=apici(RsCassetto("IdTipoValidazione"))
			DescTipoValidazione = LeggiCampo("Select * from TipoValidazione Where IdTipoValidazione='" & IdTipoValidazione & "'" ,"DescTipoValidazione")
			%>
            <td scope="col"><%=DescTipoValidazione%></td>			
            <td scope="col"><%=Stod(RsCassetto("ValidoDal"))%></td>
			<td scope="col"><%=Stod(RsCassetto("ValidoAl"))%></td>
        </tr>
		
		 
		 <%
		    RsCassetto.moveNext 
		 loop  
		 RsCassetto.close
	  end if
%>
      </tbody></table></div> <!-- table responsive fluid -->
	  <input type="hidden" name="cassettopieno" id="cassettopieno" value="<%=conta%>">
<%	  
	  if conta=0 then 
	     response.write "<h2>Nessun documento in archivio</h2> "
	  end if 
      
	  %>
