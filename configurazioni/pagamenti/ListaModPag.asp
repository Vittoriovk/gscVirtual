<%
'input
'OpDocAmm               = Operazioni ammesse 
'                         I = mette modifica
'                         L = solo lista attivi

'IdAccountModPag        = Account 

err.clear
LMP_SoloLettura=""
if OpDocAmm<>"I" and OpDocAmm<>"U" then 
   LMP_SoloLettura = " readonly "
end if 

MySqlDoc = ""
MySqlDoc = MySqlDoc & "select * from AccountModPag Where IdAccount = " & IdAccountModPag
  
Set RsDoc = Server.CreateObject("ADODB.Recordset")

FlagAction=""
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
	  ImptEstrattoUtil = 0
      ImptEstrattoImpe = 0
      ImptEstrattoDisp = 0
      ImptEstrattoValo = 0
	  qIns = ""
	  qIns = qIns & "Insert into AccountModPag (IdAccount)"
	  qIns = qIns & " values (" & idAccount & ")"
	  ConnMsde.execute qins
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
if LMP_SoloLettura="" then  
   NameLoaded = NameLoaded & ";LMP_ImptFido,FLZ;LMP_ImptEstratto,FLZ"
   NameRangeN = ""
   NameRangeN = NameRangeN &  "LMP_ImptImpeFido;LMP_ImptFido;0;9999999"
   NameRangeN = NameRangeN &  ";LMP_ImptImpeEstratto;LMP_ImptEstratto;0;9999999"
end if 
ContaModLMP=0
%>
   <input type="hidden" name="LMP_UPDATE" id="LMP_UPDATE" value="<%=LMP_SoloLettura%>">
   <div class="table-responsive"><table class="table"><tbody>
      <thead>
      <tr>
          <th scope="col">Tipo Pagamento</th>
          <th scope="col">Attivo</th>
          <th scope="col">Impt.totale &euro;</th>         
          <th scope="col">Impt.Impegnato &euro;</th>
		  <th scope="col">Impt.Utilizzato &euro;</th>
          <th scope="col">Impt.Disponibile &euro;</th>
          <th scope="col">Impt.Validazione</th>
      </tr>
      </thead>
      <tr scope="col">
      <%
	  LMP_opt="Borsellino"
	  FlagAttivo = ""
	  if FlagBorsellino = 1 then 
	     FlagAttivo=" checked "
	  end if 
	  FlagReadonly=""
	  if LMP_SoloLettura<>"" then 
	     FlagReadonly=" disabled "
	  end if 
	  if FlagBorsellino=1 or OpDocAmm<>"L" then
         ContaModLMP=ContaModLMP+1  
	  %> 	  
         <td>
	        <input class="form-control" type="text" readonly value="<%=LMP_opt%>">
	     </td>
		 <td>
		    <input id="LMP_check<%=LMP_opt%><%=l_Id%>" <%=FlagAttivo%> name="LMP_check<%=LMP_opt%><%=l_Id%>" 
				type="checkbox" value = "S" <%=FlagReadonly%> class="big-checkbox" >
		 </td>
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptBorsellino,2)%>">
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptBorsellinoImpe,2)%>">
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptBorsellinoUtil,2)%>">
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptBorsellinoDisp,2)%>">
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptBorsellinoValo,2)%>">
	     </td>		 
	  </tr>
      <%end if %>
      
      <%
	  LMP_opt="Fido"
	  FlagAttivo = ""
	  if FlagFido = 1 then 
	     FlagAttivo=" checked "
	  end if 
	  FlagReadonly=""
	  if LMP_SoloLettura<>"" then 
	     FlagReadonly=" disabled "
	  end if 
	  if LMP_SoloLettura<>"" then 
	     impt=InsertPoint(ImptFido,2)
	  else
	     Impt=ImptFido
	  end if 
	  if FlagFido=1 or OpDocAmm<>"L" then 
         ContaModLMP=ContaModLMP+1  
	  %> <tr scope="col">	  
         <td>
	        <input class="form-control" type="text" readonly value="<%=LMP_opt%>">
	     </td>
		 <td>
		    <input id="LMP_check<%=LMP_opt%><%=l_Id%>" <%=FlagAttivo%> name="LMP_check<%=LMP_opt%><%=l_Id%>" 
				type="checkbox" value = "S" <%=FlagReadonly%> class="big-checkbox" >
		 </td>
         <td>
	        <input class="form-control text-right" type="text" value="<%=impt%>" <%=LMP_SoloLettura%>
			id="LMP_Impt<%=LMP_opt%><%=l_Id%>"  <%=FlagAttivo%> name="LMP_Impt<%=LMP_opt%><%=l_Id%>">
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptFidoImpe,2)%>">
			<input type="hidden" id="LMP_ImptImpe<%=LMP_opt%><%=l_Id%>" name="LMP_ImptImpe<%=LMP_opt%><%=l_Id%>"
			value="<%=ImptFidoImpe%>">
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptFidoUtil,2)%>">
			<input type="hidden" id="LMP_ImptUtil<%=LMP_opt%><%=l_Id%>" name="LMP_ImptUtil<%=LMP_opt%><%=l_Id%>"
			value="<%=ImptFidoUtil%>">
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptFidoDisp,2)%>">
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptFidoValo,2)%>">
	     </td>		 
	  </tr>           
      <%end if %>
     
      <%
	  LMP_opt="Estratto"
	  FlagAttivo = ""
	  if FlagEstratto = 1 then 
	     FlagAttivo=" checked "
	  end if 
	  FlagReadonly=""
	  if LMP_SoloLettura<>"" then 
	     FlagReadonly=" disabled "
	  end if 
	  if LMP_SoloLettura<>"" then 
	     impt=InsertPoint(ImptEstratto,2)
	  else
	     Impt=ImptEstratto
	  end if 
	  if FlagEstratto=1 or OpDocAmm<>"L" then 
         ContaModLMP=ContaModLMP+1  
	  %>  <tr scope="col">	  
         <td>
	        <input class="form-control" type="text" readonly value="<%=LMP_opt%>">
	     </td>
		 <td>
		    <input id="LMP_check<%=LMP_opt%><%=l_Id%>" <%=FlagAttivo%> name="LMP_check<%=LMP_opt%><%=l_Id%>" 
				type="checkbox" value = "S"  <%=FlagReadonly%> class="big-checkbox" >
		 </td>
         <td>
	        <input class="form-control  text-right" type="text" value="<%=Impt%>"
			id="LMP_Impt<%=LMP_opt%><%=l_Id%>" <%=LMP_SoloLettura%> <%=FlagAttivo%> name="LMP_Impt<%=LMP_opt%><%=l_Id%>"
			>
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptEstrattoImpe,2)%>">
			<input type="hidden" id="LMP_ImptImpe<%=LMP_opt%><%=l_Id%>" name="LMP_ImptImpe<%=LMP_opt%><%=l_Id%>"
			value="<%=ImptEstrattoImpe%>">
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptEstrattoUtil,2)%>">
			<input type="hidden" id="LMP_ImptUtil<%=LMP_opt%><%=l_Id%>" name="LMP_ImptUtil<%=LMP_opt%><%=l_Id%>"
			value="<%=ImptEstrattoUtil%>">
	     </td>		 		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptEstrattoDisp,2)%>">
	     </td>		 
         <td>
	        <input class="form-control text-right" type="text" readonly value="<%=InsertPoint(ImptEstrattoValo,2)%>">
	     </td>		 
	  </tr>  
     <%end if %>

   </tbody></table></div>
	 <%if ContaModLMP=0 then
          MsgErrore="Nessuna modalita' di pagamento prevista per il cliente "
	 %>
	    <!--#include virtual="/gscVirtual/include/showError.asp"--> 

	 <%   MsgErrore=""
	 end if %>
