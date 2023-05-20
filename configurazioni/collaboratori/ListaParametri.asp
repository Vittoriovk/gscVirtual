

<%

'IdAccountModPagCliente = AccountCliente
'IdAccountRequest       = AccountDaConfigurare 

err.clear

Set RsDoc = Server.CreateObject("ADODB.Recordset")
RsDoc.CursorLocation = 3 

flagParams=false 

if cdbl(tmpIdAccountRequest)>0 then 
   'leggo i pagamenti disponibili per cliente 
   MySqlDoc = ""
   MySqlDoc = MySqlDoc & " select A.*,Isnull(B.ValoreParametro,'') as ValPar "
   MySqlDoc = MySqlDoc & " from tipoParametro A left join AccountTipoParametro B "
   MySqlDoc = MySqlDoc & " on  A.IdTipoParametro = B.IdTipoParametro "
   MySqlDoc = MySqlDoc & " and B.IdAccount = " & tmpIdAccountRequest
   MySqlDoc = MySqlDoc & " where A.IdTipoParametro in (" & listaParametri & ")"   
   MySqlDoc = MySqlDoc & " order By A.DescTipoParametro"

   RsDoc.Open MySqlDoc, ConnMsde
   if err.number=0 then
      do while not RsDoc.EOF
	     ValPar = RsDoc("ValPar")

	     if RsDoc("IdTipoParametro")="VAL_COB" then
		    FlagcheckVAL_COB_Back=""
			FlagcheckVAL_COB_Coll=""
			if instr(ValPar,"BACKO")>0 then 
			   FlagcheckVAL_COB_Back=" checked "
			end if 
			if instr(ValPar,"COLL")>0 then 
			   FlagcheckVAL_COB_Coll=" checked "
			end if 
			if FlagcheckVAL_COB_Back="" and FlagcheckVAL_COB_Coll="" then 
			   FlagcheckVAL_COB_Back=" checked "
			end if  
			
		 %>
		    <input type="hidden" name="campo_VAL_COB" value="checkVAL_COB0">
            <div class="row">
	           <div class="col-5">
		          <p class="font-weight-bold">Gestione validazione Coobbligati a carico di (default Back Office)</p>
	           </div>
	   
               <div class="col-6">
                   <input id="checkVAL_COB0" <%=FlagcheckVAL_COB_Back%> name="checkVAL_COB0" 
			       type="checkbox" value = "BACKO" class="big-checkbox" >
				   <span class="font-weight-bold">&nbsp;Back Office</span>
				   &nbsp;&nbsp;&nbsp;
                   <input id="checkVAL_COB0" <%=FlagcheckVAL_COB_Coll%> name="checkVAL_COB0" 
			       type="checkbox" value = "COLL" class="big-checkbox" >
				   <span class="font-weight-bold">&nbsp;Collaboratore</span>
               </div>
			</div>
		 <%
		 end if 
	     if RsDoc("IdTipoParametro")="VAL_ATI" then
		    FlagcheckVAL_ATI_Back=""
			FlagcheckVAL_ATI_Coll=""
			if instr(ValPar,"BACKO")>0 then 
			   FlagcheckVAL_ATI_Back=" checked "
			end if 
			if instr(ValPar,"COLL")>0 then 
			   FlagcheckVAL_ATI_Coll=" checked "
			end if 
			if FlagcheckVAL_ATI_Back="" and FlagcheckVAL_ATI_Coll="" then 
			   FlagcheckVAL_ATI_Back=" checked "
			end if  		 
		 %>
		    <input type="hidden" name="campo_VAL_ATI" value="checkVAL_ATI0">
            <div class="row">
	           <div class="col-5">
		          <p class="font-weight-bold">Gestione validazione A.T.I. a carico di (default Back Office)</p>
	           </div>			
               <div class="col-6">
                   <input id="checkVAL_ATI0" <%=FlagcheckVAL_ATI_Back%> name="checkVAL_ATI0" 
			       type="checkbox" value = "BACKO" class="big-checkbox" >
				   <span class="font-weight-bold">&nbsp;Back Office</span>
				   &nbsp;&nbsp;&nbsp;
                   <input id="checkVAL_ATI0" <%=FlagcheckVAL_ATI_Coll%> name="checkVAL_ATI0" 
			       type="checkbox" value = "COLL" class="big-checkbox" >
				   <span class="font-weight-bold">&nbsp;Collaboratore</span>
               </div>
			</div>
		 <%
		 end if 
	     if RsDoc("IdTipoParametro")="ASS_PRO" then
		    FlagcheckASS_PRO=""
			if ValPar="S" then 
			   FlagcheckASS_PRO=" checked "
			end if 
		 %>
		    <input type="hidden" name="campo_ASS_PRO" value="checkASS_PRO0">
            <div class="row">
	           <div class="col-5">
		          <p class="font-weight-bold">Pu&ograve; Associare prodotti a collaboratori e clienti (default NO)</p>
	           </div>			
               <div class="col-6">
                   <input id="checkASS_PRO0" <%=FlagcheckASS_PRO%> name="checkASS_PRO0" 
			       type="checkbox" value = "S" class="big-checkbox" >
				   <span class="font-weight-bold">&nbsp;SI</span>
               </div>
			</div>
		 <%

		 end if
		 
	     RsDoc.MoveNext 
	  loop 
   end if 
   RsDoc.close 
   err.clear 
end if 
%>
 


