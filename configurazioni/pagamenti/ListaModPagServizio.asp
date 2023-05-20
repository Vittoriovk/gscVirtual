
<script>
function LMPS_CheckPagamento(idAccCli,idAccReq)
{
   // controllo se è stato selezionato un pagamento per il cliente.
   var prefix = "ListaModPagServizio";
   var whichCL = ValoreDi("LMPS_WHICH_CLI");
   var whichRE = ValoreDi("LMPS_WHICH_REQ");
   
   var cliBors = false;
   var cliFido = false;
   var cliEstr = false;
   var reqBors = false;
   var reqFido = false;
   var reqEstr = false;
	  
   if (whichCL.includes("|BORS|"))
      cliBors = IsChecked(prefix + "BORS_" + idAccCli);
   if (whichCL.includes("|FIDO|"))
      cliFido = IsChecked(prefix + "FIDO_" + idAccCli);
   if (whichCL.includes("|ESTR|"))
      cliEstr = IsChecked(prefix + "ESTR_" + idAccCli);
   if (idAccReq>0) {
      if (whichRE.includes("|BORS|"))
         reqBors = IsChecked(prefix + "BORS_" + idAccReq);
      if (whichRE.includes("|FIDO|"))
         reqFido = IsChecked(prefix + "FIDO_" + idAccReq);
      if (whichRE.includes("|ESTR|"))
         reqEstr = IsChecked(prefix + "ESTR_" + idAccReq);
   }
   if (cliBors==false && cliFido==false & cliEstr==false) {
	  myAlert('Selezione Pagamento','Indicare il tipo di pagamento per il cliente');
      return false;
   }
   
   if (1==0 && cliBors==true && (reqBors==true || reqFido==true || reqEstr==true)) {
	  myAlert("Selezione Pagamento","Per pagamento cliente borsellino non e' necessario indicare pagamento personale.");
      return false;
   }
   // controllo se e' stato selezionato anche la modalita' del richiedente 
   if (idAccReq>0 && cliBors==true && reqBors==false && reqFido==false && reqEstr==false) {
	  myAlert("Selezione Pagamento","Indicare la modalita' di pagamento personale.");
      return false;
   }
   
   // deve essere implementata nel modulo chiamante 
   LMPS_CheckPagamentoOk();
}

</script>

<%
'input
'OpDocAmm               = Operazioni ammesse 
'                         S = selezione
'                         L = solo lista attivi

'IdAccountModPagCliente = AccountCliente
'IdAccountRequest       = AccountRichiedente  : mostra i pagamenti previsti 

'output 
'countPagaClie          contatore dei pagamenti cliente 
'countPagaRich          contatore dei pagamenti richiedente 
err.clear

countPagaClie = 0
countPagaRich = 0 

tmpIdAccountRequest=0
LMP_SoloLettura=""
if OpDocAmm<>"S" then 
   LMP_SoloLettura = " readonly "
end if 

Set RsDoc = Server.CreateObject("ADODB.Recordset")
RsDoc.CursorLocation = 3 


'controllo se il richiedente è diverso dal cliente e non è un segnalatore
if cdbl(IdAccountModPagCliente)<>cdbl(IdAccountRequest) then 
   MySqlDoc = ""
   MySqlDoc = MySqlDoc & "select * from Collaboratore Where IdAccount = " & IdAccountRequest
   TipoColl = ucase(Leggicampo(MySqlDoc,"IdTipoCollaboratore"))
   if TipoColl="" or TipoColl="SEGN" then 
      tmpIdAccountRequest = 0
   else
      tmpIdAccountRequest = IdAccountRequest
   end if 
else 
   tmpIdAccountRequest=0
end if 
'inizializzo variabili cliente 
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

'inizializzo variabili Richiedente  
FlagBorsellinoR = 0
ImptBorsellinoR = 0
ImptBorsellinoImpeR = 0
ImptBorsellinoUtilR = 0
ImptBorsellinoDispR = 0
ImptBorsellinoValoR = 0
FlagFidoR = 0
ImptFidoR = 0
ImptFidoImpeR = 0
ImptFidoUtilR = 0
ImptFidoDispR = 0
ImptFidoValoR = 0
FlagEstrattoR = 0
ImptEstrattoR = 0
ImptEstrattoImpeR = 0
ImptEstrattoUtilR = 0
ImptEstrattoDispR = 0
ImptEstrattoValoR = 0
'aggiungo il wallet per pagamenti su altre piattaforme 

'leggo i pagamenti disponibili per cliente 
MySqlDoc = ""
MySqlDoc = MySqlDoc & "select * from AccountModPag Where IdAccount = " & IdAccountModPagCliente

RsDoc.Open MySqlDoc, ConnMsde
if err.number=0 then
   if not RsDoc.EOF then 
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

if cdbl(tmpIdAccountRequest)>0 then 
   'leggo i pagamenti disponibili per il richiedente  
   MySqlDoc = ""
   MySqlDoc = MySqlDoc & "select * from AccountModPag Where IdAccount = " & tmpIdAccountRequest

   RsDoc.Open MySqlDoc, ConnMsde
   if err.number=0 then
      if not RsDoc.EOF then 
         FlagBorsellinoR     = RsDoc("FlagBorsellino")
         ImptBorsellinoR     = RsDoc("ImptBorsellino")
         ImptBorsellinoImpeR = RsDoc("ImptBorsellinoImpe")
		 ImptBorsellinoUtilR = RsDoc("ImptBorsellinoUtil")
         ImptBorsellinoDispR = RsDoc("ImptBorsellinoDisp")
         ImptBorsellinoValoR = RsDoc("ImptBorsellinoValo")
         FlagFidoR           = RsDoc("FlagFido")
         ImptFidoR           = RsDoc("ImptFido")
         ImptFidoImpeR       = RsDoc("ImptFidoImpe")
		 ImptFidoUtilR       = RsDoc("ImptFidoUtil")
         ImptFidoDispR       = RsDoc("ImptFidoDisp")
         ImptFidoValoR       = RsDoc("ImptFidoValo")
         FlagEstrattoR       = RsDoc("FlagEstratto")
         ImptEstrattoR       = RsDoc("ImptEstratto")
         ImptEstrattoImpeR   = RsDoc("ImptEstrattoImpe")
		 ImptEstrattoUtilR   = RsDoc("ImptEstrattoUtil")
         ImptEstrattoDispR   = RsDoc("ImptEstrattoDisp")
         ImptEstrattoValoR   = RsDoc("ImptEstrattoValo")	  
	  end if 
   end if 
   RsDoc.close 
end if 


ContaModLMP=0
%>
   <input type="hidden" name="LMP_UPDATE" id="LMP_UPDATE" value="<%=LMP_SoloLettura%>">
   <div class="table-responsive"><table class="table"><tbody>
     
      <%
	  'se non sono cliente e ho dettagli metto il riferimento per cliente 
	  if (FlagBorsellino=1 or FlagFido=1 or FlagEstratto=1) then 
	     if tmpIdAccountRequest>0 then 
	  %>
	  <tr>
	      <td colspan="7" scope="col">
			   <div class="bg-primary text-center text-white font-weight-bold">
			   Pagamenti per cliente 
			   </div>

		  
		  </td>
	  </tr>
	  <% end if %>
      <tr>
	      <th scope="col">Sel.</th>
          <th scope="col">Pagamento</th>
          <th scope="col">Impt.totale &euro;</th>         
          <th scope="col">Impt.Impegnato &euro;</th>
		  <th scope="col">Impt.Utilizzato &euro;</th>
          <th scope="col">Impt.Disponibile &euro;</th>
          <th scope="col">Impt.Validazione</th>
      </tr>	  

	  <%
	  end if 
	  WhichCLI="|"
	  TipoUten="CLIE"
	  if FlagBorsellino=1 then
	     LMP_opt="Borsellino"
	     TipoPaga = "BORS"
		 CodePaga = IdAccountModPagCliente
	     ImptTota = ImptBorsellino
		 ImptImpe = ImptBorsellinoImpe
		 ImptUtil = ImptBorsellinoUtil
		 ImptDisp = ImptBorsellinoDisp
		 ImptValo = ImptBorsellinoValo
		 checkOk  = false 
		 WhichCLI=WhichCLI & TipoPaga & "|"
	  %>
	  <!--#include file="ListaModPagServizioRiga.asp"-->
	  <%
	    if CheckOk = true then
		   countPagaClie = countPagaClie + 1
		end if 
	  end if 

	  if FlagFido=1 then
	     LMP_opt="Fido"
	     TipoPaga = "FIDO"
		 CodePaga = IdAccountModPagCliente
	     ImptTota = ImptFido
		 ImptImpe = ImptFidoImpe
		 ImptUtil = ImptFidoUtil
		 ImptDisp = ImptFidoDisp
		 ImptValo = ImptFidoValo
		 checkOk  = false
		 WhichCLI=WhichCLI & TipoPaga & "|"
	  %>
	  <!--#include file="ListaModPagServizioRiga.asp"-->
	  <%
	    if CheckOk = true then
		   countPagaClie = countPagaClie + 1
		end if 	  
	  end if 
	  if FlagEstratto=1 then
	     LMP_opt="Estratto Conto Cliente"
	     TipoPaga = "ESTR"
		 CodePaga = IdAccountModPagCliente
	     ImptTota = ImptEstratto
		 ImptImpe = ImptEstrattoImpe
		 ImptUtil = ImptEstrattoUtil
		 ImptDisp = ImptEstrattoDisp
		 ImptValo = ImptEstrattoValo
		 checkOk  = false
		 WhichCLI=WhichCLI & TipoPaga & "|"
	  %>
	  <!--#include file="ListaModPagServizioRiga.asp"-->
	  <%
	    if CheckOk = true then
		   countPagaClie = countPagaClie + 1
		end if 		  
	  end if 

	  'se non sono cliente e ho dettagli metto il riferimento per cliente 
	  if (FlagBorsellinoR=1 or FlagFidoR=1 or FlagEstrattoR=1) and tmpIdAccountRequest>0 then 
	  %>
	  <tr>
	      <td colspan="7" scope="col">
			   <div class="bg-primary text-center text-white font-weight-bold">
			   Pagamenti personali 
			   </div>
		  </td>
	  </tr>
      <tr>
	      <th scope="col">Sel.</th>
          <th scope="col">Pagamento</th>
          <th scope="col">Impt.totale &euro;</th>         
          <th scope="col">Impt.Impegnato &euro;</th>
		  <th scope="col">Impt.Utilizzato &euro;</th>
          <th scope="col">Impt.Disponibile &euro;</th>
          <th scope="col">Impt.Validazione</th>
      </tr>	  

	  <%
	  end if 

      TipoUten="REQU"
	  WhichREQ="|"
	  if FlagBorsellinoR=1 then
	     LMP_opt="Borsellino"
	     TipoPaga = "BORS"
		 CodePaga = tmpIdAccountRequest
	     ImptTota = ImptBorsellinoR
		 ImptImpe = ImptBorsellinoImpeR
		 ImptUtil = ImptBorsellinoUtilR
		 ImptDisp = ImptBorsellinoDispR
		 ImptValo = ImptBorsellinoValoR
		 checkOk  = false
		 WhichREQ=WhichREQ & TipoPaga & "|"
	  %>
	  <!--#include file="ListaModPagServizioRiga.asp"-->
	  <%
	    if CheckOk = true then
		   countPagaRich = countPagarich + 1
		end if 	  
	  end if 

	  if FlagFidoR=1 then
	     LMP_opt="Fido"
	     TipoPaga = "FIDO"
		 CodePaga = tmpIdAccountRequest
	     ImptTota = ImptFidoR
		 ImptImpe = ImptFidoImpeR
		 ImptUtil = ImptFidoUtilR
		 ImptDisp = ImptFidoDispR
		 ImptValo = ImptFidoValoR
		 checkOk  = false
		 WhichREQ=WhichREQ & TipoPaga & "|"
	  %>
	  <!--#include file="ListaModPagServizioRiga.asp"-->
	  <%
	    if CheckOk = true then
		   countPagaRich = countPagarich + 1
		end if	  
	  end if 
	  if FlagEstrattoR=1 then
	     LMP_opt="Estratto Conto"
	     TipoPaga = "ESTR"
		 CodePaga = tmpIdAccountRequest
	     ImptTota = ImptEstrattoR
		 ImptImpe = ImptEstrattoImpeR
		 ImptUtil = ImptEstrattoUtilR
		 ImptDisp = ImptEstrattoDispR
		 ImptValo = ImptEstrattoValoR
		 checkOk  = false
		 WhichREQ=WhichREQ & TipoPaga & "|"		 
	  %>
	  <!--#include file="ListaModPagServizioRiga.asp"-->
	  <%
	    if CheckOk = true then
		   countPagaRich = countPagarich + 1
		end if	  
	  end if 
	  %>
	  
   </tbody></table></div>
   <input type="hidden" name="LMPS_WHICH_CLI" id="LMPS_WHICH_CLI" value="<%=WhichCLI%>">
   <input type="hidden" name="LMPS_WHICH_REQ" id="LMPS_WHICH_REQ" value="<%=WhichREQ%>">
   
     <% if canSelModPaga then 
   
	       if countPagaClie=0 then
              MsgErrore="Nessuna modalita' di pagamento prevista per il cliente "
	 %>
	          <!--#include virtual="/gscVirtual/include/showError.asp"--> 

	 <%       MsgErrore=""
	      else %>
        <div class="row">
          <div class="col-2"><br>
              <button type="button" onclick="LMPS_CheckPagamento(<%=IdAccountModPagCliente%>,<%=tmpIdAccountRequest%>)" class="btn btn-warning">Paga</button>
          </div>
        </div>
	 
	 <%   end if  
	   end if %>
	 
	 
	 
	 

