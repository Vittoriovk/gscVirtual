<%
  NomePagina="TrasferisciFornitore.asp"
  titolo="Dettaglio Richiesta Di Affidamento Cliente per Fornitore"
  default_check_profile="BackO,Coll,Clie"
  act_call_down = CryptAction("CALL_DOWN")
  act_call_uplo = CryptAction("CALL_UPLO")
  act_call_firm = CryptAction("CALL_FIRM")
  act_call_back = CryptAction("CALL_BACK")
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAffidamento.asp"-->
<!--#include virtual="/gscVirtual/modelli/FunctionEvento.asp"-->
<%
  livelloPagina="00"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include virtual="/gscVirtual/include/head.asp"-->
<!-- Custom styles for this template -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">

</head>
<script>
function upload(id)
{
    xx=ImpostaValoreDi("Oper","<%=act_call_uplo%>");
	xx=ImpostaValoreDi("ItemToRemove",id);
    document.Fdati.submit();  
}
function uploadd(id)
{
    xx=ImpostaValoreDi("Oper","<%=act_call_down%>");
	xx=ImpostaValoreDi("ItemToRemove",id);
    document.Fdati.submit();  
}
function inviaFirma()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_firm%>");
    document.Fdati.submit();  
}
function inviaBack()
{
    xx=ImpostaValoreDi("Oper","<%=act_call_back%>");
    document.Fdati.submit();  
}
</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
<%
  PaginaReturn = ""
  if FirstLoad then 
     PaginaReturn               = getCurrentValueFor("PaginaReturn")
     IdAffidamentoRichiesta     = cdbl("0" & getCurrentValueFor("IdAffidamentoRichiesta"))
     IdAffidamentoRichiestaComp = cdbl("0" & getCurrentValueFor("IdAffidamentoRichiestaComp"))
  else
     PaginaReturn               = getValueOfDic(Pagedic,"PaginaReturn")
     IdAffidamentoRichiesta     = getValueOfDic(Pagedic,"IdAffidamentoRichiesta")
     IdAffidamentoRichiestaComp = getValueOfDic(Pagedic,"IdAffidamentoRichiestaComp")
  end if 

  if cdbl("0" & IdAffidamentoRichiesta)=0 then 
     qSel = "Select * from AffidamentoRichiestaComp Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp 
     IdAffidamentoRichiesta = LeggiCampo(qSel,"IdAffidamentoRichiesta")
  end if  

 
  Oper = DecryptAction(Oper)
 
   if Oper="CALL_UPLO" then 
      xx=RemoveSwap()
      itemId = Request("ItemToRemove") 
      IdDocTraD = LeggiCampo("Select * from Documento Where IdDocumentoInterno='MODULO_TRASF_D'","IdDocumento")
      Session("swap_IdTabella")          = "AFFIDAMENTO"
      Session("swap_IdTabellaKeyInt")    = IdAffidamentoRichiestaComp
      Session("swap_IdTabellaKeyString") = ""
      Session("swap_IdUpload")           = Cdbl("0" & itemId)
      Session("swap_FlagDescEstesa")     = "N"
      Session("swap_FlagDataScadenza")   = "N"
      Session("swap_IdDocumentoToLoad")  = IdDocTraD 
      Session("swap_PaginaReturn")  = "Configurazioni/clienti/Affidamento/" & nomePagina 
      response.redirect virtualPath & "Configurazioni/clienti/Affidamento/DocumentoUpload.asp"
      response.end 
   end if 
   if Oper="CALL_DOWN" then 
      xx=RemoveSwap()
      itemId = Request("ItemToRemove") 
      IdDocTraU = LeggiCampo("Select * from Documento Where IdDocumentoInterno='MODULO_TRASF_U'","IdDocumento")
      Session("swap_IdTabella")          = "AFFIDAMENTO"
      Session("swap_IdTabellaKeyInt")    = IdAffidamentoRichiestaComp
      Session("swap_IdTabellaKeyString") = ""
      Session("swap_IdUpload")           = Cdbl("0" & itemId)
      Session("swap_FlagDescEstesa")     = "N"
      Session("swap_FlagDataScadenza")   = "N"
      Session("swap_IdDocumentoToLoad")  = IdDocTraU 
      Session("swap_PaginaReturn")  = "Configurazioni/clienti/Affidamento/" & nomePagina 
      response.redirect virtualPath & "Configurazioni/clienti/Affidamento/DocumentoUpload.asp"
      response.end 
   end if 
   
   if Oper="CALL_FIRM" then
      IdAccountDocumento = 0
      IdDocumento = LeggiCampo("Select * from Documento Where IdDocumentoInterno='MODULO_TRASF_U'","IdDocumento")
      IdDocumento = TestNumeroPos(IdDocumento)	 
 
      qSel = ""
	  qSel = qSel & " select * from AffidamentoRichiestaCompDoc "
      qSel = qSel & " Where IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp
      qSel = qSel & " and IdDocumento=" & IdDocumento
      'response.write qSEl 
	
      tmp = cdbl("0" & LeggiCampo(qSel,"IdDocumento"))
      if cdbl(tmp)=0 then 
         qIns = ""
         qIns = qIns & "Insert into AffidamentoRichiestaCompDoc ("
         qIns = qIns & " IdAffidamentoRichiestaComp"
         qIns = qIns & ",IdDocumento"
         qIns = qIns & ",flagObbligatorio"
         qIns = qIns & ",FlagDataScadenza"
         qIns = qIns & ",IdAccountDocumento) "
         qIns = qIns & " values ("
         qIns = qIns & " " & IdAffidamentoRichiestaComp
         qIns = qIns & "," & IdDocumento
         qIns = qIns & ",1"
         qIns = qIns & ",1"
         qIns = qIns & "," & IdAccountDocumento & " ) "
	     'response.write qIns 
         ConnMsde.execute qIns 
	  end if 
	  'response.end    
      qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
      qUpd = qUpd & " IdFlussoProcesso = 'CAC_FORNITORE' "
      qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
      qUpd = qUpd & " and   IdFlussoProcesso = 'CAM_FORNITORE' "
      ConnMsde.execute qUpd 
      XX=createEvento("AFFI","RICH",Session("LoginIdAccount"),"Richiesta Firma Trasferimento","AffidamentoRichiestaComp","IdAffidamentoRichiestaComp=" & IdAffidamentoRichiestaComp,true,0)
      response.redirect VirtualPath & PaginaReturn
	  response.end 
   end if 
   if Oper="CALL_BACK" then
      qUpd = qUpd & " Update AffidamentoRichiestaComp set  "
      qUpd = qUpd & " IdFlussoProcesso = 'GES_FORNITORE' "
      qUpd = qUpd & " Where IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
      qUpd = qUpd & " and   IdFlussoProcesso = 'CAC_FORNITORE' "
      ConnMsde.execute qUpd 
      response.redirect VirtualPath & PaginaReturn
      response.end 
   end if    
'selezionato dal cassetto lo associo 

  xx=setValueOfDic(Pagedic,"PaginaReturn"               ,PaginaReturn)
  xx=setValueOfDic(Pagedic,"IdAccountCliente"           ,IdAccountCliente)
  xx=setValueOfDic(Pagedic,"IdAffidamentoRichiestaComp" ,IdAffidamentoRichiestaComp)
  xx=setCurrent(NomePagina,livelloPagina) 

  Oggi = Dtos() 
  Set Rs = Server.CreateObject("ADODB.Recordset")
 

  MySql = ""
  MySql = MySql & " Select A.*,B.DescStatoServizio,B.FlagStatoFinale,C.DescCompagnia "
  MySql = MySql & " from AffidamentoRichiestaComp A,StatoServizio B,Compagnia C"
  MySql = MySql & " Where A.IdAffidamentoRichiestaComp = " & IdAffidamentoRichiestaComp
  MySql = MySql & " and   A.IdCompagnia = C.IdCompagnia"
  MySql = MySql & " and   A.IdStatoAffidamento = B.IdStatoServizio"
   'response.write MySql
  err.clear 

  Rs.CursorLocation = 3 
  Rs.Open MySql, ConnMsde
  IdAccountCliente         = Rs("IdAccountCliente")
  DescCliente              = LeggiCampo("Select * from Account Where idAccount=" & IdAccountCliente,"Nominativo" )
  idCompagniaComp          = Rs("IdCompagnia")
  DescCompagniaComp        = Rs("DescCompagnia")
  DataRichiestaComp        = Rs("DataRichiesta")
  IdStatoAffidamentoComp   = Rs("IdStatoAffidamento")
  DescStatoServizioComp    = Rs("DescStatoServizio")
  FlagStatoFinale          = Rs("FlagStatoFinale")
  DataChiusuraComp         = Rs("DataChiusura")
  NoteAffidamentoComp      = Rs("NoteAffidamento")
  NoteAffidamentoClie      = Rs("NoteAffidamentoCliente")
  ValidoDalComp            = Rs("ValidoDal")
  ValidoAlComp             = Rs("ValidoAl")
  IdFornitore              = Rs("IdFornitore")
  if Cdbl(IdFornitore)>0 then 
     DescFornitore=LeggiCampo("Select * from Fornitore Where IdFornitore=" & Idfornitore,"DescFornitore")
  else
     DescFornitore=""
  end if 
  IdFlussoProcesso = rs("IdFlussoProcesso")
  Rs.close
  
   Set DizProcesso = CreateObject("Scripting.Dictionary")
   esito=getInfoProcessoAffi(DizProcesso,IdStatoAffidamentoComp,IdFlussoProcesso)
   if esito = true then 
      descProcesso = " - " & GetDiz(DizProcesso,"DescrizioneUtente")
   else
      descProcesso = "" 
   end if 
  DescLoaded=""
  NumRec  = 0
  MsgNoData  = ""
  
%>

<div class="d-flex" id="wrapper">
    <%
      callP=VirtualPath & "bar/" & Session("sideBar_" & Session("LoginIdAccount")) 
      Server.Execute(callP) 
    %>
    <!-- Page Content -->
    <div id="page-content-wrapper">
    <%
      callP=VirtualPath & "bar/" & Session("TopBar_" & Session("LoginIdAccount")) 
      Server.Execute(callP) 
    %>    

        <div class="container-fluid">
            <form name="Fdati" Action="<%=NomePagina%>" method="post">
 
            <div class="row">
                <%
                if PaginaReturn<>"" then 
                   RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"
                %>
                   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <% end if %>
                <div class="col-11"><h3>Richiesta Trasferimento Fornitore Per Affidamento</h3>
                </div>
            </div>
            <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
            <div class="row">
               <div class="col-4">
                  <div class="form-group ">
                     <%xx=ShowLabel("Utente")%>
                     <input type="text" readonly class="form-control" value="<%=DescCliente%>" >
                  </div>        
               </div>
               <div class="col-4">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Compagnia")%>
                     <input type="text" readonly class="form-control" value="<%=DescCompagniaComp%>" >
                  </div>        
               </div>
            </div>
            
            <div class="row">			   
               <div class="col-2">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Richiesta Del")%>
                     <input type="text" readonly class="form-control" value="<%=Stod(DataRichiestaComp)%>" >
                  </div>        
               </div>
               <div class="col-4">
                  <div class="form-group font-weight-bold">
                     <%xx=ShowLabel("Stato Richiesta")%>
                     <input type="text" readonly class="form-control" value="<%=DescStatoServizioComp & descProcesso%>" >
                  </div>        
               </div>
               
            </div>
            <%If IsBackOffice() then %>
            <div class="row">
               <div class="col-4">
                  <div class="form-group">
                     <%xx=ShowLabel("Fornitore")%>
                     <input type="text" readonly class="form-control" value="<%=DescFornitore%>" >
                  </div>                       
               </div>             

            </div> 
			<%end if %>
			<%
			ShowDownload    = true 
			ShowUpload      = true 
			if IsBackOffice() then 
			   showUpload = false
			end if
            idTabella       = "AFFIDAMENTO"
			IdTabellaKeyInt = IdAffidamentoRichiestaComp
			IdDocTraU = 0 
			IdDocTraD = 0
			%>
			<!--#include virtual="/gscVirtual/configurazioni/documenti/MostraTrasferimento.asp"-->

			
			<%
			  flagInviaCli = false 
			  flagInviaBac = false 
			  flagComandi  = false 
			  'response.write IdFlussoProcesso & IdDocTraU
			  if readonly="" then
			     If isBackOffice() then 
                    if cdbl(IdDocTraD)>0 and idFlussoProcesso="CAM_FORNITORE" then 
					   flagComandi  = true 
					   flagInviaCli = true  
					end if 
				 else 
                    if cdbl(IdDocTraU)>0 and idFlussoProcesso="CAC_FORNITORE" then 
					   flagComandi  = true 
					   flagInviaBac = true  
					end if 

			     end if
			  end if
			%>
			
			<%if flagComandi then %>
			<br>
			<div class="row">
			   <%if flagInviaCli = true then %>
                   <div class="col-1">
                   </div>  
			   
                   <div class="col-2">
                     <div class="form-group ">
                         <button type="button" onclick="inviaFirma()" class="btn btn-success">Invia Per Firma</button>
                     </div>               
                   </div>  
			   
			   <%end if %>
			   <%if flagInviaBac = true then %>
                   <div class="col-1">
                   </div>  
			   
                   <div class="col-2">
                     <div class="form-group ">
                         <button type="button" onclick="inviaBack()" class="btn btn-success">Invia Richiesta</button>
                     </div>               
                   </div>  
			   
			   <%end if %>			   
			</div>
			<%end if %>
           
            <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
            <!--#include virtual="/gscVirtual/include/paginazione.asp"-->

            </form>
        </div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->

<!--#include virtual="/gscVirtual/include/scripts.asp"-->

  <!-- Menu Toggle Script -->
  <script>
    $("#menu-toggle").click(function(e) {
      e.preventDefault();
      $("#wrapper").toggleClass("toggled");
    });
  </script>
  <script>
    $(document).ready(function(){
      $('[data-toggle="tooltip" = Rs("")').tooltip();   
    });
  </script>
  <script>
$('.btn').onClick(function(e){
  e.preventDefault();
});  
</script>
</body>

</html>
