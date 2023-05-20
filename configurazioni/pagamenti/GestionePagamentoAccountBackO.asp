<%
  NomePagina="GestionePagamentoAccountBackO.asp"
  titolo="Movimentazione Borsellino "
  default_check_profile = "BackO"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/FunMailWithAttach.asp"-->
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
<script language="JavaScript">

function localRicarica()
{
    document.Fdati.submit();

}

function localFun(op,id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    ImpostaValoreDi("Oper",op);
    document.Fdati.submit();

}

</script>

<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

 
<%

  NameLoaded=""
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdCliente=0
  if FirstLoad then 
     PaginaReturn  = getCurrentValueFor("PaginaReturn") 
	 IdTipoCredito = getCurrentValueFor("IdTipoCredito")
	 IdTipoUtente  = getCurrentValueFor("IdTipoUtente")
  else
     PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
	 IdTipoCredito = Request("ListaModPag")
	 IdTipoUtente  = Request("ListaAccount")
   end if 
   
   if Oper="CALL_CAR" then 
      IdMovEco = cdbl("0" & Request("ItemToRemove"))
	  if Cdbl(IdMovEco)>0 then 
         MyQ = MyQ & " update AccountMovEco "
		 MyQ = MyQ & " set IdAccountGestore = " & Session("LoginIdAccount")
         MyQ = MyQ & " Where IdAccountMovEco = " & IdMovEco 
		 ConnMsde.execute MyQ      
	  end if 
	  
   end if 
   
   if Oper="CALL_GES" then 
      Session("TimeStamp")=TimePage
      IdMovEco      = cdbl("0" & Request("ItemToRemove"))
 
      xx=RemoveSwap()
      Session("swap_IdMovEco")      = IdMovEco
      Session("swap_PaginaReturn")  = "configurazioni/pagamenti/" & NomePagina
  
      response.redirect RitornaA("configurazioni/pagamenti/PagamentoAccountModBackO.asp")
      response.end 
   End if 
   
   xx=setValueOfDic(Pagedic,"IdTipoCredito"   ,IdTipoCredito)
   xx=setValueOfDic(Pagedic,"OperTabella"     ,OperTabella)
   xx=setCurrent(NomePagina,livelloPagina) 
   DescLoaded="0"  
   
  
  %>
<% 
  'xx=DumpDic(SessionDic,NomePagina)
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
			RiferimentoA="col-1;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
			<%end if %>
                <div class="col-11"><h3>Movimentazioni Pagamenti </b> </h3>
                </div>
            </div>
	
            <%
               FlagBorsellino = 1
               FlagFido       = 1
               FlagEstratto   = 1
               FlagAction = true 
			   
                   flagchecked   = IdtipoCredito
				   IdtipoCredito = ""
				   AlllSele=""
                   BorsSele=""
                   FidoSele=""
                   EstrSele=""
				   IdtipoCredito=""
				   if Flagchecked="BORS" and flagBorsellino=1 then
				      BorsSele= " checked "
					  IdtipoCredito= Flagchecked
				   end if 
				   if Flagchecked="FIDO" and flagFido=1 then
				      FidoSele= " checked "
					  IdtipoCredito= Flagchecked
				   end if 
				   if Flagchecked="ESTR" and flagEstratto=1 then
				      EstrSele= " checked "
					  IdtipoCredito= Flagchecked
				   end if 
                   if IdtipoCredito="" then 
					  AlllSele= " checked "
					  IdtipoCredito= "ALLL"
				   end if 

			   %> 
			      <div class="row">
				     <div class="col-1">
					 </div>
				     <div class="col-2">
					    <div class="form-group font-weight-bold">
						     Tipo di Movimento
						</div>
					 </div>					 
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=AlllSele%> type="radio" 
                     	   name="ListaModPag" id="ListaModPagALLL" value="ALLL"
						   onclick="localRicarica();" >				  
 					       Tutti
					    </div>
                     </div>	
                     
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=BorsSele%> type="radio" 
                     	   name="ListaModPag" id="ListaModPagBORS" value="BORS"
						   onclick="localRicarica();" >				  
 					       Borsellino
					    </div>
                     </div>		

                     
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=FidoSele%> type="radio" 
                     	   name="ListaModPag" id="ListaModPagFIDO" value="FIDO"
						   onclick="localRicarica();">				  
 					       Fido
					    </div>
                     </div>		

                     
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=EstrSele%> type="radio" 
                     	   name="ListaModPag" id="ListaModPagESTR" value="ESTR"
						   onclick="localRicarica();">				  
 					       Estratto
					    </div>
                     </div>		
				  
			      </div>
				  <%
				   AlluSele=""
                   ClieSele=""
                   CollSele=""

				   flagchecked   = IdtipoUtente
				   IdtipoUtente=""
				   if Flagchecked="CLIE" then
				      ClieSele= " checked "
					  IdtipoUtente= Flagchecked
				   end if 
				   if Flagchecked="COLL" then
				      CollSele= " checked "
					  IdtipoUtente= Flagchecked
				   end if 
                   if IdtipoUtente="" then 
					  AlluSele= " checked "
					  IdtipoUtente= "ALLL"
				   end if 				  
                   %>
				  
			      <div class="row">
				     <div class="col-1">
					 </div>
				     <div class="col-2">
					    <div class="form-group font-weight-bold">
						     Tipo Utente
						</div>
					 </div>					 
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=AlluSele%> type="radio" 
                     	   name="ListaAccount" id="ListaAccountALLL" value="ALLL"
						   onclick="localRicarica();" >				  
 					       Tutti
					    </div>
                     </div>
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=ClieSele%> type="radio" 
                     	   name="ListaAccount" id="ListaAccountCLIE" value="CLIE"
						   onclick="localRicarica();" >				  
 					       Cliente
					    </div>
                     </div>					 
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=CollSele%> type="radio" 
                     	   name="ListaAccount" id="ListaAccountCOLL" value="COLL"
						   onclick="localRicarica();" >				  
 					       Collaboratore
					    </div>
                     </div>						 
				  </div>
<%
            'caricamento tabella 
            err.clear

				Set Rs = Server.CreateObject("ADODB.Recordset")

				MySql = "" 
				MySql = MySql & " Select A.*,B.DescStatoCredito as StatoCredito,C.Nominativo,DescTipoCredito,E.DescTipoAccount "
				MySql = MySql & " From AccountMovEco A, StatoCredito B, Account C , TipoCredito D , TipoAccount E"
				MySql = MySql & " Where A.IdStatoCredito = B.IdStatoCredito "
				MySql = MySql & " And   A.IdAccount = C.IdAccount"
				MySql = MySql & " And   A.IdTipoCredito = D.IdTipoCredito"
				MySql = MySql & " And   B.FlagStatoFinale = 0"
				MySql = MySql & " And   A.IdMovimento = 0"
				MySql = MySql & " And   A.IdStatoCredito = B.IdStatoCredito"
				MySql = MySql & " And   C.IdTipoAccount = E.IdTipoAccount"
				if IdTipoCredito<>"ALLL" then 
				   MySql = MySql & " And   A.IdTipoCredito = '" & IdtipoCredito & "'"
				end if 
				if IdTipoUtente<>"ALLL" then 
				   MySql = MySql & " And   C.IdTipoAccount = '" & IdTipoUtente & "'"
				end if 				
				MySql = MySql & Condizione & " order By IdAccountMovEco"

				Rs.CursorLocation = 3 
				Rs.Open MySql, ConnMsde

				DescLoaded=""
				NumCols = 3
				NumRec  = 0
				ShowNew    = true
				ShowUpdate = false
				MsgNoData  = ""
	%>

		<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->
	<%
				if MsgNoData="" then 
				if PageSize>0 then 
					Rs.PageSize = PageSize
					pageTotali = rs.PageCount
					NumRec=0
					if Cpag<=0 then 
						Cpag =1
					end if 
					if Cpag>PageTotali then 
						CPag=PageTotali
					end if  
					Rs.absolutepage=CPag
				end if
				NumRec=0
				InCert="0"
				Do While Not rs.EOF and (NumRec<PageSize or Pagesize<=0)
					Primo=Primo+1
					NumRec=NumRec+1
					Id=Rs("IdAccountMovEco")
					IdAccBackO=Rs("IdAccountGestore")
					DescStato= rs("StatoCredito")
					Note     = rs("DescStatoCredito")

			%> 
			
				<div class="row">
				  <%
				  xx=writeDiv(2,"Tipo Movimento"  ,Rs("DescTipoCredito")     ,"","")
				  xx=writeDiv(2,"Tipo Utente"     ,Rs("DescTipoAccount")     ,"","")				  
				  xx=writeDiv(4,"Nominativo"      ,Rs("Nominativo")          ,"","")				  
				  xx=writeDiv(4,"Descrizione"     ,Rs("DescMovEco")          ,"","")				  
				  %>
 
			   </div>
				<div class="row">
		         <%
				 xx=writeDiv(2,"Movimento del"  ,Stod(Rs("DataMovEco"))          ,"","")
				 xx=writeDiv(2,"Importo &euro;" ,InsertPoint(Rs("ImptMovEco"),2) ,"","")
				 xx=writeDiv(2,"Stato"          ,Rs("StatoCredito")              ,"","")
				 DescGestito=""
				 if Cdbl(IdAccBackO)>0 then 
				    DescGestito=LeggiCampo("select * from Account where IdAccount=" & IdAccBackO,"Nominativo")
				 end if 
				 
				 xx=writeDiv(2,"Gestito da"     ,DescGestito ,"","")
				 xx=writeDiv(3,"Annotazioni"    ,Note        ,"","")
				 %>
			      <div class="col-1">
                     <div class="form-group ">
				         <%xx=ShowLabel("Azioni")%>
						 <br>
							<%
							if Rs("IdStatoCredito")="LAVO" then
							   RiferimentoA=";#;;2;upda;Modifica;;localFun('CALL_GES'," & id & ");N"%>
							   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->   
							
							
				               <%if IsBackOffice() and Cdbl(IdAccBackO)=0 then
					                RiferimentoA=";#;;2;sele;Prendi in carico;;localFun('CALL_CAR'," & Id &");N"
					             %>
					            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
				               <%end if %>							
							<%end if %>					    
					 </div>
                  </div>
				</div>
				
						<%    
						rs.MoveNext
						if rs.EOF = false then 
						%>
						<!--#include virtual="/gscVirtual/include/rigaSepDiv.asp"-->
						<%
						end if 
					Loop
				end if 
				rs.close
           
%>
      
            <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
            </form>
        </div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

</body>

</html>
