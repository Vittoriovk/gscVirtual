<%
  NomePagina="ListaPagamentoAccountStorico.asp"
  titolo="Movimentazione Pagamenti Storico "
%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
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
     IdCliente     = cdbl("0" & getCurrentValueFor("IdCliente"))
     IdAccount     = cdbl("0" & getCurrentValueFor("IdAccount"))
     DescCliente   = getCurrentValueFor("DescCliente")
     OperTabella   = getCurrentValueFor("OperTabella")
     PaginaReturn  = getCurrentValueFor("PaginaReturn") 
	 IdTipoCredito = getCurrentValueFor("IdTipoCredito")
  else
     IdCliente     = cdbl("0" & getValueOfDic(Pagedic,"IdCliente"))
     IdAccount     = cdbl("0" & getValueOfDic(Pagedic,"IdAccount"))
     DescCliente   = getValueOfDic(Pagedic,"DescCliente")     
     OperTabella   = getValueOfDic(Pagedic,"OperTabella")
     PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
	 IdTipoCredito = Request("ListaModPag")
   end if 
   if Cdbl(IdCliente)=0 then 
      if isCliente() then 
         IdCliente = Session("LoginIdCliente")
		 IdAccount = Session("LoginIdAccount")
      end if 
   end if 
   
   IdCliente = cdbl(IdCliente)
   if Cdbl(IdCliente)=0 then 
      response.redirect RitornaA(PaginaReturn)
      response.end 
   end if 
  'inizio elaborazione pagina
   if DescCliente="" then 
      DescCliente=LeggiCampo("select * from Cliente Where IdCliente=" & IdCliente,"Denominazione")
   end if  
   if IdAccount=0 then 
      IdAccount  =LeggiCampo("select * from Cliente Where IdCliente=" & IdCliente,"IdAccount")   
   end if 

   DescPageOper=DescCliente
   
   xx=setValueOfDic(Pagedic,"IdCliente"       ,IdCliente)
   xx=setValueOfDic(Pagedic,"DescCliente"     ,DescCliente)
   xx=setValueOfDic(Pagedic,"IdAccount"       ,IdAccount)
   xx=setValueOfDic(Pagedic,"IdTipoCredito"   ,IdTipoCredito)
   xx=setValueOfDic(Pagedic,"OperTabella"     ,OperTabella)
   xx=setValueOfDic(Pagedic,"PaginaReturn"    ,PaginaReturn)
   xx=setCurrent(NomePagina,livelloPagina) 
   DescLoaded="0"  
   
  
  %>
<% 
  'xx=DumpDic(SessionDic,NomePagina)
%>
<div class="d-flex" id="wrapper">

    <%
	  if Iscliente() then 
	     Session("opzioneSidebar")="paga"
	  end if 
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
			RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
			<%end if %>
                <div class="col-11"><h3>Movimentazioni Pagamenti :</b> <%=DescPageOper%> </h3>
                </div>
            </div>

            <%
               FlagBorsellino = 1
               FlagFido       = 1
               FlagEstratto   = 1
               FlagAction = true 

                   flagchecked   = IdtipoCredito
				   IdtipoCredito = ""
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
					   if flagBorsellino=1 then
						  BorsSele= " checked "
						  IdtipoCredito= "BORS"
					   elseif flagFido=1 then
						  FidoSele= " checked "
						  IdtipoCredito= "FIDO"
					   else  
						  EstrSele= " checked "
						  IdtipoCredito= "ESTR"
					   end if 
				   end if 

			   %> 
			      <div class="row">
				     <div class="col-1">
					 </div>
			      <% if  FlagBorsellino=1 then %>  
                     
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=BorsSele%> type="radio" 
                     	   name="ListaModPag" id="ListaModPagBORS" value="BORS"
						   onclick="localRicarica();" >				  
 					       Borsellino
					    </div>
                     </div>		
			      <% end if %>
			      <% if  FlagFido=1 then %>  
                     
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=FidoSele%> type="radio" 
                     	   name="ListaModPag" id="ListaModPagFIDO" value="FIDO"
						   onclick="localRicarica();">				  
 					       Fido
					    </div>
                     </div>		
			      <% end if %>				  
			      <% if  FlagEstratto=1 then %>  
                     
					 <div class="col-2">
					    <div class="form-group font-weight-bold">
					   	   <input class="form-check-input" <%=EstrSele%> type="radio" 
                     	   name="ListaModPag" id="ListaModPagESTR" value="ESTR"
						   onclick="localRicarica();">				  
 					       Estratto
					    </div>
                     </div>		
			      <% end if %>					  
			      </div>
		
<%
            'caricamento tabella 
            err.clear

			Set Rs = Server.CreateObject("ADODB.Recordset")

			MySql = "" 
			MySql = MySql & " Select A.*,B.DescStatoCredito as StatoCredito "
			MySql = MySql & " From AccountMovEco A, StatoCredito B  "
			MySql = MySql & " Where A.IdAccount  = " & IdAccount
			MySql = MySql & " And   A.IdTipoCredito = '" & IdtipoCredito & "'"
			MySql = MySql & " And   A.IdStatoCredito = B.IdStatoCredito"
			MySql = MySql & " And   B.FlagStatoFinale = 1"
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
			<div class="table-responsive"><table class="table"><tbody>
			<thead>
				<tr>
					<th scope="col">Descrizione</th>
					<th scope="col" width="12%">Data Movimento</th>
					<th scope="col" width="12%">Importo &euro;</th>
					<th scope="col">Stato</th>
				</tr>
			</thead>
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
				DescStato=rs("StatoCredito")
				if rs("DescStatoCredito")<>"" then 
				   DescStato=DescStato & ":" & rs("DescStatoCredito")
				end if 

		%> 
				<tr scope="col"> 
					<td>
						<input class="form-control" type="text" readonly value="<%=Rs("DescMovEco")%>">
					</td>				
					<td>
						<input class="form-control text-center" type="text" readonly value="<%=Stod(Rs("DataMovEco"))%>">
					</td>				
					<td>
						<input class="form-control text-right" type="text" readonly value="<%=InsertPoint(Rs("ImptMovEco"),2)%>">
					</td>	
					<td>
						<input class="form-control" type="text" readonly value="<%=DescStato%>">
					</td>				

				</tr> 
					<%    
					rs.MoveNext
				Loop
			end if 
			rs.close
           
%>

</tbody></table></div> <!-- table responsive fluid -->
            

 
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
