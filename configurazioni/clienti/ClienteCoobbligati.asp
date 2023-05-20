<%
  NomePagina="ClienteCoobbligati.asp"
  titolo="Coobbligati per Azienda"
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

function localFun(op,id)
{
    xx=ImpostaValoreDi("ItemToRemove",id);
    ImpostaValoreDi("Oper",op);
    document.Fdati.submit();

}
function attivaForm()
{
	xx=$('#confirmModalCoob').modal('toggle');
}

function localIns()
{
	xx=ImpostaValoreDi("Oper","CALL_INS");
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
  else
     IdCliente     = cdbl("0" & getValueOfDic(Pagedic,"IdCliente"))
     IdAccount     = cdbl("0" & getValueOfDic(Pagedic,"IdAccount"))
     DescCliente   = getValueOfDic(Pagedic,"DescCliente")     
     OperTabella   = getValueOfDic(Pagedic,"OperTabella")
     PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   end if 
   if Cdbl(IdCliente)=0 then 
      if isCliente() then 
         IdCliente = Session("LoginIdCliente")
		 IdAccount = Session("LoginIdAccount")
      end if 
   end if 
   
   IdCliente = cdbl(IdCliente)
   if Cdbl(IdCliente)=0 then 
      IdCliente  =LeggiCampo("select * from Cliente Where IdAccount=" & IdAccount,"IdCliente")   
   end if 
   
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

   if Oper="CALL_DEL" then 
      IdCoobbligato = cdbl("0" & Request("ItemToRemove"))
	  if Cdbl(IdCoobbligato)>0 then 
         MyQ = MyQ & " Delete from AccountCoobbligato "
         MyQ = MyQ & " Where IdAccountCoobbligato = " & IdCoobbligato 
         MyQ = MyQ & " and   IdAccount=" & IdAccount
         ConnMsde.execute MyQ      
      end if 
   end if 
   
   if Oper="CALL_DETT" then 
      xx=RemoveSwap()
      Session("TimeStamp")=TimePage
      KK=Request("ItemToRemove")
      if Cdbl("0" & KK ) > 0 then 
         Session("swap_IdListaDocumento")= KK
         Session("swap_OperTabella")     = Oper
         Session("swap_TipoRife") = "COOB"
         Session("swap_IdRife")   = KK
         Session("swap_PaginaReturn")    = "configurazioni/clienti/ClienteCoobbligati.asp"
         response.redirect virtualPath   & "configurazioni/clienti/AffidamentoAtiCoob.asp"
         response.end 
      end if 
   End if 
   
   if Oper="CALL_INS" then 
      Session("TimeStamp")=TimePage
      IdCoobbligato = 0
	  tipoId = Request("gruppo1")
      xx=RemoveSwap()
	  if Cdbl(IdAccount)>0 then 
         Session("swap_IdAccCliente")  = IdAccount
         Session("swap_IdCoobbligato") = IdCoobbligato
         Session("swap_PaginaReturn")  = "configurazioni/Clienti/" & NomePagina
         Session("swap_IdPersCliente") = tipoID
         response.redirect RitornaA("configurazioni/Clienti/ClienteCoobbligatiMod.asp")
         response.end 
      end if 
   End if    
   if Oper="CALL_GES" then 
      Session("TimeStamp")=TimePage
      IdCoobbligato = cdbl("0" & Request("ItemToRemove"))
	  if cdbl(IdCoobbligato)>0 then 
         xx=RemoveSwap()
         Session("swap_IdAccCliente")  = IdAccount
	     Session("swap_IdCoobbligato") = IdCoobbligato
		 Session("swap_OperAmmesse")   = "U"
         Session("swap_PaginaReturn")  = "configurazioni/Clienti/" & NomePagina
  
         response.redirect RitornaA("configurazioni/Clienti/ClienteCoobbligatiMod.asp")
         response.end 
	  end if 
   End if 
   DescPageOper=DescCliente
   if Iscliente() then 
      Session("opzioneSidebar")="coob"
	  PaginaReturn=""
   end if 

  
   xx=setValueOfDic(Pagedic,"IdCliente"       ,IdCliente)
   xx=setValueOfDic(Pagedic,"DescCliente"     ,DescCliente)
   xx=setValueOfDic(Pagedic,"IdAccount"       ,IdAccount)
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
	     Session("opzioneSidebar")="coob"
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
			if PaginaReturn<>"" and instr(PaginaReturn,NomePagina)=0 then 
			   RiferimentoA="col-1  text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->    
            <%else%>
                <div class="col-1"></div>
			
			<%end if %>
                <div class="col-11"><h3>Elenco coobbligati</h3>
                </div>
            </div>
			<%if Iscliente()=false or 1=1 then %>
			<div class="row">
			   <div class="col-1"></div>
			   <div class="col-4">
                  <div class="form-group ">
				     <%xx=ShowLabel("Cliente")%>
					 <input type="text" readonly class="form-control" value="<%=DescCliente%>" >
                  </div>		
			   </div>
			</div>			
            <%end if %>

<%
            'caricamento tabella 
            err.clear
            Set Rs = Server.CreateObject("ADODB.Recordset")

            MySql = "" 
            MySql = MySql & " Select * "
            MySql = MySql & " From AccountCoobbligato "
            MySql = MySql & " Where IdAccount  = " & IdAccount
            MySql = MySql & Condizione & " order By RagSoc"

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


            <div class="table-responsive"><table class="table"><tbody>
            <thead>
                <tr>
                    <th scope="col">Coobbligato&nbsp;&nbsp;
					      <%
						  RiferimentoA="col-2;#;;2;inse;Inserisci;;attivaForm();N"
						  %>
						<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
					</th>
					<th scope="col">CF</th>
					<th scope="col">PI</th>
					<th scope="col">Stato</th>
					<th scope="col" width="12%" >Valido dal</th>
					<th scope="col" width="12%">Valido al</th>
                    <th scope="col">Azioni</th>
                </tr>
            </thead>

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
                Id=Rs("IdAccountCoobbligato")

        %> 
                <tr scope="col"> 
                    <td>
                        <input class="form-control" type="text" readonly value="<%=Rs("Ragsoc")%>">
                    </td>				
                    <td>
                        <input class="form-control" type="text" readonly value="<%=Rs("CF")%>">
                    </td>				
                    <td>
                        <input class="form-control" type="text" readonly value="<%=Rs("PI")%>">
                    </td>

                    <td><%
					      IdStatoValidazione = rs("IdStatoValidazione")
						  FlagStatoFinale = cdbl("0" & getInfoStatoServizio(IdStatoValidazione,"FlagStatoFinale"))
					      if Rs("FlagValidato")=0 or Rs("ValidoAl") < Cdbl(Oggi) then 
					         Stato = "Non Validato"
							 
							 if IdStatoValidazione<>"" then 
							    if IdStatoValidazione = "RICH" then 
								   DescStatoValidazione="Richiesta Validazione"
								else
							       DescStatoValidazione=LeggiCampoTabellaText("StatoServizio",Rs("IdStatoValidazione"))
							    end if 
							    Stato = trim(DescStatoValidazione & " " &  Rs("NoteValidazione"))
							 end if 
					      else
						     Stato = "Validato"
						  end if 
						  %>
                        <input class="form-control" type="text" readonly value="<%=Stato%>">
                    </td>				
                    <td>
                        <input class="form-control" type="text" readonly value="<%=Stod(Rs("ValidoDal"))%>">
                    </td>
                    <td>
                        <input class="form-control" type="text" readonly value="<%=Stod(Rs("ValidoAl"))%>">
                    </td>
                    <td>
                        <%RiferimentoA="col-2;#;;2;upda;Modifica;;localFun('CALL_GES'," & id & ");N"%>
                        <!--#include virtual="/gscVirtual/include/Anchor.asp"-->   
                        <%RiferimentoA="col-2;#;;2;hand;Affidamento;;AttivaFunzione('CALL_DETT','" & Id & "');N"%>
                        <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                        <%
						Oggi = Dtos()
						if FlagStatoFinale=0 or IdStatoValidazione<>"AFFI" or (IdStatoValidazione="AFFI" and Rs("ValidoAl") < Cdbl(Oggi)) then
				           RiferimentoA="col-2;#;;2;dele;Cancella;;localFun('CALL_DEL'," & id & ");N"%>
                           <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
						<%end if %>
                    </td>
                </tr> 
                    <%    
                    rs.MoveNext
                Loop
            end if 
            rs.close

%>
<%if ShowNew then 
    Id=0
%>
    <tr> 
        <td>
            <%     IdRef="IdCoobbligato" & Id     
            query = ""
            query = query & " Select * from Coobbligato " 
            if InCert<>"" then 
               query = query & " Where IdCoobbligato not in (" & inCert & ")" 
            end if 
            query = query & " order By DescBreveCoobbligato"
            'response.write query 
            response.write ListaDbChangeCompleta (Query,IdRef,"0","IdCoobbligato","DescBreveCoobbligato",0,"","IdCoobbligato","","","dati assenti","class='form-control form-control-sm'")
            
            xx="0" & LeggiCampo(query,"IdCoobbligato")
            %>
        </td>
        <td align="left">
            <%if Cdbl(xx)>0 then %>
            <%RiferimentoA="col-2;#;;2;insert;Inserisci;;SaveWithOper('INS')"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
            <%end if %>
        </td>
    </tr>            
       
<%end if%>
</tbody></table></div> <!-- table responsive fluid -->
            
            <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
			
			
<div class="modal fade" id="confirmModalCoob"  aria-hidden="true" role="dialog">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">

        <h2>Seleziona Tipo Coobbligato </h2> 
        <button type="button" class="close" data-dismiss="modal">
          <span aria-hidden="true">×</span><span class="sr-only">Chiudi</span>
        </button>
      </div>

      <div class="modal-body"> 
		<div>
		  <div class="form-check">
			<input name="gruppo1" type="radio" id="radio1"  value="PEFI" checked>
			<label for="radio1">Persona fisica</label>
		  </div>
		  <div class="form-check">
			<input name="gruppo1" type="radio" id="radio2" value="PEGI" >
			<label for="radio2">Persona giuridica</label>
		  </div>
		</div>		  
      </div> 

      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>
        <button type="button" class="btn btn-primary" onclick="localIns();";>Seleziona</button>
      </div>
    </div>
  </div>
</div>
			
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
