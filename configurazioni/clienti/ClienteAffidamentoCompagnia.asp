<%
  NomePagina="ClienteAffidamentoCompagnia.asp"
  titolo="Affidamento cliente per compagnia"
  act_call_affi = CryptAction("CALL_AFFI")
  act_call_modi = CryptAction("CALL_MODI")
  act_call_dele = CryptAction("CALL_DELE")
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<!--#include virtual="/gscVirtual/common/functionAffidamento.asp"-->
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
<!--#include virtual="/gscVirtual/js/functionTable.js"-->
<script language="JavaScript">
function affiIniz(id)
{
    xx=ImpostaValoreDi("Oper","<%=act_call_affi%>");
	xx=ImpostaValoreDi("ItemToRemove",id);
    document.Fdati.submit();  
}

function modify(id)
{
    ImpostaValoreDi("Oper","<%=act_call_modi%>");
    xx=ImpostaValoreDi("ItemToRemove",id);
    document.Fdati.submit();
}
function remove(id)
{
    ImpostaValoreDi("Oper","<%=act_call_dele%>");
    xx=ImpostaValoreDi("ItemToRemove",id);
    document.Fdati.submit();
}
</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<!--#include virtual="/gscVirtual/modelli/FunctionAccount.asp"-->
  
 <!-- javascript locale -->


<%
  Set Rs = Server.CreateObject("ADODB.Recordset")

  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdCliente=0
  if FirstLoad then 
     IdCliente   = "0" & Session("swap_IdCliente")
     if Cdbl(IdCliente)=0 then 
        IdCliente = cdbl("0" & getValueOfDic(Pagedic,"IdCliente"))
     end if   
     PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
     if PaginaReturn="" then 
        PaginaReturn = Session("swap_PaginaReturn")
     end if 

  else
     IdAccount = "0" & getValueOfDic(Pagedic,"IdAccount")
     DenomClie =       getValueOfDic(Pagedic,"DenomClie")
     IdCliente = "0" & getValueOfDic(Pagedic,"IdCliente")
     PaginaReturn    = getValueOfDic(Pagedic,"PaginaReturn")
   end if 
   IdCliente = cdbl(IdCliente)
   IdAccount = cdbl(IdAccount)
   if Cdbl(IdAccount)=0 then 
      MySql = "Select * from Cliente Where IdCliente=" & IdCliente 
      Rs.CursorLocation = 3
      Rs.Open MySql, ConnMsde  
      IdAccount=Rs("IdAccount")
      DenomClie=Rs("Denominazione")
      Rs.close 
   end if 
   
   Oper = DecryptAction(Oper)
   
   DescPageOper="Aggiornamento"

   'registro i dati della pagina 
   xx=setValueOfDic(Pagedic,"IdCliente"        ,IdCliente)
   xx=setValueOfDic(Pagedic,"IdAccount"        ,IdAccount)
   xx=setValueOfDic(Pagedic,"DenomClie"        ,DenomClie)
   xx=setValueOfDic(Pagedic,"PaginaReturn"     ,PaginaReturn)
   if Oper="CALL_MODI" then 
      idRichiesta  = "0" & request("ItemToRemove")
      xx=RemoveSwap()
      Session("swap_IdCliente")       = IdCliente
      if Cdbl(idRichiesta)>0 then 
         Session("swap_IdRecMod")     = idRichiesta
      else 
         Session("swap_IdRecMod")     = 0
      end if 
      Session("swap_PaginaReturn")    = "configurazioni/Clienti/" & NomePagina

      response.redirect RitornaA("configurazioni/Clienti/ClienteAffidamentoCompagniaDettaglio.asp")
      response.end 
   end if
   if Oper="CALL_DELE" then 
      idRichiesta  = "0" & request("ItemToRemove")
      if Cdbl(IdRichiesta)>0 then 
         QDel = ""
         qDel = qDel & " Delete from AccountCreditoAffi"
         qDel = qDel & " Where idAccountCreditoAffi = " & IdRichiesta 
         ConnMsde.execute qDel 
      end if 
   end if
   if Oper="CALL_AFFI" then 
      idRichiesta  = "0" & request("ItemToRemove")
      xx=RemoveSwap()
      Session("swap_IdCliente")       = IdCliente
      if Cdbl(idRichiesta)>0 then 
         Session("swap_IdRecMod")     = idRichiesta
         Session("swap_PaginaReturn")    = "configurazioni/Clienti/" & NomePagina
         response.redirect RitornaA("configurazioni/Clienti/Affidamento/ClienteAffidamentoCompagniaIniziale.asp")
         response.end 
      end if   
   end if   
   
      
  
  xx=setCurrent(NomePagina,livelloPagina) 
  DescLoaded="0"  
  
  OperAmmesse=""
  if Session("LoginTipoUtente")=ucase("BackO") then
    OperAmmesse = "C"
  end if 
  
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
            <%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <div class="col-11"><h3>Gestione Cliente : Affidamenti per compagnia </b></h3>
                </div>
            </div>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

            <div class="row">
                <div class="col-4">
                    <div class="form-group ">
                        <%xx=ShowLabel("Utente")%>
                        <input type="text" readonly class="form-control" value="<%=DenomClie%>" >
                    </div>        
                </div>
                <div class="col-4">
                    <div class="form-group ">
                        <%xx=ShowLabel("Compagnia")

                        stdClass="class='form-control form-control-sm'"
                        IdCompagnia=Cdbl("0" & Request("IdCompagnia0"))
                        q="Select * from Compagnia "
                        q=Q & " order by DescCompagnia "
                        'Where 
                        response.write ListaDbChangeCompleta(q,"IdCompagnia0",IdCompagnia ,"IdCompagnia","DescCompagnia" ,1,"document.Fdati.submit()","","","","",stdClass)
                        %>
                    </div>        
                </div>
                
            </div>  
            <% 
            CompInCorso=""
            qSel = ""
            qSel = qSel & " Select D.* from "
            qSel = qSel & " AffidamentoRichiesta A, StatoServizio B, AffidamentoRichiestaComp C, Compagnia D"
            qSel = qSel & " Where A.IdAccountCliente = " & IdAccount
            qSel = qSel & " and   C.IdStatoAffidamento = B.IdStatoServizio"
            qSel = qSel & " and   B.FlagStatofinale = 0 "
            qSel = qSel & " and   A.IdAffidamentoRichiesta = C.IdAffidamentoRichiesta"
            qSel = qSel & " and   C.IdCompagnia = D.IdCompagnia"
            qSel = qSel & " order by D.DescCompagnia"
            Rs.CursorLocation = 3 
            Rs.Open qSel, ConnMsde
            Do While Not rs.EOF
               if CompInCorso<>"" then 
                  CompInCorso = CompInCorso & " , "
               end if 
               CompInCorso = CompInCorso & Rs("DescCompagnia")
               rs.MoveNext
            loop
            Rs.Close 

            if CompInCorso<>"" then 
            %>
            <div class="row">
                <div class="col-10">
                    <div class="form-group ">
                        <%xx=ShowLabel("ATTENZIONE : esistono affidamenti in corso per le seguenti compagnie.")%>
                        <input type="text" readonly class="form-control" value="<%=CompInCorso%>" >
                    </div>        
                </div>
            </div>
            <%
            end if 

            err.clear
            MySql = ""
            MySql = MySql & " select a.*,B.DescCompagnia"
            MySql = MySql & " From AccountCreditoAffi A, Compagnia B"
            MySql = MySql & " Where A.IdAccount = " & IdAccount
            MySql = MySql & " and A.IdCompagnia = B.IdCompagnia"

            if Cdbl(IdCompagnia)>0 then 
               MySql = MySql & " And A.IdCompagnia=" & IdCompagnia 
            end if 
            MySql = MySql & Condizione & " order By ValidoDal desc,DescCompagnia"

            Rs.CursorLocation = 3 
            Rs.Open MySql, ConnMsde

            DescLoaded=""
            NumCols = numC + 1
            NumRec  = 0
            ShowNew    = true
            ShowUpdate = false
            MsgNoData  = ""
            'elenco azioni 
            
            %>
<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->            
            <div class="table-responsive"><table class="table"><tbody>
            <thead>
                <tr>
                <th scope="col">Compagnia
                                        <%
                  if instr(OperAmmesse,"C")>0  then 
                          RiferimentoA="col-2;#;;2;inse;Inserisci;;modify(0);N"
                          %>
                        <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                  <%end if %>
                </th>
                <th scope="col" width="12%">Importo Affidato</th>
                <th scope="col" width="12%">x Polizza</th>
                <th scope="col" width="13%">Importo Impegnato</th>
                <th scope="col" width="13%">Importo Disponibile</th>
                <th scope="col" width="11%">Valido Dal</th>
                <th scope="col" width="11%">Valido Al</th>
                <% if Len(OperAmmesse)>0  then  %>
                <th scope="col">azioni</th>
                <% end if %>
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
            Primo=0
            Do While Not rs.EOF and (NumRec<PageSize or Pagesize<=0)
               Primo=Primo+1
               NumRec=NumRec+1
               Id=Rs("IdAccountCreditoAffi")
               DescLoaded=DescLoaded & Id & ";"
               Dim DizAff
               Set DizAff = CreateObject("Scripting.Dictionary")
   
			   esito=GetTotaliAffidamentoComp(DizAff,IdAccount,Rs("IdCompagnia"))
			   if Esito=true then 
			      ImptAddInit   = cdbl(Getdiz(DizAff,"ImptIniziale"))
				  ImptImpegnato = cdbl(Getdiz(DizAff,"TotaleImpegnato"))
			   else
			      ImptAddInit   = 0 
				  ImptImpegnato = 0
			   end if 
			   ShowGesAffIni = false 
			   if cdbl(ImptAddInit)>0 then 
			      ShowGesAffIni = true 
			   end if 

			   ImptDisponibile = cdbl(Rs("ImptComplessivo")) - Cdbl(ImptImpegnato)
            %>
			
            <tr scope="col">

                <td>
                    <input class="form-control" type="text" readonly value="<%=Rs("DescCompagnia")%>">
                </td>
                <td>
                    <input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("ImptComplessivo"),2)%>">
                </td>
                <td>
                    <input class="form-control" type="text" readonly value="<%=InsertPoint(Rs("ImptSingolaPolizza"),2)%>">
                </td>                
                <td>
                    <input class="form-control" type="text" readonly value="<%=InsertPoint(ImptImpegnato,2)%>">
                </td>
                <td>
                    <input class="form-control" type="text" readonly value="<%=InsertPoint(ImptDisponibile,2)%>">
                </td>
                <td>
                    <input class="form-control" type="text" readonly value="<%=StoD(Rs("ValidoDal"))%>">
                </td>                
                <td>
                    <input class="form-control" type="text" readonly value="<%=StoD(Rs("ValidoAl"))%>">
                </td>    
                <% if Len(OperAmmesse)>0  then  %>    
                <td>
                  <%
                  if instr(OperAmmesse,"C")>0  then 
                          RiferimentoA=";#;;2;upda;Aggiorna;;modify(" & Id & ");N"
                          %>
                        <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                          <%if Rs("ImptImpegnato") = 0 then 
                              RiferimentoA=";#;;2;dele;Cancella;;remove(" & Id & ");N"
                          %>
                        <!--#include virtual="/gscVirtual/include/Anchor.asp"-->                    
                        <%  end if 
                  end if %>                
				  <% if ShowGesAffIni then 
                          RiferimentoA=";#;;2;dett;Gestisci Affidamento Iniziale;;affiIniz(" & Id & ");N"
                          %>
                        <!--#include virtual="/gscVirtual/include/Anchor.asp"-->                    
				  <%end if %> 
				
                </td>
                <%end if %> 

            </tr>
			
			
			
			
			
			
            <%    
               rs.MoveNext
           Loop
end if 
rs.close

%>
            

            </tbody></table></div>
            <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
            </form>
<!--#include virtual="/gscVirtual/include/FormSoggetti.asp"-->
        </div> <!-- container fluid -->
    </div>
    <!-- /#page-content-wrapper -->

</div>
<!-- /#wrapper -->

<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

</body>

</html>
