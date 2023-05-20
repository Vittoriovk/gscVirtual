<%
  NomePagina="Prodotto.asp"
  titolo="Anagrafica Prodotto"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<%
  livelloPagina="00"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include virtual="/gscVirtual/include/head.asp"-->
<!-- Custom styles for this  -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">
</head>

<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->


<%

if FirstLoad then 
   IdCompagnia   = getCurrentValueFor("IdCompagnia") 
   PaginaReturn  = getCurrentValueFor("PaginaReturn") 
else
   IdCompagnia   = getValueOfDic(Pagedic,"IdCompagnia")
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
end if 
IdCompagnia = Cdbl("0" & IdCompagnia)
if Cdbl("0" & IdCompagnia)=0 then 
   response.redirect PaginaReturn
   response.end 
end if 

on error resume next 
if Oper="CALL_INS" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
    Session("swap_IdCompagnia")   = IdCompagnia
    Session("swap_OperTabella")   = Oper
    Session("swap_PaginaReturn")  = "configurazioni/Prodotti/Prodotto.asp"
    response.redirect virtualPath & "configurazioni/Prodotti/ProdottoAggiungi.asp"
    response.end 
End if 
if Oper="CALL_UPD" or Oper="CALL_DEL" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
    KK=Request("ItemToRemove")
    Session("swap_IdProdotto")   = KK
    Session("swap_OperTabella")   = Oper
    Session("swap_PaginaReturn")  = "configurazioni/Prodotti/Prodotto.asp"
    response.redirect virtualPath & "configurazioni/Prodotti/ProdottoDettaglio.asp"
    response.end 
End if 
if Oper="CALL_COND" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
    KK=Request("ItemToRemove")
    if Cdbl("0" & KK ) > 0 then 
        Session("swap_IdTabella")       = "PRODOTTO_COND"
        Session("swap_IdProdotto")      = KK
        Session("swap_IdTabellaKeyInt") = KK
        Session("swap_OperTabella")   = Oper
        Session("swap_PaginaReturn")  = "configurazioni/Prodotti/Prodotto.asp"
        response.redirect virtualPath & "configurazioni/Documenti/DocumentiElenco.asp"
        response.end 
    end if 
End if 
if Oper="CALL_TECN" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
    KK=Request("ItemToRemove")
    if Cdbl("0" & KK ) > 0 then 
        Session("swap_IdProdotto")      = KK
        Session("swap_IdTabellaKeyInt") = KK
        Session("swap_OperTabella")   = Oper
        Session("swap_PaginaReturn")  = "configurazioni/Prodotti/Prodotto.asp"
        response.redirect virtualPath & "configurazioni/prodotti/ProdottoListaDati.asp"
        response.end 
    end if 
End if 
if Oper="CALL_OPZI" then 
    xx=RemoveSwap()
    Session("TimeStamp")=TimePage
    KK=Request("ItemToRemove")
    if Cdbl("0" & KK ) > 0 then 
        Session("swap_IdProdotto")      = KK
        Session("swap_IdTabellaKeyInt") = KK
        Session("swap_OperTabella")   = Oper
        Session("swap_PaginaReturn")  = "configurazioni/Prodotti/Prodotto.asp"
        response.redirect virtualPath & "configurazioni/prodotti/ProdottoListaOpzioni.asp"
        response.end 
    end if 
End if 

  'registro i dati della pagina 
xx=setValueOfDic(Pagedic,"PaginaReturn"   ,PaginaReturn)
xx=setValueOfDic(Pagedic,"IdCompagnia"    ,IdCompagnia)
DescCompagnia = LeggiCampo("select * from Compagnia where IdCompagnia=" & IdCompagnia,"DescCompagnia")
xx=setCurrent(NomePagina,livelloPagina) 
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
            <%RiferimentoA="col-1 text-center;" & VirtualPath & paginaReturn & ";;2;prev;Indietro;;;"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <div class="col-11"><h3>Elenco Prodotti Per compagnia</b></h3>
                </div>
            </div>
	        <div class="row">
	           <div class="col-1">
	           </div>
               <div class="col-4 form-group ">
		          <%xx=ShowLabel("Compagnia")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescCompagnia%>" >
               </div>
			</div>				

            <%
            AddRow=true
            dim CampoDb(10)
            CampoDB(1)="DescProdotto"    
            CampoDB(2)="DescRamo"    
            ElencoOption=";0;Descrizione;1;Ramo;2"
            %>        
            <!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->

<%
'caricamento tabella 
if Condizione<>"" then 
    Condizione=" and " & Condizione
end if 
        
Set Rs = Server.CreateObject("ADODB.Recordset")

MySql = "" 
MySql = MySql & " Select a.*,B.DescRamo "
MySql = MySql & " From Prodotto a,Ramo B "
MySql = MySql & " Where a.IdProdotto > 0 "
MySql = MySql & " and   a.IdRamo = B.idRamo "
MySql = MySql & " and   a.IdCompagnia = " & idCompagnia
MySql = MySql & Condizione & " order By DescProdotto"

'response.write MySql 

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

DescLoaded=""
NumCols = numC + 1
NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
%>

<!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

<div class="table-responsive"><table class="table"><tbody>
<thead>
    <tr>
        <th scope="col">Prodotto
        <a href="#" title="Inserisci" onclick="AttivaFunzione('CALL_INS','0');">
        <i class="fa fa-2x fa-plus-square"></i></a>
        </th>
        <th scope="col">Ramo - Servizio Rif.</th>
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
    Primo=0
    Do While Not rs.EOF and (NumRec<PageSize or Pagesize<=0)
        Primo=Primo+1
        NumRec=NumRec+1
        Id=Rs("IdProdotto")
        err.clear 
        %>
        
        <tr scope="col">
            <td>
            <input class="form-control" type="text" readonly value="<%=Rs("DescProdotto")%>">
            </td>
            <td>
            <%
            idAnagServ=Rs("idAnagServizio")
            DescServ=Leggicampo("select * from AnagServizio Where IdAnagServizio='" & apici(IdAnagServ) & "'","DescAnagServizio") 
            %>
            <input class="form-control" type="text" readonly value="<%=Rs("DescRamo") & " - " & descServ%>">
            </td>            
            
            <td>
                <%if rs("flagModificabile")=1 then
                     IdAnagServizio=rs("IdAnagServizio")
                     q = ""
                     q = q & " select distinct IdTipoOpzione From Opzione "
                     q = q & " where IdAnagServizio='" & IdAnagServizio & "'" 
                     q = q & " And IdTipoOpzione='"

                     ContaOpziTecn=LeggiCampo(q & "TECN'","IdTipoOpzione")
                     ContaOpziOpzi=LeggiCampo(q & "OPZI'","IdTipoOpzione")
                %>
            
                   <%RiferimentoA="col-2;#;;2;upda;Aggiorna;;AttivaFunzione('CALL_UPD','" & Id & "');N"%>
                   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                   <%RiferimentoA="col-2;#;;2;dele;Cancella;;AttivaFunzione('CALL_DEL','" & Id & "');N"%>
                   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                   <%RiferimentoA="col-2;#;;2;hand;Condizioni di contratto;;AttivaFunzione('CALL_COND','" & Id & "');N"%>
                   <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            

                      <%RiferimentoA="col-2;#;;2;manu;Dati Aggiuntivi;;AttivaFunzione( 'CALL_TECN','" & Id & "');N"%>
                      <!--#include virtual="/gscVirtual/include/Anchor.asp"-->                 

                   <%if contaOpziOpzi<>"" then %>
                      <%RiferimentoA="col-2;#;;2;penn;Opzioni ;;AttivaFunzione( 'CALL_OPZI','" & Id & "');N"%>
                      <!--#include virtual="/gscVirtual/include/Anchor.asp"-->                 
                   <%end if%>
                
                <%end if %>
            <td>
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
            <!--#include virtual="/gscVirtual/include/paginazione.asp"-->            
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
