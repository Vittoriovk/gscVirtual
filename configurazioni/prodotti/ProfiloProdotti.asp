<%
  NomePagina="ProfiloProdotti.asp"
  titolo="Prodotti disponibili"
  default_check_profile="Admin,Coll"
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

IdAzienda=1
if FirstLoad then 
   PaginaReturn   = getCurrentValueFor("PaginaReturn")
else
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
end if 

on error resume next 
if IsAdmin() then 
   IdAccount = 0
else
   IdAccount = Session("LoginIdAccount")
end if 

if CheckTimePageLoad()=false then
   oper=""
end if 


if oper=ucase("RemoveItem") then 
   Session("TimeStamp")=TimePage
   IdProfiloProdotto = Cdbl("0" & Request("ItemToRemove"))
   
   if Cdbl(IdProfiloProdotto)>0 then 
      ConnMsde.execute "Delete From ProfiloProdotto where IdProfiloProdotto=" & IdProfiloProdotto
      ConnMsde.execute "Delete From AccountProfiloProdotto where IdProfiloProdotto=" & IdProfiloProdotto
   end if 
end if 

if Oper="CALL_MOD" then 
   xx=RemoveSwap()
   Session("TimeStamp")=TimePage
   IdProfiloProdotto = cdbl("0" & Request("ItemToRemove"))
   Session("swap_IdProfiloProdotto") = IdProfiloProdotto
   Session("swap_PaginaReturn")  = "configurazioni/prodotti/" & nomePagina 
   response.redirect virtualPath & "configurazioni/prodotti/ProfiloProdottiModifica.asp"
   response.end 
End if 
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"PaginaReturn"  ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

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
                <div class="col-11"><h3>Profilo Prodotti</b></h3>
                </div>
            </div>
            <%
            AddRow=true
            dim CampoDb(10)
            CampoDB(1)="DescProfiloProdotto"    
            ElencoOption=";0;Descrizione;1"
            %>        
            <!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->

<%
'caricamento tabella 
if Condizione<>"" then 
    Condizione=" and " & Condizione
end if 
        
Set Rs = Server.CreateObject("ADODB.Recordset")

'associazioni presenti 
MySql = "" 
MySql = MySql & " Select * "
MySql = MySql & " From ProfiloProdotto"
MySql = MySql & " Where IdAccount=" & IdAccount
MySql = MySql & " And IdTipoProfilo = 'PROFILO'"
MySql = MySql & Condizione & " order By DescProfiloProdotto"

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
        <th scope="col">Profilo Prodotto
		<%RiferimentoA="col-2;#;;2;inse;Inserisci;;AttivaFunzione('CALL_MOD','0');N" %>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->		
		</th>
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
        IdP=Rs("IdProfiloProdotto")
		 
        %>
        <tr scope="col">
            <td>
                <input class="form-control" type="text" readonly value="<%=Rs("DescProfiloProdotto")%>">
            </td>
            <td>
				<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;AttivaFunzione('CALL_MOD','" & IdP & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->				
                <%RiferimentoA="col-2;#;;2;dele;Cancella;;RemoveItem('" & IdP & "');N"%>
                <!--#include virtual="/gscVirtual/include/Anchor.asp"-->  
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
