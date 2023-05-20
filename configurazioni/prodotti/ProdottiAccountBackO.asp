<%
  NomePagina="ProdottiAccountBackO.asp"
  titolo="Prodotti "
  default_check_profile="Admin"
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
NameLoaded=NameLoaded & "ValidoDal,DTO;ValidoAl;DTO"

IdAzienda=1
if FirstLoad then 
   PaginaReturn   = getCurrentValueFor("PaginaReturn")
   IdAccount      = getCurrentValueFor("IdAccount")
   DescAccount    = getCurrentValueFor("DescAccount")   
else
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
   IdAccount      = getValueOfDic(Pagedic,"IdAccount")
   DescAccount    = getValueOfDic(Pagedic,"DescAccount")
end if 

on error resume next 
if CheckTimePageLoad()=false then
   oper=""
end if 
if oper="CALL_ADD" then 
   Session("TimeStamp")=TimePage
   IdProdotto= Request("IdProdotto0")
   ArDati=split(IdProdotto,"_")
   idP = cdbl("0" & ArDati(0))
   idF = cdbl("0" & ArDati(1))

   if Cdbl(IdP)>0 and cdbl(idF)=0 then 
      q = ""
      q = q & " insert into AccountProdotto (IdAccount,IdProdotto,MailDocumentazione,IdAccountFornitore,ValidoAl) values ("
      q = q & " " & NumForDb(IdAccount)
      q = q & "," & NumForDb(IdP)
      q = q & ",''"
      q = q & "," & NumForDb(IdF)
      q = q & ",20991231"
      q = q & " )"
      connMsde.execute q 
   end if 
end if 
if oper="CALL_DEL" then 
   Session("TimeStamp")=TimePage
   IdProdotto= Request("ItemToRemove")
   ArDati=split(IdProdotto,"_")
   idP = cdbl("0" & ArDati(0))
   idF = cdbl("0" & ArDati(1))
   if Cdbl(IdP)>0 and cdbl(idF)=0 then    
      q = ""
      q = q & " delete from AccountProdotto"
      q = q & " where IdAccount = " & NumForDb(IdAccount)
      q = q & " and IdProdotto = " & NumForDb(IdP)
      q = q & " and IdAccountFornitore = " & NumForDb(IdF)
      connMsde.execute q 
   end if 
   
end if 

  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"PaginaReturn"  ,PaginaReturn)
  xx=setValueOfDic(Pagedic,"IdAccount"     ,IdAccount)
  xx=setValueOfDic(Pagedic,"DescAccount"   ,DescAccount)  
  xx=setCurrent(NomePagina,livelloPagina) 

  'response.write IdAccount
  DescAccount   = leggiNominativoAccount(IdAccount)
  IdTipoAccount = LeggiCampo("select * from Account Where IdAccount=" & IdAccount,"IdTipoAccount")
  
  'xx=DumpDic(SessionDic,NomePagina)
%>
<div class="d-flex" id="wrapper">

    <%
      TitoloNavigazione="Configurazioni"
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
                <div class="col-11"><h3>Associazione Prodotti</b></h3>
                </div>
            </div>
            <div class="row">
               <div class="col-1">
               </div>
               <div class="col-4 form-group ">
                  <%
			     descTipoAccount="Utente Back Office"
				  xx=ShowLabel(descTipoAccount)%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescAccount%>" >
               </div>    
            </div>
            <%
            AddRow=true
            dim CampoDb(10)
            CampoDB(1)="DescProdotto"    
            ElencoOption=";0;Prodotto;1"
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
MySql = MySql & " Select a.*,B.FlagPrezzoFisso,B.IdAnagServizio,B.DescProdotto"
MySql = MySql & " From AccountProdotto a, Prodotto B "
MySql = MySql & " Where A.IdAccount = " & IdAccount
MySql = MySql & " and A.IdProdotto = B.IdProdotto "
MySql = MySql & Condizione & " order By B.DescProdotto"

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

<div class="row">
     <div class="col-12 bg-primary text-white font-weight-bold">
     Prodotti gi&agrave; associati
     </div>
</div>

<div class="table-responsive"><table class="table"><tbody>
<thead>
    <tr>
        <th scope="col">Prodotto</th>
        <th scope="col">Azioni</th>
    </tr>
</thead>

<%
PageSize=0
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
        IdP=Rs("IdProdotto") & "_0"
        DescLoaded=DescLoaded & Id & ";"
        %>
        
        <tr scope="col">
            <td>
                <input class="form-control" type="text" readonly value="<%=Rs("DescProdotto")%>">
            </td>
            <td>
                <%RiferimentoA="col-2;#;;2;dele;Cancella;;AttivaFunzione('CALL_DEL','" & IdP & "');N"%>
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
</tbody></table></div> 

<%
'lista dei prodotti associabili
   IdAnagServizio=Request("IdAnagServizio0")
   if IdAnagServizio="-1" then
      IdAnagServizio=""
   end if 

   q = ""
   q = q & " select B.IdProdotto,B.DescProdotto"
   q = q & " From ProdottoAttivo A , Prodotto B "
   q = q & " Where A.IdProdotto = B.IdProdotto "
   q = q & " and   a.ValidoDal <= " & Dtos()
   q = q & " and   a.ValidoAl  >= " & Dtos()
   if IdAnagServizio<>"" then 
      q = q & " and b.IdAnagServizio = '" & apici(IdAnagServizio) & "'"
   end if    
   'escludo quelli gia associati 
   qe = ""
   qe = qe & " select 'X' from AccountProdotto tt "
   qe = qe & " where tt.IdAccount=" & IdAccount 
   qe = qe & " and tt.idProdotto = A.IdProdotto "
      
   q = q & " and not Exists (" & qe &  ")"

   tt = leggiCampo(q,"IdProdotto")

   if tt="" and IdAnagServizio="" then 
%>
<div class="row">
     <div class="col-12 bg-primary text-white font-weight-bold">
     nessun altro prodotto associabile 
     </div>
</div>
<%else%>

  
<div class="row">
     <div class="col-12 bg-primary text-white font-weight-bold">
     Prodotti da associare 
     </div>
</div>

<div class="row">
     <div class="col-2 font-weight-bold">filtra per servizio</div>
	 <div class="col-4 ">
    <% 


	    stdClass="class='form-control form-control-sm'"
	    q1 = ""
        q1 = q1 & " Select * from AnagServizio "
        q1 = q1 & " order by DescAnagServizio  "
        response.write ListaDbChangeCompleta(q1,"IdAnagServizio0",IdAnagServizio ,"IdAnagServizio","DescAnagServizio" ,1,"Sottometti();","","","","",stdClass)
    %>
</div></div>
<br>
<div class="table-responsive"><table class="table"><tbody>
<thead>
    <tr>
        <th scope="col">Prodotto</th>
        <th scope="col">Azioni</th>
    </tr>
</thead>
   
   <%


   if tt<>"" then 
        %>
        
        <tr scope="col">
            <td>
            <%
            stdClass="class='form-control form-control-sm'" 
            response.write ListaDbChangeCompleta(q,"IdProdotto0",IdProdotto ,"IdProdotto","descProdotto" ,0,"","","","","",stdClass)
     
           %>
            </td>
            <td>
               
                <%RiferimentoA="col-2;#;;2;inse;Aggiungi;;AttivaFunzione('CALL_ADD','0');N"%>
                <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
            <td>
        </td>
        </tr>
    <%end if %>
</tbody></table></div> <!-- table responsive fluid -->            
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
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

</body>

</html>
