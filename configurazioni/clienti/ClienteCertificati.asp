<%
  NomePagina="ClienteCertificati.asp"
  titolo="Utenti per Azienda"
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

function localFun(Op,Id)
{
    xx=ImpostaValoreDi("DescLoaded","0");
    xx=ElaboraControlli();
    
     if (xx==false)
       return false;
    if (Op=="submit")
       ImpostaValoreDi("Oper","update");
    if (Op=="send")
       ImpostaValoreDi("Oper","update_send");
       
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

   if Oper="INS" then 
      Session("TimeStamp")=TimePage
      KK="0"
      IdCertificazione = Request("IdCertificazione" & KK)

      if Cdbl(IdCertificazione)>0 then 
         MyQ = "" 
         MyQ = MyQ & " Insert into AccountCertificazione ("
         MyQ = MyQ & " IdAccount,IdCertificazione"
         MyQ = MyQ & ") values ("            
         MyQ = MyQ & " " & IdAccount
         MyQ = MyQ & "," & IdCertificazione
         MyQ = MyQ & ")"
         ConnMsde.execute MyQ 
          If Err.Number <> 0 Then 
             MsgErrore = ErroreDb(Err.description)
          End If
      END if 
   End if 
   if Oper="DEL" then 
      Session("TimeStamp")=TimePage
      KK=Request("ItemToRemove")
      MyQ = "" 
      MyQ = MyQ & " delete from AccountCertificazione "
      MyQ = MyQ & " where IdAccount = " & IdAccount
      MyQ = MyQ & " and   IdCertificazione = " & KK
    
      ConnMsde.execute MyQ 
      If Err.Number <> 0 Then 
         MsgErrore = ErroreDb(Err.description)
      End If
   End if   
   DescPageOper=DescCliente

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
                <div class="col-11"><h3>elenco certificazione per :</b> <%=DescPageOper%> </h3>
                </div>
            </div>

<%
            'caricamento tabella 
            err.clear
            Set Rs = Server.CreateObject("ADODB.Recordset")

            MySql = "" 
            MySql = MySql & " Select a.*,b.DescBreveCertificazione "
            MySql = MySql & " From AccountCertificazione a, Certificazione B "
            MySql = MySql & " Where A.IdAccount  = " & IdAccount
            MySql = MySql & " And   A.IdCertificazione = B.IdCertificazione "
            MySql = MySql & Condizione & " order By B.DescBreveCertificazione"

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
                    <th scope="col">Certificazione</th>
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
                Id=Rs("IdCertificazione")
                InCert=Incert & "," & Id

        %> 
                <tr scope="col"> 
                    <td>
                        <input class="form-control" Id="IdCertificazione<%=Id%>" type="text" readonly value="<%=Rs("DescBreveCertificazione")%>">
                    </td>
                    <td>
                        <%RiferimentoA="col-2;#;;2;dele;Cancella;;SalvaSingoloEdAttiva('DEL'," & Id & ",true,'','','');N"%>
                        <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
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
            <%     IdRef="IdCertificazione" & Id     
            query = ""
            query = query & " Select * from Certificazione " 
            if InCert<>"" then 
               query = query & " Where IdCertificazione not in (" & inCert & ")" 
            end if 
            query = query & " order By DescBreveCertificazione"
            'response.write query 
            response.write ListaDbChangeCompleta (Query,IdRef,"0","IdCertificazione","DescBreveCertificazione",0,"","IdCertificazione","","","dati assenti","class='form-control form-control-sm'")
            
            xx="0" & LeggiCampo(query,"IdCertificazione")
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
