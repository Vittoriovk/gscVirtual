<%
  NomePagina="ProdottiFornitoreDocAff.asp"
  titolo="Menu Supervisor - Dashboard"
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
<script>
function updMail() {
    xx=ImpostaValoreDi("NameLoaded","mail,EM");
    xx=ImpostaValoreDi("DescLoaded","0");
    xx=ElaboraControlli();
     if (xx==false) {
       return false;
    } 
    
    ImpostaValoreDi("Oper","UPDMAIL");
    document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<%
NameLoaded=NameLoaded & "IdDocumento,LI"

Set Rs = Server.CreateObject("ADODB.Recordset")

IdProdotto   = 0
IdFornitore  = 0
IdAccount    = 0
DescFornitore= ""
DescProdotto = ""
TipoDoc      = ""
if FirstLoad then 
   IdProdotto   = "0" & Session("swap_IdProdotto")
   if Cdbl(IdProdotto)=0 then 
      IdProdotto = cdbl("0" & getValueOfDic(Pagedic,"IdProdotto"))
   end if 
   IdFornitore   = "0" & Session("swap_IdFornitore")
   if Cdbl(IdFornitore)=0 then 
      IdFornitore = cdbl("0" & getValueOfDic(Pagedic,"IdFornitore"))
   end if 
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
   if PaginaReturn="" then 
      PaginaReturn = Session("swap_PaginaReturn")
   end if 
   IdFornitore = cdbl("0" & IdFornitore)
   if cdbl(IdFornitore)>0 then 
      Rs.CursorLocation = 3 
      Rs.Open "Select * from Fornitore where IdFornitore=" & IdFornitore, ConnMsde   
      IdAccount     = Rs("IdAccount")
      DescFornitore = Rs("DescFornitore")
      Rs.close 
   end if      
   if cdbl(IdProdotto)>0 then 
      Rs.CursorLocation = 3 
      Rs.Open "Select * from Prodotto where IdProdotto=" & IdProdotto, ConnMsde   
      DescProdotto = Rs("DescProdotto")
      Rs.close 
   end if  
   TipoDoc = getCurrentValueFor("TipoDoc")   
else
   IdProdotto     = "0" & getValueOfDic(Pagedic,"IdProdotto")
   IdAccount      = "0" & getValueOfDic(Pagedic,"IdAccount")
   IdFornitore    = "0" & getValueOfDic(Pagedic,"IdFornitore")
   DescProdotto   = getValueOfDic(Pagedic,"DescProdotto")
   DescFornitore  = getValueOfDic(Pagedic,"DescFornitore")
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
   TipoDoc        = getValueOfDic(Pagedic,"TipoDoc")
end if 
if cdbl(IdProdotto)=0 or Cdbl(IdFornitore)=0  then 
   response.redirect virtualPath & PaginaReturn
   response.end 
end if 

IdAccount =cdbl(IdAccount)
IdProdotto=cdbl(IdProdotto)
on error resume next
 
FlagUpdRiferimento=false 
if Oper="CALL_FROM" then 
   IdLista=Request("ListaDoc")

   MyQ = "" 
   
   MyQ = MyQ & " delete from AccountProdottoDocAff "
   MyQ = MyQ & " where IdAccount=" & idAccount
   MyQ = MyQ & " and IdProdotto=" & IdProdotto
   ConnMsde.execute MyQ
   
   MyQ = "" 
   MyQ = MyQ & " Insert into AccountProdottoDocAff ("
   MyQ = MyQ & " IdAccount,IdProdotto,IdDocumento,FlagObbligatorio,FlagDataScadenza,DITT,PEFI,PEGI,PEGC,TipoDoc)"
   MyQ = MyQ & " select " & idAccount & " as IdAccount, " & IdProdotto & " as IdProdotto,IdDocumento"
   MyQ = MyQ & ",FlagObbligatorio,FlagDataScadenza,DITT,PEFI,PEGI,PEGC,'" & TipoDoc & "' as TipoDoc"   
   MyQ = MyQ & " from ServizioDocumento Where IdAnagServizio='LISTA' and IdTipoUtenza='" & IdLista & "'"
   MyQ = MyQ & " and IdDocumento not in (select IdDocumento from AccountProdottoDocAff Where IdAccount=" & IdAccount & " and IdProdotto=" & IdProdotto & " ) "
   ConnMsde.execute MyQ
   'response.write MyQ
end if 


if Oper="INS" then 
    Session("TimeStamp")=TimePage
    KK="0"
    IdDocumento = Request("IdDocumento" & KK)
    obbl        = Request("checkObb" & KK)
    scad        = Request("checkSca" & KK)
    if obbl="S" then 
       Obbligatorio = 1 
    else
       Obbligatorio = 0
    end if 
    if scad="S" then 
       scadenza = 1 
    else
       scadenza = 0
    end if     
    if Cdbl(IdDocumento)>0 and cdbl(IdProdotto)>0 then 
        MyQ = "" 
        MyQ = MyQ & " Insert into AccountProdottoDocAff ("
        MyQ = MyQ & " IdAccount,IdProdotto,IdDocumento,FlagObbligatorio,FlagDataScadenza,DITT,PEGI,PEFI,PEGC,TipoDoc"
        MyQ = MyQ & ") values ("        
        MyQ = MyQ & "  " & IdAccount
        MyQ = MyQ & " ," & IdProdotto          
        MyQ = MyQ & " ," & IdDocumento 
        MyQ = MyQ & " ," & Obbligatorio 
        MyQ = MyQ & " ," & scadenza 
        MyQ = MyQ & ",'" & Request("checkDITT" & KK) & "'"
        MyQ = MyQ & ",'" & Request("checkPEGI" & KK) & "'"
        MyQ = MyQ & ",'" & Request("checkPEFI" & KK) & "'"        
        MyQ = MyQ & ",'" & Request("checkPEGC" & KK) & "'"
        MyQ = MyQ & ",'" & TipoDoc & "'"
        MyQ = MyQ & ")"

        ConnMsde.execute MyQ 
        If Err.Number <> 0 Then 
            MsgErrore = ErroreDb(Err.description)
        else
            DescIn=""
        End If
    END if 
End if 
if Oper="UPD" then 
    Session("TimeStamp")=TimePage
    KK=Request("ItemToRemove")
    IdDocumento = KK
    obbl        = Request("checkObb" & KK)
    scad        = Request("checkSca" & KK)
    if obbl="S" then 
       Obbligatorio = 1 
    else
       Obbligatorio = 0
    end if 
    if scad="S" then 
       scadenza = 1 
    else
       scadenza = 0
    end if 

    if Cdbl(IdDocumento)>0 and cdbl(IdProdotto)>0 then 
        MyQ = "" 
        MyQ = MyQ & " update AccountProdottoDocAff "
        MyQ = MyQ & " set "
        MyQ = MyQ & " FlagObbligatorio=" & Obbligatorio
        MyQ = MyQ & ",FlagDataScadenza=" & scadenza  
        MyQ = MyQ & ",DITT='"            & Request("checkDITT" & KK) & "'"
        MyQ = MyQ & ",PEGI='"            & Request("checkPEGI" & KK) & "'"
        MyQ = MyQ & ",PEFI='"            & Request("checkPEFI" & KK) & "'"
        MyQ = MyQ & ",PEGC='"            & Request("checkPEGC" & KK) & "'"        
        MyQ = MyQ & " where IdProdotto = " & IdProdotto
        MyQ = MyQ & " and   IdAccount  = " & IdAccount 
        MyQ = MyQ & " and   IdDocumento = " & IdDocumento 
        MyQ = MyQ & " and   TipoDoc = '" & TipoDoc & "'" 

        ConnMsde.execute MyQ 
        If Err.Number <> 0 Then 
            MsgErrore = ErroreDb(Err.description)
        else
            DescIn=""
        End If
    END if 
End if 
if Oper="DEL" then 
    Session("TimeStamp")=TimePage
    KK=Request("ItemToRemove")
    MyQ = "" 
    MyQ = MyQ & " delete from AccountProdottoDocAff "
    MyQ = MyQ & " where IdDocumento = " & KK
    MyQ = MyQ & " and   IdProdotto = " & IdProdotto
    MyQ = MyQ & " and   IdAccount = "  & IdAccount
    MyQ = MyQ & " and   TipoDoc = '" & TipoDoc & "'" 
    
    ConnMsde.execute MyQ 
    If Err.Number <> 0 Then 
        MsgErrore = ErroreDb(Err.description)
    End If
    DescIn=""
End if
if Oper="UPDMAIL" then 
    Session("TimeStamp")=TimePage
    KK=trim(Request("mail0"))
    MyQ = "" 
    MyQ = MyQ & " update AccountProdotto "
    MyQ = MyQ & " set  MailDocumentazione = '" & apici(kk) & "'"
    MyQ = MyQ & " Where IdProdotto = " & IdProdotto
    MyQ = MyQ & " and   IdAccount = "  & IdAccount
    
    ConnMsde.execute MyQ 
    If Err.Number <> 0 Then 
        MsgErrore = ErroreDb(Err.description)
    End If
    DescIn=""
End if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdProdotto"     ,IdProdotto)
  xx=setValueOfDic(Pagedic,"DescProdotto"   ,DescProdotto)
  xx=setValueOfDic(Pagedic,"IdFornitore"    ,IdFornitore)
  xx=setValueOfDic(Pagedic,"DescFornitore"  ,DescFornitore)
  xx=setValueOfDic(Pagedic,"IdAccount"      ,IdAccount)
  xx=setValueOfDic(Pagedic,"PaginaReturn"   ,PaginaReturn)
  xx=setValueOfDic(Pagedic,"TipoDoc"        ,TipoDoc)
  
  xx=setCurrent(NomePagina,livelloPagina) 
  err.clear 

  DescMail=LeggiCampo("select * from AccountProdotto Where IdProdotto=" & IdProdotto & " and IdAccount=" & IdAccount,"MailDocumentazione")
  
  IdAnagServizio=LeggiCampo("select * from Prodotto Where IdProdotto=" & IdProdotto,"IdAnagServizio")
%>

<% 
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
            <%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;;"
              DescAzione = "Documentazione per gestione/affidamento "
              if IdAnagServizio="CAUZ_DEFI" then 
                 DescAzione = " Documentazione per istruttoria cauzione definitiva"
              end if 
              if TipoDoc = "PROD" then 
                 DescAzione = " Documentazione di prodotto"
              end if 
            %>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <div class="col-11"><h5><%=DescAzione%></h5>
                </div>
            </div>
            <div class="row">
               <div class="col-5">
                  <div class="form-group ">
                     <%xx=ShowLabel("Fornitore")%>
                     <input type="text" readonly class="form-control input-sm" value="<%=DescFornitore%>" >
                  </div>        
               </div>            
               <div class="col-5">
                  <div class="form-group ">
                     <%xx=ShowLabel("Prodotto")%>
                     <input type="text" readonly class="form-control input-sm" value="<%=DescProdotto%>" >
                  </div>        
               </div>            
               
            </div> 
            

            <%
            AddRow=true
            dim CampoDb(10)
            CampoDB(1)="DescDocumento"    
            ElencoOption=";0;Descrizione;1"
            %>        
            <!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
<%
'caricamento tabella 
if Condizione<>"" then 
    Condizione=" and " & Condizione
end if 
        


MySql = "" 
MySql = MySql & " Select a.FlagObbligatorio,a.FlagDataScadenza,a.DITT,a.PEGI,A.PEFI,A.PEGC,b.* "
MySql = MySql & " From AccountProdottoDocAff a, Documento B "
MySql = MySql & " Where A.IdDocumento = B.IdDocumento "
MySql = MySql & " and   A.IdProdotto = " & IdProdotto
MySql = MySql & " and   A.IdAccount = " & IdAccount
MySql = MySql & " and   A.TipoDoc = '" & TipoDoc & "'"
MySql = MySql & Condizione & " order By B.DescDocumento"

'response.write MySql 

Rs.CursorLocation = 3 
Rs.Open MySql, ConnMsde

RecCount=Rs.RecordCount 
if RecCount=0 then 
   RecCount=99
end if 

DescLoaded=""
NumCols = 4
NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
%>
    <!--#include virtual="/gscVirtual/include/CheckRs.asp"-->

<div class="table-responsive">
<table class="table"><tbody>
    <tr>
        <th scope="col" style="width:20%">E-Mail invio documenti</th>
        <td scope="col">
        <input class="form-control" Id="mail0" name ="mail0" type="text" value="<%=DescMail%>">
        </td>
        <th scope="col" style="width:10%">
            <%RiferimentoA="col-2;#;;2;upda;Aggiorna;;updMail();N"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
        </th>
    </tr>
</tbody></table>
 
<table class="table"><tbody>
<thead>
    <tr>
        <th scope="col">Documento
        <%
        myQ="select * from ListaDocumento order By DescListaDocumento"
        response.write ListaDbChangeCompleta (MyQ,"ListaDoc","","IdListaDocumento","DescListaDocumento",0,"","","","","","")
        
        %>
        <a href='#' data-toggle="tooltip" data-placement="top" title="Carica Da Lista" onclick="AttivaFunzione('CALL_FROM','0');">
    <i class="fa fa-2x fa-caret-square-o-left"></i></a> 
        </th>
        <th class="text-center" scope="col" width="10%">Obbligatorio</th>
        <th class="text-center" scope="col" width="10%">Data Scadenza</th>
        <th class="text-center" scope="col" width="10%" >   Ditta</th>
        <th class="text-center" scope="col" width="10%">Pers.Fis.</th>
        <th class="text-center" scope="col" width="10%">Pers.Giu.Cap.</th>
		<th class="text-center" scope="col" width="10%">Pers.Giu.Pers.</th>
        <th scope="col">Azioni</th>        
    </tr>
</thead>

<%
elencoDettaglio=""
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
    Do While Not rs.EOF 
        Primo=Primo+1
        NumRec=NumRec+1
        Id=Rs("IdDocumento")
        DescLoaded=DescLoaded & Id & ";"
        if Rs("FlagObbligatorio")=0 then 
           Obbligatorio=""
        else
           Obbligatorio=" checked "
        end if        
        if Rs("FlagDataScadenza")=0 then 
           Scadenza=""
        else
           Scadenza=" checked "
        end if        
        if Rs("DITT")="" then 
           FlagDITT=""
        else
           FlagDITT=" checked "
        end if                
        if Rs("PEGI")="" then 
           FlagPEGI=""
        else
           FlagPEGI=" checked "
        end if                
        if Rs("PEGC")="" then 
           FlagPEGC=""
        else
           FlagPEGC=" checked "
        end if  
        if Rs("PEFI")="" then 
           FlagPEFI=""
        else
           FlagPEFI=" checked "
        end if                
        %> 
    <tr scope="col"> 
        <td>
            <input class="form-control" Id="IdDocumento<%=Id%>" type="text" readonly value="<%=Rs("DescDocumento")%>">
        </td>
        <td><div class="form-check">
                <input id="checkObb<%=Id%>" <%=Obbligatorio%> name="checkObb<%=Id%>" type="checkbox" value = "S" class="big-checkbox">
            </div>        
        </td>
        <td><div class="form-check">
                <input id="checkSca<%=Id%>" <%=Scadenza%> name="checkSca<%=Id%>" type="checkbox" value = "S" class="big-checkbox">
            </div>        
        </td>        
        <td><div class="form-check text-center">
                <input  id="checkDITT<%=Id%>" <%=FlagDITT%> name="checkDITT<%=Id%>" type="checkbox" value = "DITT" class="big-checkbox">
            </div>        
        </td>
        <td><div class="form-check text-center">
                <input  id="checkPEFI<%=Id%>" <%=FlagPEFI%> name="checkPEFI<%=Id%>" type="checkbox" value = "PEFI" class="big-checkbox">
            </div>        
        </td>
        <td><div class="form-check text-center">
                <input  id="checkPEGC<%=Id%>" <%=FlagPEGC%> name="checkPEGC<%=Id%>" type="checkbox" value = "PEGC" class="big-checkbox">
            </div>        
        </td>		
        <td><div class="form-check text-center">
                <input  id="checkPEGI<%=Id%>" <%=FlagPEGI%> name="checkPEGI<%=Id%>" type="checkbox" value = "PEGI" class="big-checkbox">
            </div>        
        </td>
        <td>
            <%RiferimentoA="col-2;#;;2;dele;Cancella;;SalvaSingoloEdAttiva('DEL'," & Id & ",true,'','','');N"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
            <%RiferimentoA="col-2;#;;2;upda;Aggiorna;;SalvaSingoloEdAttiva('UPD'," & Id & ",true,'','','');N"%>
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
            <%     

            inQue = "(select IdDocumento from AccountProdottoDocAff Where IdAccount=" & IdAccount & " and  IdProdotto= " & IdProdotto & " and TipoDoc='" & TipoDoc & "' ) "
            
            IdRef="IdDocumento" & Id     
            query = ""
            query = query & " Select * from Documento " 
            query = query & " Where IdDocumentoInterno='' and IdDocumento not in " & inQue 
            query = query & " order By DescDocumento"
            'response.write query
            err.clear
            response.write ListaDbChangeCompleta(Query,IdRef,"0","IdDocumento","DescDocumento",0,"","","","","dati assenti","class='form-control form-control-sm'")
            'response.write err.description
            xx="0" & LeggiCampo(query,"IdDocumento")
            %>
        </td>
        <td><div class="form-check">
                <input id="checkObb<%=Id%>" name="checkObb<%=Id%>" type="checkbox" value = "S" class="big-checkbox">
            </div>        
        </td>
        <td><div class="form-check">
                <input id="checkSca<%=Id%>" name="checkSca<%=Id%>" type="checkbox" value = "S" class="big-checkbox">
            </div>        
        </td>        
            <td><div class="form-check text-center">
                    <input id="checkDITT0" checked name="checkDITT0" type="checkbox" value = "DITT" class="big-checkbox">
                </div>        
            </td>
            <td><div class="form-check text-center">
                    <input id="checkPEFI0" checked name="checkPEFI0" type="checkbox" value = "PEFI" class="big-checkbox">
                </div>        
            </td>
            <td><div class="form-check text-center">
                    <input id="checkPEGC0" checked name="checkPEGC0" type="checkbox" value = "PEGC" class="big-checkbox">
                </div>        
            </td> 			
            <td><div class="form-check text-center">
                    <input id="checkPEGI0" checked name="checkPEGI0" type="checkbox" value = "PEGI" class="big-checkbox">
                </div>        
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
