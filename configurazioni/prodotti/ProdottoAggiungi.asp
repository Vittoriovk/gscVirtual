<%
  NomePagina="ProdottoAggiungi.asp"
  titolo="Menu Supervisor - Dashboard"
  default_check_profile="SuperV"
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
<!-- Custom styles for this template -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">
</head>
<script language="JavaScript">

function cambia()
{
   ImpostaValoreDi("Oper","cambia");
   document.Fdati.submit();
}

function localFun(Op,Id)
{
    xx=ImpostaValoreDi("DescLoaded","0");
    xx=ElaboraControlli();
    
     if (xx==false)
       return false;

    ImpostaValoreDi("Oper","update");
    document.Fdati.submit();
}

</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->


<%
  NameLoaded= ""
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdCompagnia=0
  if FirstLoad then 
     IdCompagnia   = "0" & Session("swap_IdCompagnia")
     PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
     if PaginaReturn="" then 
        PaginaReturn = Session("swap_PaginaReturn")
     end if 
  else
     IdCompagnia   = "0" & getValueOfDic(Pagedic,"IdCompagnia")
     OperTabella   = getValueOfDic(Pagedic,"OperTabella")
     PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdCompagnia = cdbl(IdCompagnia)
  
  DescPageOper  = "Associazione prodotto"
  DescCompagnia = LeggiCampo("select * from Compagnia where IdCompagnia=" & IdCompagnia,"DescCompagnia") 
  err.clear
  
  if Oper=ucase("CALL_INS") then 
     Session("TimeStamp")=TimePage
     IdProdottoTemplate = Cdbl("0" & Request("ItemToRemove"))
	 if Cdbl(IdProdottoTemplate)>0 then 
	    DescProdotto = LeggiCampo("Select * from ProdottoTemplate Where IdProdottoTemplate=" & IdProdottoTemplate,"DescProdottoTemplate")
		DescProdotto = DescProdotto & " - " & DescCompagnia
		MyQ = ""
        MyQ = MyQ & " INSERT INTO Prodotto (DescProdotto,IdCompagnia,IdRamo,GiorniDisdetta,DescDatiTecnici"
        MyQ = MyQ & " ,DescGaranzie,FlagPrezzoFisso,Prezzo,IdAnagServizio,IdAnagCaratteristica,IdTrattamentoFiscale"
        MyQ = MyQ & " ,IdListaDocumento,IdListaAffidamento,CodiceProdotto,IdSubRamo,IdProdottoTemplate,IdRischio)" 	
        
		MyQ = MyQ & " select '" & apici(DescProdotto) & "' as DescProdotto," & IdCompagnia & " as idCompagnia,IdRamo,GiorniDisdetta,DescDatiTecnici"
		MyQ = MyQ & " ,DescGaranzie,FlagPrezzoFisso,0 as Prezzo,IdAnagServizio,IdAnagCaratteristica,IdTrattamentoFiscale"
		MyQ = MyQ & " ,IdListaDocumento,IdListaAffidamento,'' as codiceProdotto,IdSubRamo,IdProdottoTemplate, IdRischio"
		MyQ = MyQ & " From ProdottoTemplate " 
		MyQ = MyQ & " Where IdProdottoTemplate = " & NumForDb(IdProdottoTemplate)

        'response.write MyQ 
        ConnMsde.execute MyQ 
        If Err.Number <> 0 Then 
           MsgErrore = ErroreDb(Err.description)
        else 
		   IdProdotto = GetTableIdentity("Prodotto")
		   if Cdbl(IdProdotto)>0 then 
              MyQ = ""
              MyQ = MyQ & " insert into ProdottoOpzione "
		      MyQ = MyQ & " (IdProdotto,IdAccountFornitore,IdOpzione,FlagObbligatorio,Rigo,Ordine"
		      MyQ = MyQ & " ,CostoFisso,PercSuAcquisto,CostoMinimoSuPerc)"
		      MyQ = MyQ & " SELECT " & IdProdotto & " as Idprodotto,0 IdAccountFornitore,IdOpzione,FlagObbligatorio,Rigo,Ordine"
		      MyQ = MyQ & " ,0 as CostoFisso,0 as PercSuAcquisto,0 as CostoMinimoSuPerc"
		      MyQ = MyQ & " FROM ProdottoTemplateOpzione "
		      MyQ = MyQ & " Where IdProdottoTemplate = " & NumForDb(IdProdottoTemplate)
              ConnMsde.execute MyQ 
              MyQ = ""
              MyQ = MyQ & " insert into ProdottoDatoTecnico "
		      MyQ = MyQ & " (IdProdotto,IdDatoTecnico,FlagObbligatorio,Rigo,Ordine)"
		      MyQ = MyQ & " SELECT " & IdProdotto & " as Idprodotto,IdDatoTecnico,FlagObbligatorio,Rigo,Ordine"
		      MyQ = MyQ & " FROM ProdottoTemplateDatoTecnico "
		      MyQ = MyQ & " Where IdProdottoTemplate = " & NumForDb(IdProdottoTemplate)			  
			  ConnMsde.execute MyQ 
		   end if 
           
           response.redirect virtualpath & PaginaReturn
        End If
     end if 
  end if 

  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdCompagnia"  ,IdCompagnia)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
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
                <div class="col-11"><h3>Gestione Prodotto Compagnia:</b> <%=DescPageOper%> </h3>
                </div>
            </div>

      <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->        

      <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            
       <%
       ao_lbd = "Compagnia"        'descrizione label 
       ao_nid = "IdCompagnia0"     'nome ed id
       ao_Att = "0"
       ao_val = IdCompagnia     
       ao_Tex = "select * from Compagnia Where IdCompagnia=" & IdCompagnia
       ao_ids = "IdCompagnia"                  'valore della select 
       ao_des = "DescCompagnia"                'valore del testo da mostrare 
       ao_cla = ""                        'azzero classe
       ao_Eve = "" 
       ao_Plh = ""                        'indica cosa mettere in caso di vuoto
       ao_Cla = "class='form-control form-control-sm' disabled"
    %>
    <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->  
	
      <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            
       <%
       ao_lbd = "Ramo"             'descrizione label 
       ao_nid = "IdRamo0"          'nome ed id
       idRamo = GetDiz(DizDatabase,"IdRamo")
       ao_Att = "1"
       if oper="CAMBIA" then 
          idRamo=Request("IdRamo0")
       end if
       ao_val = idRamo     
       ao_Tex = "select * from Ramo "
       'non modificabile se IdProdotto>0 
       ao_Tex = ao_Tex & " order by DescRamo"
       'response.write ao_Tex
       ao_ids = "IdRamo"                  'valore della select 
       ao_des = "DescRamo"                'valore del testo da mostrare 
       ao_cla = ""                        'azzero classe
       ao_Eve = "cambia()" 'azzero evento
                             'indica se deve mettere vuoto 
       ao_Plh = ""                        'indica cosa mettere in caso di vuoto
       ao_Cla = "class='form-control form-control-sm'" 
    %>
         
    <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"--> 

    <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            
       <%
       ao_lbd = "Rischio"             'descrizione label 
       ao_nid = "IdRischio0"          'nome ed id
       idRischio = GetDiz(DizDatabase,"IdRischio")
       if oper="CAMBIA" then 
          idRischio=Request("idRischio0")
       end if	   
       ao_Att = "1"
       disab=" "
       if SoloLettura=true or Cdbl(IdProdotto)>0 then
          disab=" disabled "
       end if        
       ao_val = idRischio     
       ao_Tex = "select * from Rischio Where IdRamo=" & idRamo
       ao_Tex = ao_Tex & " order by DescRischio"
       'response.write ao_Tex
       ao_ids = "IdRischio"                  'valore della select 
       ao_des = "DescRischio"                'valore del testo da mostrare 
       ao_cla = ""                        'azzero classe
       ao_Eve = "cambia()"
       ao_Plh = ""                        'indica cosa mettere in caso di vuoto
       ao_Cla = "class='form-control form-control-sm'" & disab 
    %>
    <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->    
      <%
      qP = ""
      if Cdbl(idRamo)>0 or Cdbl(IdRischio)>0 then 
         qP = qP & " select * from ProdottoTemplate "
         qP = qP & " Where 1=1  " 
         if Cdbl(idRamo)>0 then 
            qP = qP & " and IdRamo = " & IdRamo
         end if 
         if Cdbl(idRischio)>0 then 
            qP = qP & " and IdRischio = " & IdRischio
         end if 
         qP = qP & " and IdProdottoTemplate not in "
         qP = qP & " (select IdProdottoTemplate from Prodotto Where IdCompagnia=" & IdCompagnia & " ) "
         'response.write qP
         
      end if 
      
      if qP<>"" then 
      %>
         <div class="table-responsive"><table class="table"><tbody>
         <thead>
            <tr>
               <th scope="col">Prodotto</th>
               <th scope="col">Azioni</th>
            </tr>
         </thead>
         <%
         Rs.CursorLocation = 3 
         Rs.Open qP, ConnMsde
         MsgNoData  = ""
         %>
         <!--#include virtual="/gscVirtual/include/CheckRs.asp"-->         
         <%
         if MsgNoData="" then 
            Do While Not rs.EOF 
         %>
           <tr scope="col">
              <td>
			      <input class="form-control" type="text" readonly value="<%=Rs("DescProdottoTemplate")%>">
			  </td>  
			  <td>
			  <a href="#" title="Inserisci" onclick="AttivaFunzione('CALL_INS','<%=Rs("IdProdottoTemplate")%>');">
                   <i class="fa fa-2x fa-plus-square"></i></a>
		      </td>
			  
         <%   rs.MoveNext
            Loop    
         end if 
         %>
         </tbody></table></div> <!-- table responsive fluid -->
      <%
      end if 
      
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
