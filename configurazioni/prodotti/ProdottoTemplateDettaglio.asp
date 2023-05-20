<%
  NomePagina="ProdottoTemplateDettaglio.asp"
  titolo="Gestione Template Prodotto"
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
  NameLoaded= NameLoaded & "DescProdotto,TE" 
  NameLoaded= NameLoaded & ";IdRamo,LI;IdRischio,LI"
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdProdotto=0
  if FirstLoad then 
     IdProdotto   = "0" & Session("swap_IdProdotto")
     if Cdbl(IdProdotto)=0 then 
        IdProdotto = cdbl("0" & getValueOfDic(Pagedic,"IdProdotto"))
     end if 
     OperTabella   = Session("swap_OperTabella")
     PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
     if PaginaReturn="" then 
        PaginaReturn = Session("swap_PaginaReturn")
     end if 
  else
     IdProdotto   = "0" & getValueOfDic(Pagedic,"IdProdotto")
     OperTabella   = getValueOfDic(Pagedic,"OperTabella")
     PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdProdotto = cdbl(IdProdotto)
  'inizio elaborazione pagina

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
  
   xx=SetDiz(DizDatabase,"IdProdotto",0)
   xx=SetDiz(DizDatabase,"DescProdotto","")
   xx=SetDiz(DizDatabase,"IdRamo",0)
   xx=SetDiz(DizDatabase,"GiorniDisdetta",0)
   xx=SetDiz(DizDatabase,"FlagPrezzoFisso","0")
   xx=SetDiz(DizDatabase,"IdAnagServizio","")
   xx=SetDiz(DizDatabase,"IdTrattamentoFiscale",0)
   xx=SetDiz(DizDatabase,"IdAnagCaratteristica",0)
   xx=SetDiz(DizDatabase,"IdListaDocumento",0)
   xx=SetDiz(DizDatabase,"IdRischio",0)
  
  'recupero i dati 
  if cdbl(IdProdotto)>0 then
      MySql = ""
      MySql = MySql & " Select * From  ProdottoTemplate "
      MySql = MySql & " Where IdProdottoTemplate=" & IdProdotto
      xx=GetInfoRecordset(DizDatabase,MySql)
  end if 
  
  if OperTabella="CALL_DEL" then 
     SoloLettura=true
  end if 
  'inserisco il fornitore 
  descD      = Request("DescProdotto0")
  IdRamo     = cdbl("0" & Request("IdRamo0"))
  IdRischio  = cdbl("0" & Request("IdRischio0"))
  idAnaCar   = LeggiCampo("select * from Rischio Where IdRischio=" & IdRischio ,"IdAnagCaratteristica")
  idAna      = LeggiCampo("select * from AnagCaratteristica Where IdAnagCaratteristica=" & IdAnaCar ,"IdAnagServizio")
  giorn  = 0
  descT  = ""
  descG = ""  
  if Request("FlagPrezzoFisso0") = "S" then 
     FlagP = 1
  else
     FlagP = 0
  end if 
  idTra    = cdbl("0" & Request("IdTrattamentoFiscale0"))
  IdListaDocumento   = cdbl("0" & Request("IdListaDocumento0"))
  IdListaAffidamento = cdbl("0" & Request("IdListaAffidamento0"))  
  err.clear
  
  if Oper=ucase("update") and OperTabella="CALL_INS" then 
    Session("TimeStamp")=TimePage
    KK="0"
    MyQ = "" 
    MyQ = MyQ & " INSERT INTO ProdottoTemplate (DescProdottoTemplate,IdRamo,GiorniDisdetta,DescDatiTecnici"
    MyQ = MyQ & " ,DescGaranzie,FlagPrezzoFisso,IdAnagServizio,IdAnagCaratteristica,IdTrattamentoFiscale"
    MyQ = MyQ & " ,IdListaDocumento,IdRischio,IdListaAffidamento)" 
    MyQ = MyQ & " values (" 
    MyQ = MyQ & " '" & apici(descD) & "'"
    MyQ = MyQ & ", " & idRamo
    MyQ = MyQ & ", " & giorn 
    MyQ = MyQ & ",'" & apici(descT) & "'"
    MyQ = MyQ & ",'" & apici(descG) & "'"
    MyQ = MyQ & ", " & FlagP
    MyQ = MyQ & ",'" & apici(idAna) & "'"
    MyQ = MyQ & ", " & numFordb(idAnaCar)
    MyQ = MyQ & ", " & numFordb(idTra)
    MyQ = MyQ & ", " & numFordb(IdListaDocumento)
    MyQ = MyQ & ", " & numFordb(idRischio)
	MyQ = MyQ & ", " & numFordb(IdListaAffidamento)
    MyQ = MyQ & ")"
response.write Myq
    ConnMsde.execute MyQ 
    If Err.Number <> 0 Then 
        MsgErrore = ErroreDb(Err.description)
    else 
       response.redirect virtualpath & PaginaReturn
    End If
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdProdotto)>0 then 

     MyQ = "" 
     MyQ = MyQ & " Update ProdottoTemplate "
     MyQ = MyQ & " Set DescProdottoTemplate = '" & apici(descD) & "'"
     MyQ = MyQ & ",IdTrattamentoFiscale="        & numFordb(idTra)  
     MyQ = MyQ & ",IdListaDocumento="            & NumforDb(IdListaDocumento)
	 MyQ = MyQ & ",IdListaAffidamento="          & NumforDb(IdListaAffidamento)
     MyQ = MyQ & ",FlagPrezzoFisso='"            & flagP & "'"
     MyQ = MyQ & " Where IdProdottoTemplate = "  & IdProdotto

    ConnMsde.execute MyQ 
    If Err.Number <> 0 Then 
        MsgErrore = ErroreDb(Err.description)
    else 
	    'aggiorno ramo 
        MyQ = "" 
        MyQ = MyQ & " Update Prodotto Set "
        MyQ = MyQ & " IdRischio="                   & NumForDb(idSub)
        MyQ = MyQ & " Where IdProdottoTemplate = "  & IdProdotto
        ConnMsde.execute MyQ

        'aggiorno Trattamento fiscale se non cambiato 
        oldTra = GetDiz(DizDatabase,"IdTrattamentoFiscale") 
        MyQ = "" 
        MyQ = MyQ & " Update Prodotto Set "
        MyQ = MyQ & " IdTrattamentoFiscale="        & numFordb(idTra)  
        MyQ = MyQ & " Where IdProdottoTemplate = "  & IdProdotto
        MyQ = MyQ & " and IdTrattamentoFiscale = "  & numFordb(oldTra) 
        ConnMsde.execute MyQ

        'aggiorno lista documenti se non cambiata
        oldTra = GetDiz(DizDatabase,"IdListaDocumento") 
        MyQ = "" 
        MyQ = MyQ & " Update Prodotto Set "
        MyQ = MyQ & " IdListaDocumento="            & NumforDb(IdListaDocumento)
        MyQ = MyQ & " Where IdProdottoTemplate = "  & IdProdotto
        MyQ = MyQ & " and IdListaDocumento = "      & NumforDb(oldTra)		
        ConnMsde.execute MyQ

        'aggiorno flagprezzo
        oldTra = GetDiz(DizDatabase,"FlagPrezzoFisso") 
        MyQ = "" 
        MyQ = MyQ & " Update Prodotto Set "
        MyQ = MyQ & " FlagPrezzoFisso='"            & flagP & "'"
        MyQ = MyQ & " Where IdProdottoTemplate = "  & IdProdotto
        MyQ = MyQ & " and FlagPrezzoFisso = '"      & oldTra & "'"
        ConnMsde.execute MyQ

        response.redirect virtualpath & PaginaReturn
    End If    
  end if
  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdProdotto)>0 then 
     MsgErrore = VerificaDel("ProdottoTemplate",IdProdotto) 
     if MsgErrore = "" then   
        MyQ = "" 
        MyQ = MyQ & " Delete from ProdottoTemplate "
        MyQ = MyQ & " Where IdProdottoTemplate = " & IdProdotto
        ConnMsde.execute MyQ 
        MyQ = "" 
        MyQ = MyQ & " Delete from Prodotto "
        MyQ = MyQ & " Where IdProdottoTemplate = " & IdProdotto
        ConnMsde.execute MyQ 
        
        If Err.Number <> 0 Then 
            MsgErrore = ErroreDb(Err.description)
        else 
           response.redirect virtualpath & PaginaReturn
        End If    
    end if 
  end if    
  
   DescPageOper="Aggiornamento"
   if OperTabella="V" then 
      DescPageOper = "Consultazione"
   elseIf OperTabella="CALL_INS" then 
      DescPageOper = "Inserimento"
   elseIf OperTabella="CALL_DEL" then 
      DescPageOper = "Cancellazione"      
   end if
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdProdotto"   ,IdProdotto)
  xx=setValueOfDic(Pagedic,"OperTabella"  ,OperTabella)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  checkedSi=""
  checkedNo=""
  if GetDiz(DizDatabase,"FlagScadenza") = "1" then 
     checkedSi = " checked "
  else
     checkedNo = " checked "
  end if 
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
                <div class="col-11"><h3>Gestione Template Prodotto:</b> <%=DescPageOper%> </h3>
                </div>
            </div>

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            
   <%
   l_Id = "0"
   if oper="CAMBIA" then 
      valo = Request("DescProdotto0")
   else
      valo = GetDiz(DizDatabase,"DescProdottoTemplate") 
   end if
   ao_lbd = "Descrizione Template"       'descrizione label 
   ao_nid = "DescProdotto" & l_Id            'nome ed id
   ao_val = "|value=" & valo      'valore di default
   ao_Plh = "|placeholder=Descrizione Prodotto"              'placeholder
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->        

   
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
   disab=""
   if Cdbl(IdProdotto)>0 then 
      ao_Tex = ao_Tex & " Where IdRamo=" & ao_val
      ao_Att = "0"
      disab=" disabled "
   end if 
   ao_Tex = ao_Tex & " order by DescRamo"
   'response.write ao_Tex
   ao_ids = "IdRamo"                  'valore della select 
   ao_des = "DescRamo"                'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = "cambia()" 'azzero evento
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'" & disab  
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->  

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            
   <%
   ao_lbd = "Rischio"             'descrizione label 
   ao_nid = "IdRischio0"          'nome ed id
   idRischio = GetDiz(DizDatabase,"IdRischio")
   if oper="CAMBIA" then 
      idRischio=Request("idRischio0")
   end if   
   ao_Att = "1"
   ao_val = idRischio     
   ao_Tex = "select * from Rischio Where IdRamo=" & idRamo
   'non modificabile se IdProdotto>0 
   disab=""   
   if Cdbl(IdProdotto)>0 then 
      ao_Tex = ao_Tex & " and IdRischio=" & ao_val
      ao_Att = "0"
      disab=" disabled "
   end if 
   
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
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            
   <%
   ao_lbd = "Tratt.Fiscale"             'descrizione label 
   ao_nid = "IdTrattamentoFiscale0"          'nome ed id
   ao_val = GetDiz(DizDatabase,"IdTrattamentoFiscale")
   ao_Tex = "select * from TrattamentoFiscale"
   disab="  "
   if SoloLettura=true then
      ao_Tex = ao_Tex & " where IdTrattamentoFiscale=" & ao_val  
      disab=" disabled "
   end if 
   ao_Tex = ao_Tex & " order By DescTrattamentoFiscale"
   ao_ids = "IdTrattamentoFiscale"             'valore della select 
   ao_des = "DescTrattamentoFiscale"           'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'" & disab  
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->  
  
   <%
   'response.write IdRischio 
   if Cdbl("0" & IdRischio)=0 then 
      FlagDocu = 0
	  FlagAffi = 0
	  PlagPrez = 0
   else 
      idAnaCar   = LeggiCampo("select * from Rischio Where IdRischio=" & IdRischio ,"IdAnagCaratteristica")
	  qAna = "select * from AnagCaratteristica Where IdAnagCaratteristica=" & IdAnaCar
	  
	  FlagDocu   = Cdbl("0" & LeggiCampo(qAna ,"FlagDocProd"))
	  'response.write qAna & " " & FlagDocu
	  FlagAffi   = Cdbl("0" & LeggiCampo(qAna ,"FlagDocAffi"))
	  FlagPrez   = Cdbl("0" & LeggiCampo(qAna ,"FlagPrezzo" ))
   end if    
   
   
   %>
  
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            

   <%if cdbl(FlagDocu)>0 then 
   ao_lbd = "Lista Documenti Prodotto"             'descrizione label 
   ao_nid = "IdListaDocumento0"          'nome ed id
   ao_val = GetDiz(DizDatabase,"IdListaDocumento")
   ao_Tex = "select * from ListaDocumento"
   disab="  "
   if SoloLettura=true then
      ao_Tex = ao_Tex & " where IdListaDocumento=" & ao_val  
      disab=" disabled "
   end if 
   ao_Tex = ao_Tex & " order By DescListaDocumento"
   ao_ids = "IdListaDocumento"             'valore della select 
   ao_des = "DescListaDocumento"           'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'" & disab  
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->  
   <%end if %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->            
   <%
   if cdbl(FlagAffi)>0 then
   ao_lbd = "Lista Documenti Affidamento"             'descrizione label 
   ao_nid = "IdListaAffidamento0"          'nome ed id
   ao_val = GetDiz(DizDatabase,"IdListaAffidamento")
   ao_Tex = "select * from ListaDocumento"
   disab="  "
   if SoloLettura=true then
      ao_Tex = ao_Tex & " where IdListaDocumento=" & ao_val  
      disab=" disabled "
   end if 
   ao_Tex = ao_Tex & " order By DescListaDocumento"
   ao_ids = "IdListaDocumento"             'valore della select 
   ao_des = "DescListaDocumento"           'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'" & disab  
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->  
   <%end if %>
   <%if cdbl(FlagPrez)>0 then %>
   <div class="row  " >
      <div class="col-2">
      <p class="font-weight-bold">Prezzo fisso</p>
   </div>
   <div class = "col-1">
   <% 
   if Trim(GetDiz(DizDatabase,"FlagPrezzoFisso")) = "1" then
      FlagPrezzoFisso = " checked "
   else
      FlagPrezzoFisso = ""
   end if 
   %>
      <input id="FlagPrezzoFisso<%=l_Id%>" <%=FlagPrezzoFisso%> name="FlagPrezzoFisso<%=l_Id%>" 
             type="checkbox" value = "S" class="big-checkbox" >
            <span class="font-weight-bold">SI</span>
   </div>
   <%end if %>
  
</div> 

   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->        
    
   <%if SoloLettura=false then%>
        <div class="row"><div class="mx-auto">
        <%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
        <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
        </div></div>
        <div class="row">
            <div class="col">
                &nbsp;
            </div>
        </div>
    <%elseif OperTabella="CALL_DEL" then  %>
        <div class="row"><div class="mx-auto">
        <%RiferimentoA="center;#;;2;save;Rimuovi; Rimuovi;localFun('submit','0');S"%>
        <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
        </div></div>
        <div class="row">
            <div class="col">
                &nbsp;
            </div>
        </div>    
   <%end if %>


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
