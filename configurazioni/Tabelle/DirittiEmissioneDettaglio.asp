<%
  NomePagina="DirittiEmissioneDettaglio.asp"
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
<!-- Custom styles for this template -->
<link href="<%=VirtualPath%>/css/simple-sidebar.css" rel="stylesheet">
</head>
<script language="JavaScript">
function cambia()
{
   ImpostaValoreDi("Oper","cambia");
   document.Fdati.submit();
}
function setTipo(x)
{
   if (x=='and')
      ImpostaValoreDi("EmisTipoCalcValue","FixAndPerc");
   else
      ImpostaValoreDi("EmisTipoCalcValue","FixOrPerc");
}

function localFun(Op,Id)
{
	xx=ImpostaValoreDi("DescLoaded","0");
	xx=ElaboraControlli();
	
 	if (xx==false)
	   return false;
	 
	var vCalc = ValoreDi("EmisTipoCalcValue");
	if (vCalc=="FixOrPerc") {
	    var perc  = GetNumberAsFloat(ValoreDi("EmisSysPerc0"));
		var fisso = GetNumberAsFloat(ValoreDi("EmisSysFix0"));
		if ((perc==0 && fisso==0) || (perc>0 && fisso>0)) {
		   xx=ImpostaColoreFocus("EmisSysPerc0","S","yellow");
		   xx=ImpostaColoreFocus("EmisSysFix0","S","yellow");
		   alert("Indicare un solo valore positivo per fisso e percentuale");
		   return false;
		}
	}
	
	
	var vProf = ValoreDi("IdProfiloProdotto0");
	var vProd = ValoreDi("IdProdotto0");
	
 	if (!(vProf=='-1' || vProf=='0') && !(vProd=='-1')) {
	   alert('selezionare solo uno fra gruppo e prodotto');
	   return false;
    }
	
	ImpostaValoreDi("Oper","update");
	document.Fdati.submit();
}
</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

 
 <!-- javascript locale -->
<script>
function localSubmit(Op)
{
var xx;
   xx=false;
   if (Op=="submit")
   xx=ElaboraControlli();
   
   if (xx==false)
      return false;
  
   ImpostaValoreDi("Oper","update");
   document.Fdati.submit(); 
}
</script>

<%

   NameLoaded = "DescDirittiEmissione,TE"
   NameLoaded = NameLoaded & ";EmisSysFix,FLZ;EmisSysPerc,FLQ;EmisSysMin,FLZ"
   NameLoaded = NameLoaded & ";EmisReteFix,FLZ;EmisRetePerc,FLQ;EmisReteMin,FLZ"
   NameLoaded = NameLoaded & ";AggiMin,FLZ;AggiMax,FLZ;AggiDef,FLZ"
   
   NameRangeN = ""
   NameRangeN = NameRangeN &  "EmisReteFix;EmisReteMin;0;99999"
   NameRangeN = NameRangeN &  ";EmisReteFix;EmisSysFix;0;99999"
   NameRangeN = NameRangeN &  ";EmisReteMin;EmisSysMin;0;99999"
   
   NameRangeN = NameRangeN &  ";AggiMin;AggiMax;0;99999"
   NameRangeN = NameRangeN &  ";AggiMin;AggiDef;0;99999"
   NameRangeN = NameRangeN &  ";AggiDef;AggiMax;0;99999"

  
   FirstLoad=(Request("CallingPage")<>NomePagina)
   IdDirittiEmissione=0
   if FirstLoad then 
      IdDirittiEmissione   = "0" & Session("swap_IdDirittiEmissione")
      if Cdbl(IdDirittiEmissione)=0 then 
         IdDirittiEmissione = cdbl("0" & getValueOfDic(Pagedic,"IdDirittiEmissione"))
      end if 
      OperTabella   = Session("swap_OperTabella")
      PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
      if PaginaReturn="" then 
        PaginaReturn = Session("swap_PaginaReturn")
      end if 
   else
      IdDirittiEmissione   = "0" & getValueOfDic(Pagedic,"IdDirittiEmissione")
      OperTabella   = getValueOfDic(Pagedic,"OperTabella")
      PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   end if 
   IdDirittiEmissione = cdbl(IdDirittiEmissione)
  
  if OperTabella="CALL_DEL" then 
     SoloLettura=true
  end if 

  DescDirittiEmissione   = Request("DescDirittiEmissione0")
  IdAccountFornitore     = TestNumeroPos(Request("IdAccountFornitore0"))
  IdRamo                 = TestNumeroPos(Request("IdRamo0"))
  IdProfiloProdotto              = TestNumeroPos(Request("IdProfiloProdotto0"))
  IdAnagServizio         = Request("IdAnagServizio0")
  if IdAnagServizio="-1" then 
     IdAnagServizio=""
  end if 
  IdAnagCaratteristica   = TestNumeroPos(Request("IdAnagCaratteristica0")) 
  IdProdotto             = TestNumeroPos(Request("IdProdotto0")) 
  IdCompagnia            = TestNumeroPos(Request("IdCompagnia0")) 
  Importo                = TestNumeroPos(Request("Importo0"))
  ImportoIntermediazione    = TestNumeroPos(Request("ImportoIntermediazione0"))
  ImportoIntermediazioneMin = TestNumeroPos(Request("ImportoIntermediazioneMin0"))
  ImportoIntermediazioneMax = TestNumeroPos(Request("ImportoIntermediazioneMax0"))

  EmisTipoCalc = Request("EmisTipoCalc0")
  if EmisTipoCalc="" then 
     EmisTipoCalc="FixAndPerc"
  end if 
  EmisSysFix   = TestNumeroPos(Request("EmisSysFix0"))
  EmisSysPerc  = TestNumeroPos(Request("EmisSysPerc0"))
  EmisSysMin   = TestNumeroPos(Request("EmisSysMin0"))
  EmisReteFix  = TestNumeroPos(Request("EmisReteFix0"))
  EmisRetePerc = TestNumeroPos(Request("EmisRetePerc0"))
  EmisReteMin  = TestNumeroPos(Request("EmisReteMin0"))
  InteTipoCalc = Request("InteTipoCalcValue")
  'response.write InteTipoCalc
  'response.end 
  if InteTipoCalc="" then 
     InteTipoCalc="FixAndPerc"
  end if 
  InteSysFix   = 0 'TestNumeroPos(Request("InteSysFix0"))
  InteSysPerc  = 0 'TestNumeroPos(Request("InteSysPerc0"))
  InteSysMin   = 0 'TestNumeroPos(Request("InteSysMin0"))
  InteReteFix  = 0 'TestNumeroPos(Request("InteReteFix0"))
  InteRetePerc = 0 'TestNumeroPos(Request("InteRetePerc0"))
  InteReteMin  = 0 'TestNumeroPos(Request("InteReteMin0"))
  
  AggiTipoCalc = Request("AggiTipoCalc0")
  if AggiTipoCalc="" then 
     AggiTipoCalc="Fix"
  end if 
  AggiMin     = TestNumeroPos(Request("AggiMin0"))
  AggiMax     = TestNumeroPos(Request("AggiMax0"))
  AggiDef     = TestNumeroPos(Request("AggiDef0"))
  AggiPercSys = 0 'TestNumeroPos(Request("AggiPercSys0"))
  
  MsgNoData=""
  
  IdAccountRiferimento=0


  if Oper=ucase("update") and OperTabella="CALL_INS" then 
    Session("TimeStamp")=TimePage
	MyQ = "" 
	MyQ = MyQ & " INSERT INTO DirittiEmissione (IdAccount,IdAccountFornitore,IdProfiloProdotto,IdProdotto,TipoRegola) " 
	MyQ = MyQ & " values (" & IdAccountRiferimento
	MyQ = MyQ & ", " & numForDb(IdAccountFornitore)	
	MyQ = MyQ & ", " & numForDb(IdProfiloProdotto)
    MyQ = MyQ & ", " & IdProdotto
	MyQ = MyQ & ",'" & apici(Session("LoginTipoUtente")) & "'"
    MyQ = MyQ & " )" 
	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   OperTabella="CALL_UPD"
	   IdDirittiEmissione = GetTableIdentity("DirittiEmissione")
	End If
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdDirittiEmissione)>0 then 
	MyQ = "" 
	MyQ = MyQ & " Update DirittiEmissione "
	MyQ = MyQ & " Set IdAccountFornitore = " & IdAccountFornitore
	MyQ = MyQ & ",IdProfiloProdotto= " & NumForDb(IdProfiloProdotto)
	MyQ = MyQ & ",DescDirittiEmissione='" & apici(DescDirittiEmissione) & "'"
	MyQ = MyQ & ",IdProdotto= "   & IdProdotto
    MyQ = MyQ & ",EmisTipoCalc='" & EmisTipoCalc & "'"
    MyQ = MyQ & ",EmisSysFix= "   & NumForDb(EmisSysFix)
    MyQ = MyQ & ",EmisSysPerc= "  & NumForDb(EmisSysPerc)
    MyQ = MyQ & ",EmisSysMin= "   & NumForDb(EmisSysMin)
    MyQ = MyQ & ",EmisReteFix= "  & NumForDb(EmisReteFix)
    MyQ = MyQ & ",EmisRetePerc= " & NumForDb(EmisRetePerc)
    MyQ = MyQ & ",EmisReteMin= "  & NumForDb(EmisReteMin)
    MyQ = MyQ & ",InteTipoCalc='" & InteTipoCalc & "'"
    MyQ = MyQ & ",InteSysFix= "   & NumForDb(InteSysFix)
    MyQ = MyQ & ",InteSysPerc= "  & NumForDb(InteSysPerc)
    MyQ = MyQ & ",InteSysMin= "   & NumForDb(InteSysMin)
    MyQ = MyQ & ",InteReteFix= "  & NumForDb(InteReteFix)
    MyQ = MyQ & ",InteRetePerc= " & NumForDb(InteRetePerc)
    MyQ = MyQ & ",InteReteMin= "  & NumForDb(InteReteMin)
	MyQ = MyQ & ",AggiTipoCalc='" & AggiTipoCalc & "'"
    MyQ = MyQ & ",AggiMin= "      & NumForDb(AggiMin)
	MyQ = MyQ & ",AggiMax= "      & NumForDb(AggiMax)
	MyQ = MyQ & ",AggiDef= "      & NumForDb(AggiDef)
	MyQ = MyQ & ",AggiPercSys= "  & NumForDb(AggiPercSys)
	
	
	MyQ = MyQ & " Where IdDirittiEmissione = " & IdDirittiEmissione
    'response.write MyQ 

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
  end if

  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdDirittiEmissione)>0 then 
	MyQ = "" 
	MyQ = MyQ & " Delete from DirittiEmissione "
	MyQ = MyQ & " Where IdDirittiEmissione = " & IdDirittiEmissione

	ConnMsde.execute MyQ 
	If Err.Number <> 0 Then 
		MsgErrore = ErroreDb(Err.description)
	else 
	   response.redirect virtualpath & PaginaReturn
	End If	
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
   xx=setValueOfDic(Pagedic,"IdDirittiEmissione" ,IdDirittiEmissione)
   xx=setValueOfDic(Pagedic,"OperTabella"        ,OperTabella)
   xx=setValueOfDic(Pagedic,"PaginaReturn"       ,PaginaReturn)
   xx=setCurrent(NomePagina,livelloPagina) 

   DescLoaded="0"  
  
  'recupero i dati 
  if cdbl(IdDirittiEmissione)>0 and Oper<>ucase("cambia") then
	  MySql = ""
	  MySql = MySql & " Select * From  DirittiEmissione "
	  MySql = MySql & " Where IdDirittiEmissione=" & IdDirittiEmissione
 
      Set Rs = Server.CreateObject("ADODB.Recordset")

      Rs.CursorLocation = 3 
      Rs.Open MySql, ConnMsde 
	  DescDirittiEmissione   = rs("DescDirittiEmissione")
	  IdAccount              = RS("IdAccount")
      IdAccountFornitore     = RS("IdAccountFornitore")
	  IdProfiloProdotto      = RS("IdProfiloProdotto")
      IdProdotto             = RS("IdProdotto") 
      
      EmisTipoCalc = RS("EmisTipoCalc")
      EmisSysFix   = RS("EmisSysFix")
      EmisSysPerc  = RS("EmisSysPerc")
      EmisSysMin   = RS("EmisSysMin")
      EmisReteFix  = RS("EmisReteFix")
      EmisRetePerc = RS("EmisRetePerc")
      EmisReteMin  = RS("EmisReteMin")
      InteTipoCalc = RS("InteTipoCalc")
      InteSysFix   = RS("InteSysFix")
      InteSysPerc  = RS("InteSysPerc")
      InteSysMin   = RS("InteSysMin")
      InteReteFix  = RS("InteReteFix")
      InteRetePerc = RS("InteRetePerc")
      InteReteMin  = RS("InteReteMin")
      AggiTipoCalc = RS("AggiTipoCalc")
      AggiMin      = RS("AggiMin")
      AggiMax      = RS("AggiMax")
      AggiDef      = RS("AggiDef")
      AggiPercSys  = RS("AggiPercSys")
      rs.close 
  end if 
  if cdbl(IdDirittiEmissione)=0 then 
      EmisSysFix   = ""
      EmisSysPerc  = ""
      EmisSysMin   = ""
      EmisReteFix  = ""
      EmisRetePerc = ""
      EmisReteMin  = ""
      InteSysFix   = ""
      InteSysPerc  = ""
      InteSysMin   = ""
      InteReteFix  = ""
      InteRetePerc = ""
      InteReteMin  = ""
      AggiMin      = ""
      AggiMax      = ""
      AggiDef      = ""
      AggiPercSys  = ""
  
  end if 
   
  'xx=DumpDic(SessionDic,NomePagina)
%>
<div class="d-flex" id="wrapper">
	<%
	  Session("opzioneSidebar")="dash"
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
			<%RiferimentoA="col-1 text-center;" & virtualpath & PaginaReturn & ";;2;prev;Indietro;;"%>
			<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<div class="col-11"><h3>Gestione Diritti Di Emissione : <%=DescPageOper%> </b> </h3>
				</div>
			</div>
   <br>
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->                  
   <%
   ao_lbd = "Descrizione diritti "       'descrizione label 
   ao_nid = "DescDirittiEmissione0"          'nome ed id
   ao_val = "|value=" & DescDirittiEmissione       'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->       
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   	<%
	   qPro = ""
	   qPro = " from prodotto Where 1=1 "
   
	   ao_lbd = "Fornitore"             'descrizione label 
       ao_nid = "IdAccountFornitore0"          'nome ed id
       ao_val = IdAccountFornitore
	   ao_Att = "1"                       'indica se deve mettere vuoto 
	   ao_Tex = "select * from Fornitore order By DescFornitore"
	   'response.write ao_Tex
	   ao_ids = "IdAccount"             'valore della select 
	   ao_des = "DescFornitore"           'valore del testo da mostrare 
	   ao_cla = ""                        'azzero classe
	   ao_Eve = "cambia()"                'azzero evento
	   
	   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
	   ao_Cla = "class='form-control form-control-sm'"	  
	%>
	<!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->   
	
  <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   entity="ProfiloProdotto"
   ao_lbd = "Raggruppamento Prodotti"        
   ao_nid = "Id" & entity & "0"             
   if Oper=ucase("refresh") then 
      IdComp = request("Id" & entity & "0")     'valore di default
   else 
      idComp = IdProfiloProdotto 'valore di default
   end if 
   ao_val = IdComp
   ao_Tex = "SELECT * From " & entity
   ao_Tex = ao_Tex & " Where IdTipoProfilo = 'GRUPPO' " 
   ao_Tex = ao_Tex & " order By Desc" & entity
   ao_ids = "Id" & entity             'valore della select 
   ao_des = "Desc" & entity           'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = "refresh()"               'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   
  
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->
   
<div class="row   " >
   <div class="col-2">
      <p class="font-weight-bold"></p>
   </div>
   <div class = "col-8">
         <b>Oppure</B>
   </div>
   <div class="col-2">
      <p class="font-weight-bold"> </p>
   </div>

</div> 

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   	<%
	   EsisteProd=true 
	   ao_lbd = "Prodotto"             'descrizione label 
       ao_nid = "IdProdotto0"          'nome ed id
       ao_val = IdProdotto
	   ao_Att = "1"                       'indica se deve mettere vuoto 
	   ao_Tex = "select *  " & qPro
	   if cdbl(IdAccountFornitore)>0 then 
	      qIn = "select IdProdotto From AccountProdotto Where IdAccount=" & IdAccountFornitore
		  ao_Tex = ao_Tex & " and IdProdotto in (" & qIn & ") "
		  if cdbl("0" & LeggiCampo(ao_Tex,"IdProdotto"))=0 then 
		     EsisteProd=false 
		  end if 
	   end if 
	   ao_Tex = ao_Tex & " order By DescProdotto"
	   'response.write ao_Tex
	   ao_ids = "IdProdotto"             'valore della select 
	   ao_des = "DescProdotto"           'valore del testo da mostrare 
	   ao_cla = ""                        'azzero classe
	   ao_Eve = ""                        'azzero evento  
	   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
	   ao_Cla = "class='form-control form-control-sm'"	  
	%>
	<!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->  
 
    <%
	FlagAnd=""
	FlagOr =""
	if EmisTipoCalc="FixAndPerc" then 
	   FlagAnd = " checked "
	else
	   EmisTipoCalc = "FixOrPerc"
	   FlagOr  = " checked "
	end if 
	
	%>
	<div class="row">
	   <div class="col-3">
            <div class="form-group ">
		     <%xx=ShowLabel("Diritti di Emissione piattaforma")%><br>
             <input id="EmisTipoCalc0" <%=FlagAnd%> name="EmisTipoCalc0" 
		       type="radio" value = "FixAndPerc" class="big-checkbox" onclick="setTipo('and')">
               <span class="font-weight-bold">Fisso pi&ugrave; Perc.</span>
             <input id="EmisTipoCalc0" <%=FlagOr%>  name="EmisTipoCalc0" 
		       type="radio" value = "FixOrPerc" class="big-checkbox" onclick="setTipo('or')">
               <span class="font-weight-bold">Fisso o Perc.</span>
	        </div>
	   </div>
	   <input type="hidden" name="EmisTipoCalcValue" id="EmisTipoCalcValue" value="<%=EmisTipoCalc%>">
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Importo fisso &euro;")%>
            <input value="<%=EmisSysFix%>"  type="text" name="EmisSysFix0"  id="EmisSysFix0"  class="form-control"  >
		    </div>
       </div>
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("% su costo ")%>
            <input value="<%=EmisSysPerc%>" type="text" name="EmisSysPerc0" id="EmisSysPerc0" class="form-control"  >
		    </div>
       </div>   
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Minimo &euro;")%>
            <input value="<%=EmisSysMin%>"  type="text" name="EmisSysMin0"  id="EmisSysMin0"  class="form-control"  >
		    </div>
       </div> 	   

	   <div class="col-1">
            <div class="form-group ">
		     <%xx=ShowLabel("Di cui da rilasciare in rete")%><br>
            </div> 
	   </div>
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Importo fisso &euro;")%>
            <input value="<%=EmisReteFix%>"  type="text" name="EmisReteFix0"  id="EmisReteFix0"  class="form-control"  >
		    </div>
       </div>
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("% dei diritti")%>
            <input value="<%=EmisRetePerc%>" type="text" name="EmisRetePerc0" id="EmisRetePerc0" class="form-control"  >
		    </div>
       </div>   
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Minimo &euro;")%>
            <input value="<%=EmisReteMin%>"  type="text" name="EmisReteMin0"  id="EmisReteMin0"  class="form-control"  >
		    </div>
       </div> 	   
	</div>	

    <%
	FlagAnd=""
	FlagOr =""
	if InteTipoCalc="FixAndPerc" then 
	   FlagAnd = " checked "
	else
	   InteTipoCalc = "FixOrPerc"
	   FlagOr  = " checked "
	end if 
	
	if false then 
	%>
	<div class="row">
	   <div class="col-3">
            <div class="form-group ">
		     <%xx=ShowLabel("Diritti di Intermediazione piattaforma")%><br>
             <input id="InteTipoCalc0" <%=FlagAnd%> name="InteTipoCalc0" 
		       type="radio" value = "FixAndPerc" class="big-checkbox" onclick="setTipo('and');" >
               <span class="font-weight-bold">Fisso pi&ugrave; Perc.</span>
             <input id="InteTipoCalc0" <%=FlagOr%>  name="InteTipoCalc0" 
		       type="radio" value = "FixOrPerc" class="big-checkbox" onclick="setTipo('or');">
               <span class="font-weight-bold">Fisso o Perc.</span>
	        </div>
	   </div>
	   
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Importo fisso &euro;")%>
            <input value="<%=InteSysFix%>"  type="text" name="InteSysFix0"  id="InteSysFix0"  class="form-control"  >
		    </div>
       </div>
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Perc.su costo %")%>
            <input value="<%=InteSysPerc%>" type="text" name="InteSysPerc0" id="InteSysPerc0" class="form-control"  >
		    </div>
       </div>   
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Minimo &euro;")%>
            <input value="<%=InteSysMin%>"  type="text" name="InteSysMin0"  id="InteSysMin0"  class="form-control"  >
		    </div>
       </div> 	   

	   <div class="col-1">
            <div class="form-group ">
		     <%xx=ShowLabel("Di cui da rilasciare in rete")%><br>
            </div> 
	   </div>
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Importo fisso")%>
            <input value="<%=InteReteFix%>"  type="text" name="InteReteFix0"  id="InteReteFix0"  class="form-control"  >
		    </div>
       </div>
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Perc.su interm.")%>
            <input value="<%=InteRetePerc%>" type="text" name="InteRetePerc0" id="InteRetePerc0" class="form-control"  >
		    </div>
       </div>   
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Minimo")%>
            <input value="<%=InteReteMin%>"  type="text" name="InteReteMin0"  id="InteReteMin0"  class="form-control"  >
		    </div>
       </div> 	   
	</div>	

    <%
	end if 
	
	FlagAnd=""
	FlagOr =""
	if AggiTipoCalc<>"Perc" then 
	   FlagAnd = " checked "
	else
	   FlagOr  = " checked "
	end if 
	
	%>
	<div class="row">
	   <div class="col-3">
            <div class="form-group ">
		     <%xx=ShowLabel("Diritti aggiuntivi intermediario")%><br>
             <input id="AggiTipoCalc0" <%=FlagAnd%> name="AggiTipoCalc0" 
		       type="radio" value = "Fix" class="big-checkbox">
               <span class="font-weight-bold">Importo Fisso</span>
             <input id="AggiTipoCalc0" <%=FlagOr%>  name="AggiTipoCalc0" 
		       type="radio" value = "Perc" class="big-checkbox">
               <span class="font-weight-bold">Importo Perc.</span>
	        </div>
	   </div>
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Minimo")%>
            <input value="<%=AggiMin%>"  type="text" name="AggiMin0"  id="AggiMin0"  class="form-control"  >
		    </div>
       </div>
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Massimo")%>
            <input value="<%=AggiMax%>"  type="text" name="AggiMax0"  id="AggiMax0" class="form-control"  >
		    </div>
       </div>   
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("Base")%>
            <input value="<%=AggiDef%>"  type="text" name="AggiDef0"  id="AggiDef0"  class="form-control"  >
		    </div>
       </div>
	   <%if false then %>  
	   <div class="col-1">
            <div class="form-group ">
		     <%xx=ShowLabel("Percentuale trattenuta dal sistema")%><br>
            </div> 
	   </div>
       
       <div class = "col-1">
	        <div class="form-group ">
			<%xx=ShowLabel("% tratt.")%>
            <input value="<%=AggiPercSys%>"  type="text" name="AggiPercSys0"  id="AggiPercSys0"  class="form-control"  >
		    </div>
       </div>
	   <%end if %>
	   
	   
	</div>   
   <%if EsisteProd = false then%>
   		<div class="row"><div class="mx-auto">
		<div class="bg-danger text-white">Nessun prodotto presente : non &egrave; possibile aggiornare</div>
		</div></div>
   <%end if %>
   
   <%if SoloLettura=false and EsisteProd = true then%>
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('update','0');S"%>
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
