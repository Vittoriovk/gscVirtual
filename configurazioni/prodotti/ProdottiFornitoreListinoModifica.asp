<%
  NomePagina="ProdottiFornitoreListinoModifica.asp"
  titolo="Prodotto : Listino Prodotto"
  default_check_profile="SuperV,Admin,Coll"
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

function localFun(Op,Id)
{
    xx=ImpostaValoreDi("DescLoaded","0");
    xx=ElaboraControlli();
    
     if (xx==false)
       return false;
    
    var pc = GetNumberAsFloat(ValoreDi("PrezzoCompagnia0"));
    var pf = GetNumberAsFloat(ValoreDi("PrezzoFornitore0"));
    var pd = GetNumberAsFloat(ValoreDi("PrezzoDistribuzione0"));
    var pl = GetNumberAsFloat(ValoreDi("PrezzoListino0"));
    var ch = ValoreDi("checkDef0");
    var cc = ValoreDi("checkColl0");
    
    if (pf<pc) {
       alert("il prezzo fornitore deve essere maggiore o uguale al prezzo di compagnia");
       return false;
    }
    if (pd<pf && cc=="N") {
       alert("il prezzo di distribuzione deve essere maggiore o uguale al prezzo del fornitore");
       return false;
    }
    
    if (ch=="S") {
       var pdM = GetNumberAsFloat(ValoreDi("PrezzoDistribuzioneDef0")); 
       if (pd<pdM) {
          alert("il prezzo di distribuzione deve essere maggiore o uguale al prezzo di distribuzione minimo");
          return false;
       }
    }
    
    if (pl<pd) {
       alert("il prezzo di listino deve essere maggiore o uguale al prezzo di distribuzione");
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


<%
  NameLoaded= "ValidoDal,DTO;PrezzoCompagnia,FLZ;PrezzoFornitore,FLZ;PrezzoDistribuzione,FLZ;PrezzoListino,FLZ"
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdAccount          = 0
  IdProdotto         = 0
  IdAccountFornitore = 0
  
  DescFornitore=""
  DescProdotto =""
  if FirstLoad then 

     IdProdotto         = getCurrentValueFor("IdProdotto")
     IdAccount          = getCurrentValueFor("IdAccount")
     IdAccountFornitore = getCurrentValueFor("IdAccountFornitore")
     ValidoDal          = getCurrentValueFor("ValidoDal")
     PaginaReturn       = getCurrentValueFor("PaginaReturn")   
     
     IdProdotto         = cdbl("0" & IdProdotto)
     IdAccount          = cdbl("0" & IdAccount)
     IdAccountFornitore = cdbl("0" & IdAccountFornitore)
     ValidoDal          = cdbl("0" & ValidoDal)
 
     if cdbl(IdAccountFornitore)>0 then 
        Rs.CursorLocation = 3 
        Rs.Open "Select * from Fornitore where IdAccount=" & IdAccountFornitore, ConnMsde   
        DescFornitore = Rs("DescFornitore")
        Rs.close 
     end if      
     if cdbl(IdProdotto)>0 then 
        Rs.CursorLocation = 3 
        Rs.Open "Select * from Prodotto where IdProdotto=" & IdProdotto, ConnMsde   
        DescProdotto = Rs("DescProdotto")
        Rs.close 
     end if       
  else
     PaginaReturn       = getValueOfDic(Pagedic,"PaginaReturn") 
     IdProdotto         = "0" & getValueOfDic(Pagedic,"IdProdotto")
     IdAccount          = "0" & getValueOfDic(Pagedic,"IdAccount")
     IdAccountFornitore = "0" & getValueOfDic(Pagedic,"IdAccountFornitore")
     DescProdotto       = getValueOfDic(Pagedic,"DescProdotto")
     DescFornitore      = getValueOfDic(Pagedic,"DescFornitore")     
     ValidoDal          = "0" & getValueOfDic(Pagedic,"ValidoDal")

  end if 
 
  IdProdotto         = cdbl("0" & IdProdotto)
  IdAccount          = cdbl("0" & IdAccount)
  IdAccountFornitore = cdbl("0" & IdAccountFornitore)
  IdAccountRegistratore = 0 
  if IsCollaboratore() then 
     IdAccountRegistratore = Session("LoginIdAccount") 
  end if 
  'response.write IdProdotto & " " & IdAccount & " " & IdAccountFornitore
  ValidoDal          = cdbl("0" & ValidoDal)
  if Cdbl(IdAccount)=0 or cdbl(IdProdotto)=0 or Cdbl(IdAccountFornitore)=0  then 
     response.redirect virtualPath & PaginaReturn
     response.end 
  end if   
 
  if Oper=ucase("Update") then 
     ValidoDalNew        = DataStringa(Request("ValidoDal0"))
     PrezzoCompagnia     = Cdbl("0" & Request("PrezzoCompagnia0"))
     PrezzoFornitore     = Cdbl("0" & Request("PrezzoFornitore0"))
     PrezzoDistribuzione = Cdbl("0" & Request("PrezzoDistribuzione0"))
     PrezzoListino       = Cdbl("0" & Request("PrezzoListino0"))
     qUpd=""
 
     if cdbl(ValidoDal)=0 and cdbl("0" & ValidoDalNew)>0 then 
        qUpd = qUpd & " insert into AccountProdottoListino (IdAccountRegistratore,IdAccount,IdProdotto,ValidoDal,PrezzoCompagnia"
        qUpd = qUpd & ",PrezzoFornitore,PrezzoDistribuzione,PrezzoListino,IdAccountFornitore,TipoRegola)"
        qUpd = qUpd & " values("
        qUpd = qUpd & "  " & IdAccountRegistratore
        qUpd = qUpd & ", " & IdAccount
        qUpd = qUpd & ", " & IdProdotto
        qUpd = qUpd & ", " & NumForDb(ValidoDalNew) 
        qUpd = qUpd & ", " & NumForDb(PrezzoCompagnia) 
        qUpd = qUpd & ", " & NumForDb(PrezzoFornitore) 
        qUpd = qUpd & ", " & NumForDb(PrezzoDistribuzione) 
        qUpd = qUpd & ", " & NumForDb(PrezzoListino)
        qUpd = qUpd & ", " & NumForDb(IdAccountFornitore)
		qUpd = qUpd & ",'" & Session("LoginTipoUtente") & "'"
        qUpd = qUpd & " )"
     elseif cdbl("0" & ValidoDalNew)>0 then 
        qUpd = qUpd & " update AccountProdottoListino set "
        qUpd = qUpd & " ValidoDal = " & ValidoDalNew
        qUpd = qUpd & ",PrezzoCompagnia = " & numForDb(PrezzoCompagnia)
        qUpd = qUpd & ",PrezzoFornitore = " & numForDb(PrezzoFornitore)
        qUpd = qUpd & ",PrezzoDistribuzione = " & numForDb(PrezzoDistribuzione)
        qUpd = qUpd & ",PrezzoListino = " & numForDb(PrezzoListino)        
        qUpd = qUpd & " Where IdAccount = " & IdAccount 
        qUpd = qUpd & " and IdProdotto=" & IdProdotto
        qUpd = qUpd & " and IdAccountFornitore=" & IdAccountFornitore
        qUpd = qUpd & " and ValidoDal=" & ValidoDal
     end if 
     if qUpd<>"" then 
        'response.write qUpd 
        ConnMsde.execute qUpd 
        if Err.number=0 then
           response.redirect RitornaA(PaginaReturn)
           response.end 
        else
           MsgErrore=ErroreDb(err.description)
        end if 
     end if 

  End if 
  
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdProdotto"         ,IdProdotto)
  xx=setValueOfDic(Pagedic,"IdAccount"          ,IdAccount)
  xx=setValueOfDic(Pagedic,"IdAccountFornitore" ,IdAccountFornitore)
  xx=setValueOfDic(Pagedic,"DescProdotto"   ,DescProdotto)
  xx=setValueOfDic(Pagedic,"DescFornitore"  ,DescFornitore)
  xx=setValueOfDic(Pagedic,"ValidoDal"      ,ValidoDal)
  xx=setValueOfDic(Pagedic,"PaginaReturn"   ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 
  
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
            <%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;;"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <div class="col-11"><h3>Listino Prodotto</b></h3>
                </div>
            </div>
            <div class="row">
               <div class="col-1">
               </div>
               <div class="col-4 form-group ">
                  <%xx=ShowLabel("Prodotto")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescProdotto%>" >
               </div>
               <%if IsCollaboratore()=false then %>
               <div class="col-4 form-group ">
                  <%xx=ShowLabel("Fornitore")%>
                 <input type="text" readonly class="form-control input-sm" value="<%=DescFornitore%>" >
               </div>
               <%end if %>
            </div>            

<!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
<%
   set Rs = Server.CreateObject("ADODB.Recordset")
   Rs.CursorLocation = 3
   if MsgErrore="" then 

      MySql = ""
      MySql = MySql & " select * "
      MySql = MySql & " from AccountProdottoListino "
      MySql = MySql & " Where IdAccount = "           & NumForDb(Idaccount)
      MySql = MySql & " and IdAccountRegistratore = " & NumForDb(IdAccountRegistratore)
      MySql = MySql & " and IdProdotto  = "           & NumForDb(IdProdotto)
      MySql = MySql & " and ValidoDal  = "            & NumForDb(ValidoDal)
      'response.write MySql 

      Rs.Open MySql, ConnMsde

      PrezzoCompagnia     = 0 
      PrezzoFornitore     = 0 
      PrezzoDistribuzione = 0 
      PrezzoListino       = 0 

      if Rs.eof=false then 
         PrezzoCompagnia     = rs("PrezzoCompagnia")
         PrezzoFornitore     = rs("PrezzoFornitore") 
         PrezzoDistribuzione = rs("PrezzoDistribuzione")
         PrezzoListino       = rs("PrezzoListino")
      end if 

      Rs.close 
      err.clear 
   else
      ValidoDal = ValidoDalNew   
   end if 
   
   
   'per inserimento recupero prezzo compagnia/fornitore 
   TrovatoPrezzo=false 
   
   if IsAdmin() or IsCollaboratore() then 
      PrezzoCompagniaDef     = PrezzoCompagnia
      PrezzoFornitoreDef     = PrezzoFornitore
      PrezzoDistribuzioneDef = PrezzoDistribuzione
      PrezzoListinoDef       = PrezzoListino
       
      q = ""
      q = q & " select * from AccountProdottoListino "
      q = q & " Where IdProdotto = "          & IdProdotto
	  q = q & " and IdAccountRegistratore = " & NumForDb(IdAccountRegistratore)
      if IsAdmin() then 
         q = q & " and idAccount = "          & idAccountFornitore
      end if 
      if IsCollaboratore() then 
         q = q & " and idAccount = "          & Session("LoginIdAccount")
      end if 
      
      q = q & " and idAccountFornitore = " & idAccountFornitore
      q = q & " and ValidoDal <= " & Dtos()
      q = q & " order By ValidoDal desc "

      Rs.Open q, ConnMsde

      if Rs.eof=false then 
         TrovatoPrezzo=true 
         PrezzoCompagniaDef     = rs("PrezzoCompagnia")
         PrezzoFornitoreDef     = rs("PrezzoFornitore") 
         PrezzoDistribuzioneDef = rs("PrezzoDistribuzione")
         PrezzoListinoDef       = rs("PrezzoListino")
         if ValidoDal=0 then
            PrezzoCompagnia     = rs("PrezzoCompagnia")
            PrezzoFornitore     = rs("PrezzoFornitore") 
            PrezzoDistribuzione = rs("PrezzoDistribuzione")
            PrezzoListino       = rs("PrezzoListino")
         end if 
         Rs.close
         err.clear 
     end if 
   else 
     TrovatoPrezzo = true 
   end if 

   if ValidoDal = 0 then 
      ValidoDal = Stod(Dtos())
   else
      ValidoDal = STod(ValidoDal)
   end if    
DescLoaded=""
NumCols = numC + 1
NumRec  = 0
ShowNew    = true
ShowUpdate = false
MsgNoData  = ""
l_Id = "0"
%>
<br>
    <div class="row">
       <div class="col-1">
          <p class="font-weight-bold"></p>
       </div>
       
       <div class="col-2">
          <p class="font-weight-bold">Valido Dal</p>
       </div>

       <div class ="col-2"> 
          <input type="text"  name="ValidoDal<%=l_Id%>" id="ValidoDal<%=l_Id%>" value="<%=ValidoDal%>" 
                 class="form-control mydatepicker " placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" >
       </div>

    </div>
    <%if IsCollaboratore()=false then %>
    <input type="hidden" name="checkColl0" id="checkColl0" value = "N" >
    <div class="row">
       <div class="col-1">
          <p class="font-weight-bold"></p>
       </div>
       
       <div class="col-2">
          <p class="font-weight-bold">Prezzo Compagnia &euro;</p>
       </div>
       <%
       LockCompagnia = ""
       LockFornitore = ""       
       if IsAdmin() or IsCollaboratore() then 
          LockCompagnia = " readonly "
          LockFornitore = " readonly "
          if IsCollaboratore() then 
             
          end if 
       end if 
       
       %>
       <div class ="col-1"> 
       <input id="PrezzoCompagnia<%=l_Id%>" name="PrezzoCompagnia<%=l_Id%>" <%=LockCompagnia%> type="text" value = "<%=PrezzoCompagnia%>" class="form-control input-sm" >
       </div>
    </div>
    <%else%>
    <input type="hidden" name="checkColl0" id="checkColl0" value = "S" >
    <input id="PrezzoCompagnia<%=l_Id%>" name="PrezzoCompagnia<%=l_Id%>" type="hidden" value = "<%=PrezzoCompagnia%>">
    <%end if %>
    
    <%if IsCollaboratore()=false then %>
    <div class="row">
       <div class="col-1">
          <p class="font-weight-bold"></p>
       </div>
       
       <div class="col-2">
          <p class="font-weight-bold">Prezzo fornitore &euro;</p>
       </div>

       <div class ="col-1"> 
       <input id="PrezzoFornitore<%=l_Id%>" name="PrezzoFornitore<%=l_Id%>" <%=LockFornitore%> type="text" value = "<%=PrezzoFornitore%>" class="form-control input-sm" >
       </div>
    </div>
    <%else%>
    <input id="PrezzoFornitore<%=l_Id%>" name="PrezzoFornitore<%=l_Id%>" type="hidden" value = "<%=PrezzoDistribuzioneDef%>">
    <%end if %>    
    <div class="row">
       <div class="col-1">
          <p class="font-weight-bold"></p>
       </div>
       
       <div class="col-2">
          <p class="font-weight-bold">Prezzo Distribuzione &euro;</p>
       </div>

       <div class ="col-1">
       <input id="PrezzoDistribuzione<%=l_Id%>" name="PrezzoDistribuzione<%=l_Id%>" type="text" value = "<%=PrezzoDistribuzione%>" class="form-control input-sm" >
       </div>
       <input id="PrezzoDistribuzioneDef<%=l_Id%>" name="PrezzoDistribuzioneDef<%=l_Id%>" type="hidden" 
       value = "<%=PrezzoDistribuzioneDef%>">
       
       <%if IsAdmin() or IsCollaboratore() then %>
           <input type="hidden" name="checkDef0" id="checkDef0" value = "S" >
              <div class ="col-2">
               <p class="font-weight-bold">Prezzo Minimo <%=InsertPoint(PrezzoDistribuzioneDef,2)%> &euro;</p>
              </div>
       <%else%>
           <input type="hidden" name="checkDef0" id="checkDef0" value = "N" >
       <%end if %>

    </div>
    <div class="row">
       <div class="col-1">
          <p class="font-weight-bold"></p>
       </div>
       
       <div class="col-2">
          <p class="font-weight-bold">Prezzo Listino &euro;</p>
       </div>

       <div class ="col-1">
       <input id="PrezzoListino<%=l_Id%>" name="PrezzoListino<%=l_Id%>" type="text" value = "<%=PrezzoListino%>" class="form-control input-sm" >
       </div>
       <input id="PrezzoListinoDef<%=l_Id%>" name="PrezzoListinoDef<%=l_Id%>" type="hidden" value = "<%=PrezzoListinoDef%>" >       
       <%if IsAdmin() or IsCollaboratore() then %>
              <div class ="col-2">
               <p class="font-weight-bold">Prezzo Consigliato <%=InsertPoint(PrezzoListinoDef,2)%> &euro;</p>
              </div>

       <%end if %>

    </div>
        <%if TrovatoPrezzo = false then 
             MsgInfo="Listino non presente : contattare supervisore"
         %>
            <!--#include virtual="/gscVirtual/include/ShowInfoDivRow.asp"-->
        <%else%>
        <div class="row">
            <div class="mx-auto">
               <%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
               <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
             </div>
        </div>
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
