<%
  NomePagina="ProfiloProdottiAttiva.asp"
  titolo="Attivazione prodotti"
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
<script>
function localMod(op,id)
{
var xx;
   xx=false;
   if (op=="DEL")
      xx=true;
    
   if (op=="INS" || op=="UPD") {

      var dtf = ValoreDi("Dal" + id).trim();
	  var dtt = ValoreDi("Al"  + id).trim();
	  
	  if (op=="INS" && dtf.length==0)
	     yy=ImpostaValoreDi("Dal" + id,ValoreDi("DataDiOggi"));
	  if (op=="INS" && dtt.length==0)
	     yy=ImpostaValoreDi("Al" + id,"31/12/2099");
      yy=ImpostaValoreDi("DescLoaded",id);
      yy=ImpostaValoreDi("NameLoaded","Dal,DTO;Al,DTO");
      xx=ElaboraControlli();  
		 
   }
   
   if (xx==false)
      return false;  

   yy=AttivaFunzione(op,id); 
   
}
</script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->

<%
NameLoaded=NameLoaded & "ValidoDal,DTO;ValidoAl;DTO"

IdAzienda=1
if FirstLoad then 
   PaginaReturn  = Session("swap_PaginaReturn")
   IdAccount     = getCurrentValueFor("IdAccount")
else
   PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
   IdAccount     = getValueOfDic(Pagedic,"IdAccount")   
end if 

on error resume next 
FlagUpdRiferimento=false 

if Oper="INS" or Oper="UPD" then 
    Session("TimeStamp")=TimePage
	KK=Request("ItemToRemove")
	IdProfiloProdotto = cdbl("0" & KK)

	ValidoDal=DataStringa(Request("Dal" & KK))
	if IsNumeric(ValidoDal)=false then 
	   ValidoDal=Dtos()
	end if 
	Validoal =DataStringa(Request("Al" & KK))
	if IsNumeric(ValidoAl)=false then 
	   ValidoAl=20991231
	end if 
	if cdbl(IdProfiloProdotto)>0 then 
	   qUpd=""
       if Oper="INS" then 
	      qUpd = qUpd & " Insert into AccountProfiloProdotto "
		  qUpd = qUpd & " (IdAccount,IdProfiloProdotto,ValidoDal,ValidoAl) values "
		  qUpd = qUpd & " (" & IdAccount & "," & IdProfiloProdotto & "," & ValidoDal & "," & ValidoAl & ")"
	   else
	      qUpd = qUpd & " update AccountProfiloProdotto set "
		  qUpd = qUpd & " ValidoDal = " & ValidoDal
		  qUpd = qUpd & ",ValidoAl  = " & ValidoAl
		  qUpd = qUpd & " Where IdProfiloProdotto = " & IdProfiloProdotto
		  qUpd = qUpd & " and IdAccount = " & IdAccount
	   end if 
	   'response.write qUpd
	   connMsde.execute qUpd 
    end if 	
End if 
if Oper="DEL" then
   Session("TimeStamp")=TimePage
   KK=Request("ItemToRemove")
   IdProfiloProdotto = cdbl("0" & KK)
  
   if cdbl(IdProfiloProdotto)>0 then 
      qUpd = ""
      qUpd = qUpd & " delete from AccountProfiloProdotto  "
	  qUpd = qUpd & " Where IdProfiloProdotto=" & IdProfiloProdotto
      ConnMsde.execute qUpd 
      If Err.Number <> 0 Then 
         MsgErrore = ErroreDb(Err.description)
      End If
   END if 
End if 
  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"PaginaReturn",PaginaReturn)
  xx=setValueOfDic(Pagedic,"IdAccount"   ,IdAccount)
  xx=setCurrent(NomePagina,livelloPagina) 

  descUtente=LeggiCampo("select * from Account where IdAccount=" & IdAccount,"Nominativo")
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
				<div class="col-11"><h3>Attivazione Profilo Prodotti</b></h3>
				</div>
			</div>
			<div class="row">
				<div class="col-1"></div>			
			   <div class="col-3">
                  <div class="form-group ">
                     <%xx=ShowLabel("Utente")%>
                     <input type="text" readonly class="form-control" value="<%=DescUtente%>" >
                  </div>        
               </div>  

			</div>
			<%
			AddRow=true
			dim CampoDb(10)
			CampoDB(1)="DescProfiloProdotto"	
			ElencoOption=";0;Profilo;1"
			%>		
			<!--#include virtual="/gscVirtual/include/FiltroSearchNew.asp"-->
			
			<%
			if Firstload then 
			   flagAttivo  = ""
			   flagCessato = ""
			   flagDaAtt   = ""
			   flagTutti  = " checked "
			else 
			   flagAttivo  = ""
			   flagCessato = ""
			   flagDaAtt   = ""
			   flagTutti   = ""
			   if Request("checkAttivo")="A" then 
			   	  flagAttivo  = " checked "
			   end if 
			   if Request("checkAttivo")="C" then 
                   flagCessato = " checked "
			   end if 
			   if Request("checkAttivo")="D" then 
			      flagDaAtt   = " checked "
               end if 
			   if Request("checkAttivo")="T" then 
			      flagTutti   = " checked "
               end if 			   
			end if    
			%>
            <div class="row">
			   <div class="col-1 font-weight-bold">Stato Profilo</div>
			   <div class="col-4">
                  <div class="form-group ">
  	                   <input id="checkAttivo<%=l_Id%>" <%=FlagAttivo%> name="checkAttivo<%=l_Id%>" 
				       type="radio" value = "A" class="big-checkbox" onclick="Sottometti();">
                        <span class="font-weight-bold">Attivo</span>
  	                   <input id="checkAttivo<%=l_Id%>" <%=FlagCessato%> name="checkAttivo<%=l_Id%>" 
				       type="radio" value = "C" class="big-checkbox" onclick="Sottometti();">
                        <span class="font-weight-bold">cessato</span>
  	                   <input id="checkAttivo<%=l_Id%>" <%=FlagDaAtt%> name="checkAttivo<%=l_Id%>" 
				       type="radio" value = "D" class="big-checkbox" onclick="Sottometti();">
                        <span class="font-weight-bold">Da Attivare</span>
  	                   <input id="checkAttivo<%=l_Id%>" <%=FlagTutti%> name="checkAttivo<%=l_Id%>" 
				       type="radio" value = "T" class="big-checkbox" onclick="Sottometti();">
                        <span class="font-weight-bold">Tutti</span>						
				 </div>
			</div>
	   </div>

<%
'caricamento tabella 
if Condizione<>"" then 
	Condizione=" and " & Condizione
end if 

Set Rs = Server.CreateObject("ADODB.Recordset")

MySql = "" 
MySql = MySql & " Select B.*,A.ValidoDal,A.ValidoAl "
MySql = MySql & " From ProfiloProdotto B left join AccountProfiloProdotto A  "
MySql = MySql & " on    A.IdAccount = " & IdAccount 
MySql = MySql & " and   A.IdProfiloProdotto = B.IdProfiloProdotto"
MySql = MySql & " where B.IdTipoProfilo = 'PROFILO'"
MySql = MySql & Condizione & " order By B.DescProfiloProdotto"

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
		<th scope="col">Profilo</th>
		<th scope="col">Attivo Dal</th>
		<th scope="col">Attivo Al</th>
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
		Id=Rs("IdProfiloProdotto")
		err.clear 
		if Rs("ValidoDal")=0 then 
		   ValidoDal=""
		   ValidoAl =""
		else
		   ValidoDal=Stod(Rs("ValidoDal"))
		   ValidoAl =Stod(Rs("ValidoAl"))
        end if 
		showRow = false 
        if flagTutti="" then 
           if flagDaAtt<>"" then 
		      if ValidoDal="" then 
			     showRow = true 
              end if 
           else 
              if flagCessato<>"" then 
                 if ValidoAl < Dtos() then 
				    ShowRow = true 
			     end if  
              else 
			     if ValidoDal<=Dtos() and ValidoAl>=Dtos() then 
				    ShowRow = true 
                 end if 
              end if 
           end if 
        else
		   showRow = true 
        end if 
		if ShowRow = true then 
		%>
		
		<tr scope="col">
			<td>
			    <input class="form-control" type="text" readonly value="<%=Rs("DescProfiloProdotto")%>">
			</td>
			<td style="width: 15%;">
			<input type="text" class="form-control mydatepicker" id="Dal<%=id%>" name="Dal<%=id%>" 
			placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" value="<%=ValidoDal%>"/>
			</td>
			<td style="width: 15%;">
			<input type="text" class="form-control mydatepicker" id="Al<%=id%>" name="Al<%=id%>" 
			placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" value="<%=ValidoAl%>"/>
			</td>
            <td>			
			<%if ValidoDal="" then%>
			
				<%RiferimentoA="col-2;#;;2;plus;Inserisci;;localMod('INS','" & id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
           <%else%>				
				<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;localMod('UPD','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<%RiferimentoA="col-2;#;;2;dele;Cancella;;localMod('DEL','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
				<% 
			end if %>
			<td>
		</td>
		</tr>
		<%	
		end if 
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
