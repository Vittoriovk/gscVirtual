<%
  NomePagina="ProdottiAccount.asp"
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
   PaginaReturn   = getCurrentValueFor("PaginaReturn")
   IdAccountPadre = getCurrentValueFor("IdAccountPadre")
   IdAccount      = getCurrentValueFor("IdAccount")
   DescAccount    = getCurrentValueFor("DescAccount")   
else
   PaginaReturn   = getValueOfDic(Pagedic,"PaginaReturn")
   IdAccountPadre = getValueOfDic(Pagedic,"IdAccountPadre")
   IdAccount      = getValueOfDic(Pagedic,"IdAccount")
   DescAccount    = getValueOfDic(Pagedic,"DescAccount")
end if 

on error resume next 
if CheckTimePageLoad()=false then
   oper=""
end if 
IdAccountPadre = cdbl("0" & IdAccountPadre)
if oper="INS" then 
   Session("TimeStamp")=TimePage
   IdProdotto= Request("IdProdotto0")
   ArDati=split(IdProdotto,"_")
   idP = cdbl("0" & ArDati(0))
   idF = 0
   ValidoDal=DataStringa(Request("Dal0"))
   if IsNumeric(ValidoDal)=false then 
      ValidoDal=Dtos()
   end if 
   Validoal =DataStringa(Request("Al0"))
   if IsNumeric(ValidoAl)=false then 
      ValidoAl=20991231
   end if 
   
   if Cdbl(IdP)>0 then 
      q = ""
      q = q & " insert into AccountProdotto (IdAccount,IdProdotto,MailDocumentazione,IdAccountFornitore,ValidoDal,ValidoAl) values ("
      q = q & " " & NumForDb(IdAccount)
      q = q & "," & NumForDb(IdP)
      q = q & ",''"
      q = q & "," & NumForDb(IdF)	  
	  q = q & "," & NumForDb(ValidoDal)	  
	  q = q & "," & NumForDb(ValidoAl)	  
      q = q & " )"
      connMsde.execute q 
   end if 
end if 
if Oper="UPD" then 
   Session("TimeStamp")=TimePage

   IdProdotto= Request("ItemToRemove")
   ArDati=split(IdProdotto,"_")
   idP = cdbl("0" & ArDati(0))
   
   ValidoDal=DataStringa(Request("Dal" & IdProdotto))
   if IsNumeric(ValidoDal)=false then 
      ValidoDal=Dtos()
   end if 
   Validoal =DataStringa(Request("Al" & IdProdotto))
   if IsNumeric(ValidoAl)=false then 
      ValidoAl=20991231
   end if 
   if cdbl(idP)>0 then 
      qUpd=""
      qUpd = qUpd & " update AccountProdotto set "
      qUpd = qUpd & " ValidoDal = " & ValidoDal
      qUpd = qUpd & ",ValidoAl  = " & ValidoAl
      qUpd = qUpd & " Where IdProdotto = " & idP
	  qUpd = qUpd & " and IdAccount = " & IdAccount 
	  'response.write qUpd 
      connMsde.execute qUpd 
    end if 	
End if 
if oper="CALL_DEL" then 
   Session("TimeStamp")=TimePage
   IdProdotto= Request("ItemToRemove")
   ArDati=split(IdProdotto,"_")
   idP = cdbl("0" & ArDati(0))
   if Cdbl(IdP)>0 then    
      q = ""
      q = q & " delete from AccountProdotto"
      q = q & " where IdAccount = " & NumForDb(IdAccount)
      q = q & " and IdProdotto = " & NumForDb(IdP)
      connMsde.execute q 
   end if 
   
end if 

  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"PaginaReturn"  ,PaginaReturn)
  xx=setValueOfDic(Pagedic,"IdAccount"     ,IdAccount)
  xx=setValueOfDic(Pagedic,"IdAccountPadre", IdAccountPadre)
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
                <div class="col-11"><h3>Attivazione Prodotti</b></h3>
                </div>
            </div>
            <div class="row">
               <div class="col-1">
               </div>
               <div class="col-4 form-group ">
                  <%
				  if ucase(IdTipoAccount)="CLIE" then 
				     descTipoAccount="Cliente"
                  elseif ucase(IdTipoAccount)="BACKO" then 
				     descTipoAccount="Utente Back Office"
				  else 
				     descTipoAccount="Collaboratore"
                  end if   
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
			   <div class="col-1 font-weight-bold">Servizio</div>
			   <div class="col-4">
			   <div class="form-group ">
				     <% 
					 IdAnagServizio=Request("IdAnagServizio0")
					 if IdAnagServizio="-1" then
					    IdAnagServizio=""
					 end if 
					 stdClass="class='form-control form-control-sm'"
					 q = ""
	                 q = q & " Select * from AnagServizio "
	                 q = q & " order by DescAnagServizio  "
                     response.write ListaDbChangeCompleta(q,"IdAnagServizio0",IdAnagServizio ,"IdAnagServizio","DescAnagServizio" ,1,"Sottometti();","","","","",stdClass)
                   %>
                  </div>		
			   
			   
			   </div>

			   <div class="col-1 font-weight-bold">Stato Prodotto</div>
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

'associazioni presenti 
MySql = "" 
MySql = MySql & " Select a.*,B.FlagPrezzoFisso,B.IdAnagServizio,B.DescProdotto"
MySql = MySql & " From AccountProdotto a, Prodotto B "
MySql = MySql & " Where A.IdAccount = " & IdAccount
MySql = MySql & " and A.IdProdotto = B.IdProdotto "
if IdAnagServizio<>"" then 
   MySql = MySql & " and b.IdAnagServizio = '" & apici(IdAnagServizio) & "'"
end if 

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

<div class="table-responsive"><table class="table"><tbody>
<thead>
    <tr>
        <th scope="col">Prodotto</th>
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
        IdP=Rs("IdProdotto") & "_0"
        DescLoaded=DescLoaded & Id & ";"
		 
		if Rs("ValidoDal")=0 then 
		   ValidoDal=""
		   ValidoAl =""
		else
		   ValidoDal=Stod(Rs("ValidoDal"))
		   ValidoAl =Stod(Rs("ValidoAl"))
        end if 		
		flagProdCess=false
		
		'recupero la Data Di Scadenza 
		q = ""
		q = q & " select max(ValidoAl) as ValDalProd from ProdottoAttivo "
		q = q & " Where IdProdotto=" & Rs("IdProdotto")
		
		DataScad = leggiCampo(q,"ValDalProd")
		'response.write DataScad
		if Cdbl(DataScad)<cdbl(Dtos()) then 
		   flagProdCess=true 
		end if 
        %>
        
        <tr scope="col">
            <td>
                <input class="form-control" type="text" readonly value="<%=Rs("DescProdotto")%>">
            </td>

			<%if flagProdCess=false then %>
			<td style="width: 15%;">
			<input type="text" class="form-control mydatepicker" id="Dal<%=idP%>" name="Dal<%=idP%>" 
			placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" value="<%=ValidoDal%>"/>
			</td>
			<td style="width: 15%;">
			<input type="text" class="form-control mydatepicker" id="Al<%=idP%>" name="Al<%=idP%>" 
			placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" value="<%=ValidoAl%>"/>
			</td>
			<%else%>
			<td style="width: 15%;">
			<input type="text" class="form-control" readonly value="<%=ValidoDal%>"/>
			</td>
			<td style="width: 15%;">
			<input type="text" class="form-control" readonly value="<%=ValidoAl%>"/>
			</td>
			
			<%end if %>
            <td>
			    <%if flagProdCess=false then %>
				<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;localMod('UPD','" & IdP & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->				
                <%RiferimentoA="col-2;#;;2;dele;Cancella;;AttivaFunzione('CALL_DEL','" & IdP & "');N"%>
                <!--#include virtual="/gscVirtual/include/Anchor.asp"-->  
                <%else%>
				   Prodotto non disponibile 
				<%end if %>
            <td>
        </td>
        </tr>
        <%    
        rs.MoveNext
    Loop
end if 
rs.close

'lista dei prodotti associabili
   
   q = ""
   q = q & " select distinct concat(A.IdProdotto,'_0') as codice"
   q = q & ",B.DescProdotto as valore "
   
   q = q & " From ProdottoAttivo A , Prodotto B "
   q = q & " Where A.IdProdotto = B.IdProdotto "
   if IdAnagServizio<>"" then 
      q = q & " and b.IdAnagServizio = '" & apici(IdAnagServizio) & "'"
   end if 
   q = q & " and   a.ValidoDal <= " & Dtos()
   q = q & " and   a.ValidoAl  >= " & Dtos()
   
   'escludo quelli gia associati 
   qe = ""
   qe = qe & " select 'X' from AccountProdotto tt "
   qe = qe & " where tt.IdAccount=" & IdAccount 
   qe = qe & " and tt.idProdotto = A.IdProdotto "
      
   q = q & " and not Exists (" & qe &  ")"
   
   'per i collaboratori devo prendere solo quelli associati 
   if IsCollaboratore() = true then 
      qe = ""
      qe = qe & " select 'X' from AccountProdotto tt "
      qe = qe & " where tt.IdAccount=" & IdAccountPadre 
      qe = qe & " and tt.idProdotto = A.IdProdotto "
     
      q = q & " and Exists (" & qe &  ")"
   end if 
   if flagTutti<>"" or flagDaAtt<>"" then 
      tt = leggiCampo(q,"codice")
   else
      tt = ""
   end if   
   'response.write q 
   if tt<>"" then 
        if IsCollaboratore()=false then
		   colsp="colspan=2"
		else
		   colsp=""
		end if 
        %>

        <tr scope="col">
            <td <%=colsp%>>
            <%
            stdClass="class='form-control form-control-sm'" 
            response.write ListaDbChangeCompleta(q,"IdProdotto0",IdProdotto ,"codice","Valore" ,0,"","","","","",stdClass)
      
           %>

            </td>
			<td style="width: 15%;">
			<input type="text" class="form-control mydatepicker" id="Dal0" name="Dal0" 
			placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" value="<%=ValidoDal%>"/>
			</td>
			<td style="width: 15%;">
			<input type="text" class="form-control mydatepicker" id="Al0" name="Al0" 
			placeholder="gg/mm/aaaa" title="formato : gg/mm/aaaa" value="<%=ValidoAl%>"/>
			</td>				
            <td>

                
                <%RiferimentoA="col-2;#;;2;inse;Aggiungi;;localMod('INS','0');N"%>
                <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
            <td>
        </td>
        </tr>
    <%end if %>
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
