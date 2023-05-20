<%
  NomePagina="CertificazioneMatrice.asp"
  titolo="Menu Supervisor - Gestione Matrice Certificazione"
  'forzo il controllo al profilo 
  default_check_profile="SUPERV"  
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
  
  NameLoaded= ""
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  
  IdCertificazione = 0
  if FirstLoad then 
     IdCertificazione   = getCurrentValueFor("IdCertificazione")
     DescCertificazione = getCurrentValueFor("DescCertificazione")
     OperTabella        = getCurrentValueFor("OperTabella")
     PaginaReturn       = getCurrentValueFor("PaginaReturn") 
  else
     IdCertificazione   = getValueOfDic(Pagedic,"IdCertificazione")
     DescCertificazione = getValueOfDic(Pagedic,"DescCertificazione")
     OperTabella        = getValueOfDic(Pagedic,"OperTabella")
     PaginaReturn       = getValueOfDic(Pagedic,"PaginaReturn")
  end if 

  IdCertificazione = cdbl(IdCertificazione)
  if DescCertificazione="" then 
     MySql = ""
     MySql = MySql & " Select * From  Certificazione "
     MySql = MySql & " Where IdCertificazione=" & IdCertificazione  
     DescCertificazione=LeggiCampo(MySql,"DescBreveCertificazione")
  end if 

  IdCertificazione = 0
  
 
   if Oper=ucase("CALL_INS") then 
     kk=0
     ElCer=Request("ElCer")
	 datiE=split(elCer,";")
	 CertificazioneCompatibile="|"
	 for j=lBound(datiE) to ubound(datiE)-1
	    tt=trim(datiE(j))
		if tt<>"" and Request("checkComp" & tt & "_" & kk)="S" then 
		   CertificazioneCompatibile=CertificazioneCompatibile & tt & "|"
		end if 
	 next 
     Session("TimeStamp")=TimePage
     KK="0"
     MyQ = "" 
     MyQ = MyQ & " INSERT INTO CertificazioneMatrice (IdCertificazione,CertificazioneCompatibile) " 
     MyQ = MyQ & " values (" & IdCertificazione & ",'" &  apici(CertificazioneCompatibile) & "')" 


     ConnMsde.execute MyQ 
     If Err.Number <> 0 Then 
         MsgErrore = ErroreDb(Err.description)
      End If
   end if 
  
  if Oper="CALL_UPD" then 
     kk=Cdbl("0" & Request("ItemToRemove"))
	 if kk>0 then 
		 ElCer=Request("ElCer")
		 datiE=split(elCer,";")
		 CertificazioneCompatibile="|"
		 for j=lBound(datiE) to ubound(datiE)-1
			tt=trim(datiE(j))
			if tt<>"" and Request("checkComp" & tt & "_" & kk)="S" then 
			   CertificazioneCompatibile=CertificazioneCompatibile & tt & "|"
			end if 
		 next
		 if CertificazioneCompatibile<>"|" then 
			qUpd="update CertificazioneMatrice set CertificazioneCompatibile='" & CertificazioneCompatibile & "' Where IdCertificazioneMatrice= " & kk
			'response.write qUpd 
			ConnMsde.execute qUpd
		 else
			ConnMsde.execute "Delete from CertificazioneMatrice Where IdCertificazioneMatrice= " & kk
		end if 
    end if 
    If Err.Number <> 0 Then 
        MsgErrore = ErroreDb(Err.description)
    End If    
  end if

  if Oper=ucase("update") and OperTabella="CALL_DEL" and Cdbl(IdCertificazione)>0 then 
     MsgErrore = VerificaDel("Certificazione",IdCertificazione) 
     if MsgErrore = "" then   
        MyQ = "" 
        MyQ = MyQ & " Delete from Certificazione "
        MyQ = MyQ & " Where IdCertificazione = " & IdCertificazione

        ConnMsde.execute MyQ 
        If Err.Number <> 0 Then 
            MsgErrore = ErroreDb(Err.description)
        else 
           response.redirect virtualpath & PaginaReturn
        End If    
    end if 
  end if  


  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdCertificazione"  ,IdCertificazione)
  xx=setValueOfDic(Pagedic,"DescCertificazione",DescCertificazione)
  xx=setValueOfDic(Pagedic,"OperTabella"       ,OperTabella)
  xx=setValueOfDic(Pagedic,"PaginaReturn"      ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescLoaded=""  
  
  %>

<% 
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
            <%RiferimentoA="col-1 text-center;" & VirtualPath & PaginaReturn & ";;2;prev;Indietro;;"%>
            <!--#include virtual="/gscVirtual/include/Anchor.asp"-->            
                <div class="col-11"><h3>Gestione Compatibilità Certificazioni</b></h3>
                </div>
            </div>

            <div class="table-responsive"><table class="table"><tbody>
            <thead>
	        <tr>
            <%
			MaxIdx=0
			Dim NumCer(100)
			Idx=0
			ElCer=""
            sCer = "Select * from Certificazione Where IdCertificazione<>0 order By IdCertificazione" 
			'response.write sCer 
            Rs.CursorLocation = 3 
            Rs.Open sCer, ConnMsde
			do while not Rs.eof
			   MaxIdx=MaxIdx+1
			   ElCer=ElCer & Rs("IdCertificazione") & ";"
               NumCer(MaxIdx)=Rs("IdCertificazione")			   
			   'response.write MaxIdx & "£" & NumCer(MaxIdx) & "--"
			   response.write "<th scope='col'>" & Rs("DescBreveCertificazione") & "</th>"
               Rs.movenext
			loop
			rs.close 
			%>
			<input type="hidden" name="ElCer" id="ElCer" value="<%=ElCer%>">
			<th>Azioni</th>
	        </tr>
            </thead>
			
			<%
			'leggo i dettagli
            sCer = "Select * from CertificazioneMatrice Where IdCertificazione=" & idCertificazione &" order By IdCertificazioneMatrice " 
            Rs.CursorLocation = 3 
            Rs.Open sCer, ConnMsde
			do while not Rs.eof
			   id=Rs("IdCertificazioneMatrice")
			   'response.write "<tr>"
			   for idx=1 to MaxIdx 
			      iC=NumCer(idx)
				  'response.write Rs("CertificazioneCompatibile") 
				  kk=iC & "_" & Rs("IdCertificazioneMatrice")
				  selezionato="" 
				  'response.write "aa" & idx & " " & iC
				  if instr(Rs("CertificazioneCompatibile"),"|" & iC & "|")>0 then 
				     selezionato=" checked "
				  end if 
				%>
                  <td>
				     <input type="checkbox" id="checkComp<%=kk%>" <%=selezionato%> name="checkComp<%=kk%>" 
				     value = "S" class="big-checkbox" >
			      </td>				 
				<% next %>
				<td>
				<%RiferimentoA="col-2;#;;2;upda;Aggiorna;;AttivaFunzione('CALL_UPD','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			   
 			   </td>
			   </tr>
			   <%
               Rs.movenext
			loop
			rs.close 
			
			'carico occorrenza da inserire 
			%>
			<tr>
			<% for idx=1 to MaxIdx 
			     iC=NumCer(idx)
				 kk=iC & "_0"
				 selezionato=""
			%>
			<td>
				<input type="checkbox" id="checkComp<%=kk%>" <%=selezionato%> name="checkComp<%=kk%>" 
				 value = "S" class="medium-checkbox" >
			</td>
			<% next %>
			<TD>
				<%RiferimentoA="col-2;#;;2;inse;Inserisci;;AttivaFunzione('CALL_INS','" & Id & "');N"%>
				<!--#include virtual="/gscVirtual/include/Anchor.asp"-->
			</td>
			</tr>
           

            </tbody></table></div> <!-- table responsive fluid -->            
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->        

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
