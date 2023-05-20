<!-- funzioni per l'account -->
<!--#include virtual="/gscVirtual/modelli/FunctionAccount.asp"-->
  
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
 
  NameLoaded= NameLoaded & "UserId,TE"  
  NameLoaded= NameLoaded & ";Password,TE"  
  NameLoaded= NameLoaded & ";Cognome,TE" 
  NameLoaded= NameLoaded & ";CodiceFiscale,ML16"  
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  v_IdAccount=0
  if FirstLoad then 
	 v_IdAccount   = "0" & Session("swap_idAccount")
	 if Cdbl(v_IdAccount)=0 then 
		v_IdAccount = cdbl("0" & getValueOfDic(Pagedic,"IdAccount"))
	 end if 
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	 if PaginaReturn="" then 
		PaginaReturn = Session("swap_PaginaReturn")
	 end if 
  else
	 v_IdAccount   = "0" & getValueOfDic(Pagedic,"IdAccount")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  v_IdAccount = cdbl(v_IdAccount)
  %>
  <input type="hidden" name="IdAccount" id="IdAccount" value="<%=v_IdAccount%>">
  <%
  'inizio elaborazione pagina

  Dim DizDatabase
  Set DizDatabase = CreateObject("Scripting.Dictionary")
  
  ' preso da include virtual="/gscVirtual/model/account.asp"-->
	xx=SetDiz(DizDatabase,"IdAccount",0)
	xx=SetDiz(DizDatabase,"IdTipoAccount","")
	xx=SetDiz(DizDatabase,"UserId","")
	xx=SetDiz(DizDatabase,"PassWord","")
	xx=SetDiz(DizDatabase,"Abilitato",1)
	xx=SetDiz(DizDatabase,"Nominativo","")
	xx=SetDiz(DizDatabase,"PartitaIva","")
	xx=SetDiz(DizDatabase,"CodiceFiscale","")
	xx=SetDiz(DizDatabase,"Indirizzo1","")
	xx=SetDiz(DizDatabase,"Indirizzo2","")
	xx=SetDiz(DizDatabase,"Cap","")
	xx=SetDiz(DizDatabase,"Comune","")
	xx=SetDiz(DizDatabase,"Provincia","")
	xx=SetDiz(DizDatabase,"Settore","")
	xx=SetDiz(DizDatabase,"email1","")
	xx=SetDiz(DizDatabase,"email2","")
	xx=SetDiz(DizDatabase,"Telefono","")
	xx=SetDiz(DizDatabase,"FlagAttivo","S")
	xx=SetDiz(DizDatabase,"DescBlocco","")
	xx=SetDiz(DizDatabase,"Cognome","")
	xx=SetDiz(DizDatabase,"Nome","")
	xx=SetDiz(DizDatabase,"IdTipoUsoServizio","NESSUNO")  
  
  'recupero i dati 
  if cdbl(v_IdAccount)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  Account "
	  MySql = MySql & " Where IdAccount=" & v_IdAccount
	  
	  xx=GetInfoRecordset(DizDatabase,MySql)

  end if 
  
  if Oper=ucase("update") then 
	'response.write "update:" & v_IdAccount
	xx=SetDiz(DizDatabase,"IdAccount",v_IdAccount)
	xx=SetDiz(DizDatabase,"IdTipoAccount",Request("IdTipoAccount0")) 
	tt=trim(Request("UserId0"))
	if tt<>"" then 
	   xx=SetDiz(DizDatabase,"UserId",Request("UserId0"))
	end if 
	xx=SetDiz(DizDatabase,"PassWord",cripta(Request("PassWord0")))
	xx=SetDiz(DizDatabase,"Nominativo",trim(Request("Cognome0") & " " & Request("Nome0")))
	xx=SetDiz(DizDatabase,"PartitaIva",Request("PartitaIva0"))
	xx=SetDiz(DizDatabase,"CodiceFiscale",Request("CodiceFiscale0"))
	xx=SetDiz(DizDatabase,"Indirizzo1",Request("Indirizzo10"))
	xx=SetDiz(DizDatabase,"Indirizzo2",Request("Indirizzo20"))
	xx=SetDiz(DizDatabase,"Cap",Request("Cap0"))
	xx=SetDiz(DizDatabase,"Comune",Request("Comune0"))
	xx=SetDiz(DizDatabase,"Provincia",Request("Provincia0"))
	xx=SetDiz(DizDatabase,"Settore",Request("Settore0"))
	xx=SetDiz(DizDatabase,"email1",Request("email10"))
	xx=SetDiz(DizDatabase,"email2",Request("email20"))
	xx=SetDiz(DizDatabase,"Telefono",Request("Telefono0"))
	xx=SetDiz(DizDatabase,"Cognome",Request("Cognome0"))
	xx=SetDiz(DizDatabase,"Nome",Request("Nome0"))     
	MsgErrore=UpdateAccount()
    if cdbl(v_IdAccount)=0 then 
	   v_IdAccount = GetDiz(DizDatabase,"IdAccount")
	end if 
  end if

  'registro i dati della pagina 
  xx=setValueOfDic(Pagedic,"IdAccount"    ,v_IdAccount)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescLoaded="0"  
  %>