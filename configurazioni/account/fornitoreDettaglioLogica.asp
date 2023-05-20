<!-- funzioni per l'Fornitore -->
<!--#include virtual="/gscVirtual/modelli/FunctionFornitore.asp"-->
  
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
function localDelete()
{
var xx;
    xx=false;
	if (confirm('Stai per rimuovere un fornitore, sei sicuro ?')) {
	} else {
		return false;
	}
		
	ImpostaValoreDi("Oper","delete");
	document.Fdati.submit(); 
}
</script>

<%
  NameLoaded= ""
  NameLoaded= NameLoaded & "DescCognome,TE" 
  NameLoaded= NameLoaded & ";IdTipoDitta,LI"  
  NameLoaded= NameLoaded & ";IdTipoSocieta,LI"  
  NameLoaded= NameLoaded & ";IdTipoMandato,LI"
  NameLoaded= NameLoaded & ";IdTipoIncasso,LI"
 
  FirstLoad=(Request("CallingPage")<>NomePagina)
  IdFornitore=0
  IdAccount  =0
  if FirstLoad then 
	 IdFornitore   = cdbl("0" & Session("swap_idFornitore"))
	 if Cdbl(IdFornitore)=0 then 
		IdFornitore = cdbl("0" & getValueOfDic(Pagedic,"IdFornitore"))
	 end if 
	 OperTabella   = Session("swap_OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn") 
	 if PaginaReturn="" then 
		PaginaReturn = Session("swap_PaginaReturn")
	 end if 
  else
	 IdFornitore   = cdbl("0" & getValueOfDic(Pagedic,"IdFornitore"))
	 IdAccount     = cdbl("0" & getValueOfDic(Pagedic,"IdAccount"))
	 OperTabella   = getValueOfDic(Pagedic,"OperTabella")
	 PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
  end if 
  IdFornitore = cdbl(IdFornitore)
  'inizio elaborazione pagina

   Dim DizDatabase
   Set DizDatabase = CreateObject("Scripting.Dictionary")
  
   xx=SetDiz(DizDatabase,"IdFornitore",0)
   xx=SetDiz(DizDatabase,"IdAccount",0)
   xx=SetDiz(DizDatabase,"DescFornitore","")
   xx=SetDiz(DizDatabase,"DescCognome","")
   xx=SetDiz(DizDatabase,"DescNome","")
   xx=SetDiz(DizDatabase,"IdTipoDitta","")
   xx=SetDiz(DizDatabase,"IdTipoSocieta","")
   xx=SetDiz(DizDatabase,"IdTipoMandato","")
   xx=SetDiz(DizDatabase,"IdTipoIncasso","")
   xx=SetDiz(DizDatabase,"CodiceFiscale","")
   xx=SetDiz(DizDatabase,"PartitaIva","")
   xx=SetDiz(DizDatabase,"DescRuolo","")
   
  
  'recupero i dati 
  if cdbl(IdFornitore)>0 then
	  MySql = ""
	  MySql = MySql & " Select * From  Fornitore "
	  MySql = MySql & " Where IdFornitore=" & IdFornitore
	  xx=GetInfoRecordset(DizDatabase,MySql)
	  IdAccount=Cdbl("0" & GetDiz(DizDatabase,"IdAccount"))
  elseif FirstLoad=false then 
	  xx=SetDiz(DizDatabase,"DescFornitore",trim(Request("DescCognome0") & " " & Request("DescNome0")))
	  xx=SetDiz(DizDatabase,"DescCognome"  ,Request("DescCognome0"))
	  xx=SetDiz(DizDatabase,"DescNome"     ,Request("DescNome0"))     
	  xx=SetDiz(DizDatabase,"IdTipoDitta"  ,Request("IdTipoDitta0")) 
	  xx=SetDiz(DizDatabase,"IdTipoSocieta",Request("IdTipoSocieta0")) 
	  xx=SetDiz(DizDatabase,"IdTipoMandato",Request("IdTipoMandato0")) 
	  xx=SetDiz(DizDatabase,"IdTipoIncasso",Request("IdTipoIncasso0")) 
	  xx=SetDiz(DizDatabase,"CodiceFiscale",Request("CodiceFiscale0")) 
      xx=SetDiz(DizDatabase,"PartitaIva"   ,Request("PartitaIva0")) 
	  IdSezioneRui      = request("IdSezioneRui0")
	  NumeroRui         = request("NumeroRui0")
	  DataIscrizioneRui = request("DataIscrizioneRui0")
	  if IdSezioneRui="-1" then 
		 xx=SetDiz(DizDatabase,"IdSezioneRui","") 
		 xx=SetDiz(DizDatabase,"NumeroRui"   ,"") 
         xx=SetDiz(DizDatabase,"DataIscrizioneRui",0) 	
	else
		xx=SetDiz(DizDatabase,"IdSezioneRui",IdSezioneRui) 
		xx=SetDiz(DizDatabase,"NumeroRui"   ,NumeroRui) 
		xx=SetDiz(DizDatabase,"DataIscrizioneRui",DataStringa(DataIscrizioneRui)) 	
	end if 
	xx=SetDiz(DizDatabase,"DescRuolo"   ,Request("DescRuolo0")) 
  end if 

  if Oper=ucase("delete") and cdbl(IdAccount)>0 and cdbl(idfornitore)>0 then
  	 MsgErrore=VerificaDel("FORNITORE",IdAccount)
	 if MsgErrore="" then
	    MsgErrore=VerificaDel("FORNITORE_1",IdFornitore)
	    if MsgErrore="" then 
		   connMsde.execute "delete from Account where IdAccount=" & IdAccount
		   connMsde.execute "delete from AccountCompagnia where IdAccount=" & IdAccount
		   connMsde.execute "delete from AccountContatto where IdAccount=" & IdAccount
		   connMsde.execute "delete from AccountDocumento where IdAccount=" & IdAccount
		   connMsde.execute "delete from AccountModPag where IdAccount=" & IdAccount
		   connMsde.execute "delete from AccountMovEco where IdAccount=" & IdAccount
           connMsde.execute "delete from AccountProdottoDatoTecn where IdAccount=" & IdAccount
           connMsde.execute "delete from AccountProdottoDocAff where IdAccount=" & IdAccount
           connMsde.execute "delete from AccountProdottoFascia where IdAccount=" & IdAccount
           connMsde.execute "delete from AccountProdottoFirma where IdAccount=" & IdAccount
           connMsde.execute "delete from AccountProfiloProdotto where IdAccount=" & IdAccount   
		   connMsde.execute "delete from AccountTipoParametro where IdAccount=" & IdAccount   
		   connMsde.execute "delete from AccountSede where IdAccount=" & IdAccount   
		   connMsde.execute "delete from Fornitore where IdAccount=" & IdAccount   
		   connMsde.execute "delete from Utente where IdAccount=" & IdAccount   
   		   connMsde.execute "delete from AccountProdotto where IdAccountFornitore=" & IdAccount   
   		   connMsde.execute "delete from AccountProdottoListino where IdAccountFornitore=" & IdAccount   
   		   connMsde.execute "delete from ProdottoSessione where IdAccountFornitore=" & IdAccount   
           response.redirect PaginaReturn
	    end if 
	 end if 
  end if 
  if Oper=ucase("AddContattoAccount") then 
	Session("swap_idAccount")         = IdAccount
	Session("swap_idContattoAccount") = 0  
	Session("swap_PaginaReturn")      = VirtualPath & "/configurazioni/account/fornitoriDettaglio.asp" 
	response.redirect VirtualPath &   "/configurazioni/contatti/contatto.asp"
  end if   
  if Oper=ucase("ModContattoAccount") then 
	Session("swap_idAccount")         = IdAccount
	Session("swap_idContattoAccount") = Request("ItemToRemove")  
	Session("swap_PaginaReturn")      = VirtualPath & "/configurazioni/account/fornitoriDettaglio.asp" 
	response.redirect VirtualPath &   "/configurazioni/contatti/contatto.asp"
  end if    
   
  'sono in inserimento : creo un account fittizio 
  if cdbl(IdAccount)=0 and OperTabella="CALL_INS" then 
     IdAccount=GetTempAccount()
  end if 
  
  'inserisco il fornitore 
  if Oper=ucase("update") and OperTabella="CALL_INS" then 
  
    Session("TimeStamp")=TimePage
	KK="0"
	DescCogn = Request("DescCognome" & KK)
	DescNome = Request("DescNome"    & KK)
	DescIn   = trim(DescCogn & " " & DescNome)
	idTipo=Request("IdTipo" & KK)
	
	if Cdbl(IdAccount)>0 then 
		MyQ = "" 
		MyQ = MyQ & " Insert into Fornitore ("
		MyQ = MyQ & " IdAccount,DescCognome,DescNome,DescFornitore,IdTipoDitta"
		MyQ = MyQ & ") values ("			
		MyQ = MyQ & "  " & IdAccount
		MyQ = MyQ & ",'" & Apici(DescCogn) & "'"
		MyQ = MyQ & ",'" & Apici(DescNome) & "'"
		MyQ = MyQ & ",'" & Apici(DescIn) & "'"
		MyQ = MyQ & ",'" & Apici(IdTipo) & "'"
		MyQ = MyQ & ")"

		ConnMsde.execute MyQ 
		If Err.Number <> 0 Then 
			MsgErrore = ErroreDb(Err.description)
		else
		    IdFornitore = LeggiCampo("Select * from Fornitore Where IdAccount=" & IdAccount,"IdFornitore")
			'aggiorno account 
	        MyQ = "" 
	        MyQ = MyQ & " UPDATE Account "
			MyQ = MyQ & " Set IdAzienda = 1"
			MyQ = MyQ & " ,IdTipoAccount = 'Forn'"
			MyQ = MyQ & " ,UserId='Forn" & IdAccount & "'"
			MyQ = MyQ & " ,Password=''"
			MyQ = MyQ & " ,Nominativo='" & apici(DescIn) & "'"
			MyQ = MyQ & " ,abilitato = 1 " 
			MyQ = MyQ & " Where IdAccount=" & IdAccount
	        ConnMsde.execute MyQ
			DescIn=""
			OperTabella="CALL_UPD"
		End If
	else
		MsgErrore="Errore Interno : " & Err.Description
	end if   
  end if 
  
  if Oper=ucase("update") and OperTabella="CALL_UPD" and Cdbl(IdFornitore)>0 then 
	xx=SetDiz(DizDatabase,"IdFornitore",IdFornitore)
	xx=SetDiz(DizDatabase,"DescFornitore",trim(Request("DescCognome0") & " " & Request("DescNome0")))
	xx=SetDiz(DizDatabase,"DescCognome"  ,Request("DescCognome0"))
	xx=SetDiz(DizDatabase,"DescNome"     ,Request("DescNome0"))     
	xx=SetDiz(DizDatabase,"IdTipoDitta"  ,Request("IdTipoDitta0")) 
	xx=SetDiz(DizDatabase,"IdTipoSocieta",Request("IdTipoSocieta0")) 
	xx=SetDiz(DizDatabase,"IdTipoMandato",Request("IdTipoMandato0")) 
	xx=SetDiz(DizDatabase,"IdTipoIncasso",Request("IdTipoIncasso0")) 
	xx=SetDiz(DizDatabase,"CodiceFiscale",Request("CodiceFiscale0")) 
    xx=SetDiz(DizDatabase,"PartitaIva"   ,Request("PartitaIva0")) 
	IdSezioneRui      = request("IdSezioneRui0")
	NumeroRui         = request("NumeroRui0")
	DataIscrizioneRui = request("DataIscrizioneRui0")
	if IdSezioneRui="-1" then 
		xx=SetDiz(DizDatabase,"IdSezioneRui","") 
		xx=SetDiz(DizDatabase,"NumeroRui"   ,"") 
		xx=SetDiz(DizDatabase,"DataIscrizioneRui",0) 	
	else
		xx=SetDiz(DizDatabase,"IdSezioneRui",IdSezioneRui) 
		xx=SetDiz(DizDatabase,"NumeroRui"   ,NumeroRui) 
		xx=SetDiz(DizDatabase,"DataIscrizioneRui",DataStringa(DataIscrizioneRui)) 	
	end if 
	xx=SetDiz(DizDatabase,"DescRuolo"   ,Request("DescRuolo0")) 
	
	MsgErrore=UpdateFornitore()
    if cdbl(IdFornitore)=0 then 
	   IdFornitore = GetDiz(DizDatabase,"IdFornitore")
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
  xx=setValueOfDic(Pagedic,"IdFornitore"  ,IdFornitore)
  xx=setValueOfDic(Pagedic,"IdAccount"    ,IdAccount)
  xx=setValueOfDic(Pagedic,"OperTabella"  ,OperTabella)
  xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
  xx=setCurrent(NomePagina,livelloPagina) 

  DescLoaded="0"  
  
  %>