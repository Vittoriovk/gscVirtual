<%NomePagina="swapEvento.asp"
  

%>
<!--#include virtual="/gscVirtual/include/includeStd.asp"-->
<%
on error resume next
  Set Rs = Server.CreateObject("ADODB.Recordset")
  linkRef=""
  swap1=""
  swap2=""
  swap3=""
  IdEvento=Request("IdEvento")
  MySql = "" 
  MySql = MySql & " select * from Evento Where IdEvento = " & IdEvento
  'response.write MySql 
  Rs.CursorLocation = 3 
  Rs.Open MySql, ConnMsde
  'response.write err.description & rs("IdProcesso")
  if ucase(rs("IdProcesso"))="AFFI" and Ucase(RS("IdTabella"))=ucase("AffidamentoRichiestaComp") then 
     qKey = ""
     qKey = qKey & " select * from "
     qKey = qKey & " AffidamentoRichiestaComp A,AffidamentoRichiesta B,Cliente C "
     qKey = qKey & " Where " & Rs("IdKey")
     qKey = qKey & " and A.IdAffidamentoRichiesta=B.IdAffidamentoRichiesta"
     qKey = qKey & " and B.IdAccountCliente = C.IdAccount"
     'response.write qKey 
     descCliente = LeggiCampo(qKey,"Denominazione")
	 'response.write descCliente 
	 'response.end 
	 if descCliente<>"" then 
	    swap1    = "TipoRicercaExt"
		swap1Val = "1"
		swap2    = "testo_ricercaExt"
		swap2Val = descCliente
		linkRef = "/gscVirtual/configurazioni/clienti/Affidamento/ListaRichiesta.asp"
     end if 
  end if 
  if ucase(rs("IdProcesso"))="COOB" and Ucase(RS("IdTabella"))=ucase("AccountCoobbligato") then 
     qKey = ""
     qKey = qKey & " select C.* from "
     qKey = qKey & " AccountCoobbligato A,Cliente C "
     qKey = qKey & " Where " & Rs("IdKey")
     qKey = qKey & " and A.IdAccount = C.IdAccount"
     'response.write qKey 
     descCliente = LeggiCampo(qKey,"Denominazione")
     IdCliente   = cdbl("0" & LeggiCampo(qKey,"IdCliente"))
     IdAccount   = cdbl("0" & LeggiCampo(qKey,"IdAccount"))
	 
	 'response.write descCliente 
	 'response.end 
	 if isBackOffice() and descCliente<>"" then 
	    swap1    = "TipoRicercaExt"
		swap1Val = "1"
		swap2    = "testo_ricercaExt"
		swap2Val = descCliente
		linkRef = "/gscVirtual/configurazioni/Clienti/ValidazioneCoobbligatoBackO.asp"
     end if 
	 if isBackOffice()=false then 
	    swap1    = "IdCliente"
		swap1Val = IdCliente 
		swap2    = "IdAccount"
		swap2Val = IdAccount
		swap3    = "DescCliente"
		swap3Val = descCliente
		linkRef = "/gscVirtual/configurazioni/Clienti/ClienteCoobbligati.asp"
     end if 
  end if 
  
  Rs.close 
  if linkRef<>"" then 
     xx=RemoveSwap()
	 if swap1<>"" then 
        Session("swap_" & swap1) = swap1Val
	 end if 
	 if swap2<>"" then 
        Session("swap_" & swap2) = swap2Val
	 end if 
	 if swap3<>"" then 
        Session("swap_" & swap3) = swap3Val
	 end if 
     response.redirect linkRef 
     response.end 
  end if 
  response.redirect Session("LoginHomePage")
  response.end 
   
%>