<%
Set EmptyD = Server.CreateObject("Scripting.Dictionary")
Set EmptyC = Server.CreateObject("Scripting.Dictionary")
Set EmptyP = Server.CreateObject("Scripting.Dictionary")

'variabili di sessione globali per navigazione 
Session("virtualPath")     = virtualPath
Session("cryptPass")       = generatePassword(10)

'variabili di navigazione per login
Session("sideBar_u")        = ""

IdAzienda=Cdbl(TestNumeroPos("0" & Request("IdAzienda")))
if Cdbl(IdAzienda)=0 then 
   IdAzienda=1
end if 
Session("IdAziendaMaster")  = IdAzienda
Session("IdAzienda")        = 0
Session("IdAziendaWork")    = 0
Session("IdSubAziende")     = "0"
session("LivelloAccount")   = 0
set session("AziendaWork")  = EmptyD
set session("ClienteWork")  = EmptyC

Session("LoginTipoUtente")         = ""
Session("IdProfiloAbilitazione")   = "0"
Session("LoginIdAccount")          = "0"
Session("LoginIdAccountLev1")      = "0"
Session("LoginIdAccountLev2")      = "0"
session("FlagGeneraCollaboratore") = "0"
Session("LoginRefAccountLev1")     = "0"
Session("LoginRefAccountLev2")     = "0"
Session("LoginRefAccountLev3")     = "0"
Session("LoginTipoCollaboratore")  = ""
Session("Login_servizi_attivi")    = ""
Session("LoginIdCliente")          = "0"
Session("LoginNominativo")         = ""
Session("LoginHomePage")           = "/gscVirtual/login.asp"
Session("LoginExtePage")           = ""
'idAccount per i prodotti utilizzabili 
Session("LoginIdAccountProdotti")  = "0"
set session("Login_Parametri")     = EmptyP

%>