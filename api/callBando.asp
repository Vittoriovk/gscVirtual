<!--#include virtual="/gscVirtual/common/function.asp"-->
<!--#include virtual="/gscVirtual/common/functionNew.asp"-->
<!--#include virtual="/gscVirtual/common/connDb.asp"-->
<!--#include virtual="/gscVirtual/common/FunctionAccessoDb.asp"-->
<!--#include virtual="/gscVirtual/common/parametri.asp"-->
<!--#include virtual="/gscVirtual/common/FunCallOtherPage.asp"-->
<!--#include virtual="/gscVirtual/api/aspjson/aspJSON1.19.asp" -->
<!--#include virtual="/gscVirtual/api/FareAppalti/function.asp" -->

<%
on error resume next 
'riceve in input un pagamento di un account e ne richiede il pagamento 

retV=""
token=trim(getToken("nextbroker","0cc80b3e231b5281dfef46b82eafbb30461cce34cc7e4c35779b1c434db2516c5de0081b77efef7691e9f2f0e952a1e4aed7e03c3b4fe066cebf3235c1049a86"))
'response.write "token:" & token & "<hr>"
Dim DizDatabase
Set DizDatabase = CreateObject("Scripting.Dictionary")
  
lista=getBando(DizDatabase,token,"10332010")


%>