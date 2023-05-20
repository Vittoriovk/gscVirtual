<script>
function GCF_CostoFirma_Change(ref)
{
	var opt=$('#IdTipoFirma0').val();
	var imp=$('#GCF_CostoFirma_' + opt).val();
	if (IsNumber(imp)==false)
	   imp='0,00';
	else { 
	   var ii=Number.parseFloat(imp).toFixed(2);
       imp=ii.replace('.',',');
    }

	$('#' + ref).val(imp);
	xx = tipoFirmaIsChanged();
}
</script>
<%
flagDebug=false
Oper = ""
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
   flagDebug = false 
end if

if Cdbl(IdCompagnia)>0 and Cdbl(IdFornitore)>0 then
   IdAccForn  = Cdbl(LeggiCampo("select * from Fornitore Where IdFornitore=" & IdFornitore,"IdAccount")) 
   qProd      = "select * from Prodotto where IdCompagnia=" & IdCompagnia & " and IdAnagServizio='CAUZ_PROV'"
   idProdotto = Cdbl(LeggiCampo(qProd,"IdProdotto"))
   qFascia    = qFascia
   qFascia    = qFascia & " Select * from AccountProdottoFirma A, TipoFirma B"
   qFascia    = qFascia & " Where IdAccount = " & idAccForn
   qFascia    = qFascia & " and IdProdotto = "  & idProdotto
   qFascia    = qFascia & " and A.IdTipoFirma = B.IdTipoFirma" 
   if NoChangeSel="S" then 
      qFascia    = qFascia & " and A.IdTipoFirma = '" & apici(IdTipoFirma) & "'"
   end if 
   qFascia    = qFascia & " order by B.DescTipoFirma"

   if flagDebug=true then 
      response.write qFascia & "<br>"
   end if    

   Set Rs = Server.CreateObject("ADODB.Recordset")
   'response.write qFascia
   Rs.CursorLocation = 3
   Rs.Open qFascia, ConnMsde 
   do while not Rs.eof 
      nm="GCF_CostoFirma_" & Rs("IdTipoFirma")
   %>
   <input type="hidden" name="<%=nm%>" id="<%=nm%>" value="<%=Rs("CostoFirma")%>">
   <%
      Rs.Movenext 
   loop 
   Rs.close 
end if 

xx=ShowLabel("Firma")
stdClass="class='form-control form-control-sm'"
tt=1
if NoChangeSel="S" then 
   tt=0
   stdClass = stdClass & " disabled='true' "
end if 
cambio="" 
if IdRefCampo<>"" and NoChangeSel<>"S" then 
   cambio="GCF_CostoFirma_Change('" & IdRefCampo & "')"
end if 
response.write ListaDbChangeCompleta(qFascia,"IdTipoFirma0",IdTipoFirma ,"IdTipoFirma","DescTipoFirma" ,tt,cambio,"","","","",stdClass)

%>