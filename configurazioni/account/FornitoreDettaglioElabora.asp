   
   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->		
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   l_Id = "0"
   
   ao_lbd = "Cognome/Rag.Sociale"       'descrizione label 
   ao_nid = "DescCognome" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"DescCognome")       'valore di default
   ao_Plh = "|placeholder=Cognome/Rag.Sociale"              'placeholder
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Nome/Rag.Sociale"         'descrizione label 
   ao_nid = "DescNome" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"DescNome")        'valore di default
   ao_Plh = "|placeholder=Nome/Rag.Sociale"      'placeholder 
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   ao_lbd = "Tipo Ditta"                       'descrizione label 
   ao_nid = "IdTipoDitta" & l_Id              'nome ed id
   ao_val = GetDiz(DizDatabase,"IdTipoDitta") 'valore di default
   

   ao_Tex = "SELECT * From TipoDitta order By DescTipoDitta"
   ao_ids = "IdTipoDitta"             'valore della select 
   ao_des = "DescTipoDitta"           'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "0"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   ao_lbd = "Tipo Societa'"                   'descrizione label 
   ao_nid = "IdTipoSocieta" & l_Id              'nome ed id
   ao_val = GetDiz(DizDatabase,"IdTipoSocieta") 'valore di default
   

   ao_Tex = "SELECT * From TipoSocieta order By DescTipoSocieta"
   ao_ids = "IdTipoSocieta"			  'valore della select 
   ao_des = "DescTipoSocieta"         'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "0"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"--> 
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   ao_lbd = "Tipo Mandato"                       'descrizione label 
   ao_nid = "IdTipoMandato" & l_Id              'nome ed id
   ao_val = GetDiz(DizDatabase,"IdTipoMandato") 'valore di default
   

   ao_Tex = "SELECT * From TipoMandato  order By DescTipoMandato"
   ao_ids = "IdTipoMandato"			  'valore della select 
   ao_des = "DescTipoMandato"         'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "0"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->   
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   ao_lbd = "Tipo Incasso"                       'descrizione label 
   ao_nid = "IdTipoIncasso" & l_Id              'nome ed id
   ao_val = GetDiz(DizDatabase,"IdTipoIncasso") 'valore di default
   

   ao_Tex = "SELECT * From TipoIncasso order By DescTipoIncasso"
   ao_ids = "IdTipoIncasso"			  'valore della select 
   ao_des = "DescTipoIncasso"         'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "0"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->      
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Codice Fiscale"           'descrizione label 
   ao_3ls = "col-6"                       'size terzo elemento	
   ao_div = "col-4"				   
   ao_nid = "CodiceFiscale" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"CodiceFiscale")       'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Partita Iva"             'descrizione label 
   ao_3ls = "col-6"                   'size terzo elemento	
   ao_div = "col-4"				   
   ao_nid = "PartitaIva" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"PartitaIva")       'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   ao_3ls = "col-7"                   'size terzo elemento	
   ao_div = "col-3"    
   ao_lbd = "Rui"                   'descrizione label 
   ao_nid = "IdSezioneRui" & l_Id              'nome ed id
   ao_val = GetDiz(DizDatabase,"IdSezioneRui") 'valore di default
   

   ao_Tex = "SELECT * From TipoRUI order By DescTipoRUI"
   ao_ids = "IdTipoRUI"				  'valore della select 
   ao_des = "DescTipoRUI"              'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"--> 

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_3ls = "col-7"                   'size terzo elemento	
   ao_div = "col-3"    
   ao_lbd = "Numero RUI"         'descrizione label 
   ao_nid = "NumeroRui" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"NumeroRui")        'valore di default
   ao_Plh = "|placeholder="      'placeholder 
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->   

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_3ls = "col-7"                   'size terzo elemento	
   ao_div = "col-3"   
   ao_Cla = "|classe=form-control mydatepicker"
   ao_lbd = "Iscrizione RUI"         'descrizione label 
   ao_nid = "DataIscrizioneRui" & l_Id            'nome ed id
   tmpStr = GetDiz(DizDatabase,"DataIscrizioneRui")
   if tmpStr=0 then 
      tmpstr=""
   else
      Tmpstr=StoD(tmpStr)
   end if 
   ao_val = "|value=" & tmpStr        'valore di default
   ao_Plh = ""      'placeholder 
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->    


   
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Ruolo"         'descrizione label 
   ao_nid = "DescRuolo" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"DescRuolo")        'valore di default
   ao_Plh = "|placeholder=Ruolo"      'placeholder 
   
   
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->
	

<% if Cdbl(IdAccount)>0 then 

   %>
   
   <!--#include virtual="/gscVirtual/include/setDataForCall.asp"--> 
   <input type="hidden" name="sendDataForUpd"   id="sendDataForUpd"   value="<%=sendData%>">
   <input type="hidden" name="sendDataProd"     id="sendDataProd"     value="<%=sendData%>">
   <%
   NomeStruttura     = "SEDI_FORNITORE"
   DescStruttura     = "Sedi fornitore"
   flagOperStruttura = "CUD"
   ProfiloAccount    = "FORN"
   %>
   <!--#include virtual="/gscVirtual/configurazioni/sedi/StrutturaSede.asp"--> 

   <%
   NomeStruttura     = "CONTATTI_FORNITORE"
   DescStruttura     = "Contatti fornitore"
   flagOperStruttura = "CUD"
   ProfiloAccount    = "FORN"
   %>
   <!--#include virtual="/gscVirtual/configurazioni/contatti/StrutturaContatto.asp"--> 

<!-- Compagnie Account -->
<script>
function AssCompForn(id,sendData)
{
   
   var op="D";
   if($("#checkComp" + id).prop("checked") == true)
      op="I";
   var dataIn="op=" + op + "&sendData=" + sendData;  
   var vp=$("#localVirtualPath").val();
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "utility/updateData.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
        if (msg=="")
			esito = true;
		else {
			esito = false;
		}
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
	xx=LoadProdotti('0');
  
}
</script>
  <a class="btn btn-info" data-toggle="collapse" href="#collapseCompagnia" role="button" 
     onclick="evaluateMinusPlus('Compagnia')"
     aria-expanded="false" aria-controls="collapseCompagnia">
	 <span Id="Compagnia_Plus"><i class="fa fa-1x fa-plus-circle"></i></span>
	 <span Id="Compagnia_Minus" style= "display:none"><i class="fa fa-1x fa-minus-circle"></i></span>
	 <input type="hidden" id="Compagnia_plusMinus" value = "+">
	 </a>
	 <B> Compagnie </B>
  </p> 
  
	<div class="row">
	  <div class="col">
		<div class="collapse" id="collapseCompagnia">
			<div class="table-responsive" id="div_Compagnia">
			<table class="table"><tbody>
			<thead>
				<tr>
					<th scope="col">Compagnia</th>
					<th scope="col">Seleziona</th>
				</tr>
			</thead> 
<%
   err.clear
   MySql = "" 
   MySql = MySql & " Select A.*,isnull(b.IdCompagnia,0) as idEventoB From Compagnia A left join AccountCompagnia b "
   MySql = MySql & " on    A.IdCompagnia = B.IdCompagnia "
   MySql = MySql & " and   B.IdAccount = " & IdAccount
   MySql = MySql & " order By A.DescCompagnia"

   Set RsDet = Server.CreateObject("ADODB.Recordset")
   RsDet.CursorLocation = 3 
   RsDet.Open MySql, ConnMsde 
   
   Do While Not RsDet.EOF and err.number=0
      Id=RsDet("IdCompagnia")
	  selezionato=""
	  if Cdbl(RsDet("idEventoB"))>0 then 
	     selezionato=" checked "
	  end if 
	  
	  sendData="Oper=UpdCompForn|IdCompagnia=" & Id & "|" & "IdAccount=" & IdAccount
	  sendData=CryptWithKey(sendData,Session("CryptKey"))
	  
   %>
			<tr scope="col"> 
				<td>
					<input class="form-control" type="text" readonly value="<%=RsDet("DescCompagnia")%>">
				</td>
				<td><div class="form-check">
						<input id="checkComp<%=Id%>" <%=selezionato%> name="checkComp<%=Id%>" 
						type="checkbox" value = "S" class="big-checkbox" onclick="AssCompForn(<%=Id%>,'<%=sendData%>')">
					</div>		
				</td>
			</tr> 
	<%	
		RsDet.MoveNext
	Loop
    RsDet.close
    %>	
			
			</table>
			</div>
		</div>
	  </div>
	</div> 

<script>
function LoadProdotti(x)
{
   if (x=='1')
      evaluateMinusPlus('Prodotto');
	  
   var act=$("#Prodotto_plusMinus").val();

   var dataIn="sendData=" + $("#sendDataProd").val(); 
   //alert(dataIn);
   var vp=$("#localVirtualPath").val();   
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/ProdottiAccount.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	    $("#collapseProdotto").html(msg); 
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
  
}
</script>

  <a class="btn btn-info" data-toggle="collapse" href="#collapseProdotto" role="button" 
     onclick="LoadProdotti('1')"
     aria-expanded="false" aria-controls="collapseProdotto">
	 <span Id="Prodotto_Plus"><i class="fa fa-1x fa-plus-circle"></i></span>
	 <span Id="Prodotto_Minus" style= "display:none"><i class="fa fa-1x fa-minus-circle"></i></span>
	 <input type="hidden" id="Prodotto_plusMinus" value = "+">
	 </a>
	 <B> Prodotti forniti </B>
  </p> 	
   <div class="collapse multi-collapse" id="collapseProdotto">
   </div>    
  <%end if %>
 
    <%if SoloLettura=false then%>
		<div class="row"><div class="mx-auto">
		<%
		If OperTabella="CALL_DEL" then 
		    RiferimentoA="center;#;;2;dele;Cancella; Cancella;localDelete('submit','0');S"
		else 
			RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"
		end if 
		%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div></div>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
		</div>
		<br>
   <%end if %>
