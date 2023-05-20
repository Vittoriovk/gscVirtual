<script>
function creaZipAttivita(idAttivita,idNumAttivita)
{
   var vp=$("#hiddenVirtualPath").val();
   var dataIn="IdAttivita=" + idAttivita + "&IdNumAttivita=" + idNumAttivita;
   
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "Servizio/AttivitaScaricaZip.asp",
	  data: dataIn,
      dataType: "html",
      success: function(msg)
      {
        retVal = msg;
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
	    var descE = xhr.status + ":Chiamata esegui Upload fallita, si prega di riprovare..." + thrownError;
        alert(descE);
		retVal = "ERR:Chiamata Fallita"
      }
    });
	document.Fdati.submit();
}
function creaReqAttivita(idAttivita,idNumAttivita)
{
   
   var vp=$("#hiddenVirtualPath").val();
   var dataIn="IdAttivita=" + idAttivita + "&IdNumAttivita=" + idNumAttivita;
   
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "Servizio/AttivitaScaricaRequest.asp",
	  data: dataIn,
      dataType: "html",
      success: function(msg)
      {
        retVal = msg;
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
	    var descE = xhr.status + ":Chiamata esegui Upload fallita, si prega di riprovare..." + thrownError;
        alert(descE);
		retVal = "ERR:Chiamata Fallita"
      }
    });
	document.Fdati.submit();
}
</script>
</script>

<%
function AggiungiDocAttivita(IdProdotto,IdAccountFornitore,IdAttivita,IdNumAttivita,TipoUtente,TipoDoc)
Dim q,qn,qi 
   on error resume next 
   
   qn = ""
   qn = qn & " select IdDocumento "
   qn = qn & " from AttivitaDocumento"
   qn = qn & " Where IdAttivita='" & apici(IdAttivita) & "'"
   qn = qn & " and IdNumAttivita=" & IdNumAttivita
  
   if cdbl(IdAccountFornitore)>0 then 
      q = ""
      q = q & " select '" & IdAttivita & "' as IdAttivita"
      q = q & "," & IdNumAttivita & " as IdNumAttivita "
      q = q & ",A.IdDocumento,B.DescDocumento,FlagObbligatorio,FlagDataScadenza,0 as IdUpload "
      q = q & " from AccountProdottoDocAff A,Documento B"
      q = q & " Where A.IdAccount="  & IdAccountFornitore 
      q = q & " and   A.IdProdotto=" & IdProdotto
      q = q & " and   A.IdDocumento = b.IdDocumento"
	  q = q & " and   A.TipoDoc = '" &  TipoDoc & "'"
      if TipoUtente<>"" then 
         q = q & " and (DITT='" & TipoUtente & "' or PEGC='" & TipoUtente & "' or PEGI='" & TipoUtente & "' or PEFI='" & TipoUtente & "') "
      end if 
      q = q & " and a.IdDocumento not in (" & qn & ")"
   else
      q = ""
      q = q & " select '" & IdAttivita & "' as IdAttivita"
      q = q & "," & IdNumAttivita & " as IdNumAttivita "
      q = q & ",A.IdDocumento,B.DescDocumento,max(FlagObbligatorio),max(FlagDataScadenza),0 as IdUpload "
      q = q & " from AccountProdottoDocAff A,Documento B"
      q = q & " Where A.IdProdotto=" & IdProdotto
      q = q & " and   A.IdDocumento = b.IdDocumento"
      if TipoUtente<>"" then 
         q = q & " and (DITT='" & TipoUtente & "' or PEGI='" & TipoUtente & "' or PEFI='" & TipoUtente & "') "
      end if 
      q = q & " and a.IdDocumento not in (" & qn & ")"
	  q = q & " group by A.IdDocumento,B.DescDocumento" 
   
   end if 
   'response.write q
   qi = ""
   qi = qi & " insert into AttivitaDocumento "
   qi = qi & "(IdAttivita,IdNumAttivita,IdDocumento,DescDocumento"
   qi = qi & ",FlagObbligatorio ,FlagDataScadenza,IdUpload)"
   qi = qi & q 
   
   'response.write qi
   connMsde.execute qi
   err.clear 
   
end function 

function SetCampoAttivita(id,campo,tipo,Valore)
dim q
    q = ""
	q = q & " update AttivitaDocumento "
	if tipo="N" then 
	   q = q & " Set " & campo & "=" & valore
	else 
	   q = q & " Set " & campo & "='" & apici(valore) & "'"
	end if 
	q = q & " Where AttivitaDocumento=" & id
	connMsde.execute q
	err.clear 
end function 
function SetAttivitaDocumentoValido(id,note)
dim q
   q = ""
   q = q & " update AttivitaDocumento set "
   q = q & " IdTipoValidazione = 'VALIDO'"
   q = q & ",DataValidazione = getDate()" 
   q = q & ",NoteValidazione ='" & apici(note) & "'"
   q = q & " Where IdAttivitaDocumento=" & id
   connMsde.execute q
   err.clear 
end function 
function SetAttivitaDocumentoNonValido(id,note)
dim q
   q = ""
   q = q & " update AttivitaDocumento set "
   q = q & " IdTipoValidazione = 'NONVAL'"
   q = q & ",DataValidazione = getDate()" 
   q = q & ",NoteValidazione ='" & apici(note) & "'"
   q = q & " Where IdAttivitaDocumento=" & id
   connMsde.execute q
   err.clear 
end function 
function SetAttivitaDocumentoRichiesto(id)
dim q
   q = ""
   q = q & " update AttivitaDocumento set "
   q = q & " FlagObbligatorio = 1"
   q = q & " Where IdAttivitaDocumento=" & id
   connMsde.execute q
   err.clear 
end function
function SetAttivitaDocumentoNonRichiesto(id)
dim q
   q = ""
   q = q & " update AttivitaDocumento set "
   q = q & " FlagObbligatorio = 0"
   q = q & " Where IdAttivitaDocumento=" & id
   connMsde.execute q
   err.clear 
end function
function SetAttivitaDocumentoDelete(id)
dim q
   q = ""
   q = q & " delete from AttivitaDocumento  "
   q = q & " Where IdAttivitaDocumento=" & id
   connMsde.execute q
   err.clear 
end function 

function SetAttivitaDocumentoAdd(IdAttivita,IdNumAttivita,IdDocumento,DescDocumento,FlagObbligatorio,FlagDataScadenza)
dim qi 
   on error resume next 
   qi = ""
   qi = qi & " insert into AttivitaDocumento "
   qi = qi & "(IdAttivita,IdNumAttivita,IdDocumento,DescDocumento"
   qi = qi & ",FlagObbligatorio ,FlagDataScadenza,IdUpload)"
   qi = qi & " values ("
   qi = qi & " '" & IdAttivita & "'"
   qi = qi & ", " & NumForDb(IdNumAttivita)
   qi = qi & ", " & NumForDb(IdDocumento)
   qi = qi & ",'" & apici(DescDocumento) & "'"
   qi = qi & ", " & NumForDb(FlagObbligatorio)
   qi = qi & ", " & NumForDb(FlagDataScadenza)
   qi = qi & ",0)"
   'response.write qi 
   connMsde.execute qi 
   err.clear
end function 
   

%>