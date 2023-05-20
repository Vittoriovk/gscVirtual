<script>
function reverseCF(idCf,cognome,nome,comune,descProv,siglaProv,sesso,dataNasc,descStato,siglaStato)
{
   
   var tm="";
   var cf=$("#" + idCf).val();
   cf=cf.trim();
   
   if (cf.length>0) {
      tm=getDataDaCf(cf,0);
      if (tm.length>0)
          $("#" + dataNasc).val(tm);  
      tm=getDescComuneDaCF(cf);
      if (tm.length>0) {
          $("#" + comune).val(tm);
		  if (descProv.trim().length>0) {
             tm=getDescProvinciaDaCF(cf);
             $("#" + descProv).val(tm);  
		  }
		  if (siglaProv.trim().length>0) {
             tm=getSiglaProvinciaDaCF(cf);
             $("#" + siglaProv).val(tm);  
		  }
      }
	  if (descStato.trim().length>0) {
         tm=getDescStatoDaCF(cf);
         $("#" + descStato).val(tm);  
	  }
	  if (siglaStato.trim().length>0) {
         tm=getIdStatoDaCF(cf);
         $("#" + siglaStato).val(tm);  
	  }
	  if (sesso.trim().length>0) {
         tm=getSessoDaCF(cf);
         $("#" + sesso).val(tm);  
	  }      
   
   }   
   
   
}
function getSessoDaCF(cf)
{
   var retVal = "M"; 
   var gg = cf.substring(9,11);
   if (gg>"40")
      retVal="F";
   return retVal ;
}

function getDataDaCf(cf,limite)
{
   var dataIn="action=getDataDaCf&cf=" + encodeURI(cf);
   var esito = callFunctionCf(dataIn);
   return esito ;
}
function getDescComuneDaCF(cf)
{
   var dataIn="action=getDescComuneDaCF&cf=" + encodeURI(cf);
   var esito = callFunctionCf(dataIn);
   return esito ;
}
function getDescProvinciaDaCF(cf)
{
   var dataIn="action=getDescProvinciaDaCF&cf=" + encodeURI(cf);
   var esito = callFunctionCf(dataIn);
   return esito ;
}

function getSiglaProvinciaDaCF(cf)
{
   var dataIn="action=getSiglaProvinciaDaCF&cf=" + encodeURI(cf);
   var esito = callFunctionCf(dataIn);
   return esito ;
}

function getSiglaProvinciaDaProvincia(id)
{
   var dataIn="action=getSiglaProvinciaDaProvincia&id=" + encodeURI(id);
   var esito = callFunctionCf(dataIn);
   return esito ;
}

function getIdStatoDaCF(cf)
{
   var dataIn="action=getIdStatoDaCF&cf=" + encodeURI(cf);
   var esito = callFunctionCf(dataIn);
   return esito ;
}

function getDescStatoDaCF(cf)
{
   var dataIn="action=getDescStatoDaCf&cf=" + encodeURI(cf);
   var esito = callFunctionCf(dataIn);
   return esito ;
}

function getCodiceCatasto(stato,provincia,comune)
{
   var dataIn="action=getCodiceCatasto&stato=" + encodeURI(stato) + "&provincia=" + encodeURI(provincia) + "&comune=" + encodeURI(comune) ;
   var esito = callFunctionCf(dataIn);
   return esito ;
}

function callFunctionCf(dataIn)
{
   var vp=$("#hiddenVirtualPath").val(); 
   var esito="";
   $.ajax({
      type: "POST",
	  async: false,
      url: vp + "/utility/loadFunctionCf.asp",
      data: dataIn,
      dataType: "html",
      success: function(msg)
      {
	   esito = msg;
      },
      error: function(xhr, ajaxOptions, thrownError)
      {
        alert(xhr.status + ":Chiamata fallita, si prega di riprovare..." + thrownError);
      }
    });   
	return esito;
  
}

function estraiData(gg,mm,aa,sex){
   const MONTH_CODES = ['A','B','C','D','E','H','L','M','P','R','S','T'];
   
   if (gg==0 || mm==0 || aa==0 || sex=="")
      return "";
   
   var anno = aa.toString();
   anno = anno.substr(anno.length - 2, 2);
   var mese = MONTH_CODES[mm-1];
   
   if (sex.toUpperCase() == 'F') {
      gg += 40;
   }
   gior = '0' + gg.toString();
   gior = gior.substr(gior.length - 2, 2);
   return anno + mese + gior;
   
}   
  
function calcolaCF(cognome,nome,sesso,dtna,stato,provincia,comune) {
   var esito="";
   var tmp="";
   if (cognome.length==0 || nome.length==0 || sesso.length==0)
      return esito;
   tmp = sesso.trim().toUpperCase();
   if (!(tmp=="M" || tmp=="F"))
      return esito;  
   tmp = stato.trim().toUpperCase();
   if (tmp=="")
      return esito;  
   
   var codCognome = estraiConsonanti(cognome.toUpperCase()) + estraiVocali(cognome.toUpperCase()) + "XXX";
   codCognome = codCognome.substr(0, 3);
  
   esito = codCognome;
   var codNome    = estraiConsonanti(nome.toUpperCase());  
   if (codNome.length >= 4) {
      codNome = codNome.charAt(0) + codNome.charAt(2) + codNome.charAt(3)
    } else {
      codNome += estraiVocali(nome.toUpperCase());
      codNome = codNome.substr(0, 3);
    }
   if (codNome.length!=3)
      return esito;    
   esito = esito + codNome;
   
   var gg = dtna.substring(0,2);
   var ggn = parseInt(gg);
   if (isNaN(ggn))
      return esito;

   var mm = dtna.substring(3,5);
   var mmn = parseInt(mm);
   if (isNaN(mmn))
      return esito;
   
   var aa = dtna.substring(6);
   var aan = parseInt(aa);
   if (isNaN(aan))
      return esito;

   var codData = estraiData(ggn,mmn,aan,sesso);
   if (codData.length!=5)
      return esito;      
   esito = esito + codData;
 
   var codCata = getCodiceCatasto(stato,provincia,comune);
   if (codCata.length!=4)
      return esito;      
   esito = esito + codCata;
   var codChec = getCheckCode(esito);
   
   if (codChec.length!=1)
      return esito;      
   esito = esito + codChec;
   return esito;
   
}
  function getCheckCode (codiceFiscale) {
    const CHECK_CODE_CHARS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const CHECK_CODE_ODD  = {  0: 1,  1: 0,  2: 5,  3: 7,  4: 9,  5: 13,  6: 15,  7: 17,  8: 19,  9: 21,  A: 1,  B: 0,  C: 5,  D: 7,  E: 9,  F: 13,  G: 15,  H: 17,  I: 1,  J: 21,  K: 2,  L: 4,  M: 18,  N: 20,  O: 11,  P: 3,  Q: 6,  R: 8,  S: 12,  T: 14,  U: 16,  V: 10,  W: 22,  X: 25,  Y: 24, Z: 23};
    const CHECK_CODE_EVEN = {  0: 0,  1: 1,  2: 2,  3: 3,  4: 4,  5: 5,  6: 6,  7: 7,  8: 8,  9: 9,  A: 0,  B: 1,  C: 2,  D: 3,  E: 4,  F: 5,  G: 6,  H: 7,  I: 8,  J: 9,  K: 10,  L: 11,  M: 12,  N: 13,  O: 14,  P: 15,  Q: 16,  R: 17,  S: 18,  T: 19,  U: 20,  V: 21,  W: 22,  X: 23,  Y: 24,  Z: 25};
    codiceFiscale = codiceFiscale.toUpperCase();
    let val = 0
    for (let i = 0; i < 15; i = i + 1) {
      const c = codiceFiscale[i];
      val += i % 2 !== 0 ? CHECK_CODE_EVEN[c] : CHECK_CODE_ODD[c];
    }
    val = val % 26;
    return CHECK_CODE_CHARS.charAt(val);
  }
  
</script>
