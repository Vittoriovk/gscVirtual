<!--
ns4 = (document.layers) ? true:false
ie4 = (document.all) ? true:false 
ng5 = (document.getElementById) ? true:false 
var RetValue="";
var imgMinus="../images/images/minus.gif";
var imgPlus="../images/images/plus.gif"; 

function noOper()
{
   return false;
}

function messaggioAlert(m)
{
   alert(m);
   return true;
}
function pulisciCampo(id) {
   $("#" + id).val("");
}
function salvaCampo(id) {
var xx;

   $("#Oper").val("SAVE_FIELD");
   $("#ItemToRemove").val(id);
   document.Fdati.submit();
}

function creaPassword(elem){
    var elencoCaratteri="abcdefghijklmnopqrstuvwxyzA.!_BCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";

    var minimoCaratteri=8;
    var massimoCaratteri=20;
    var differenzaCaratteri=massimoCaratteri-minimoCaratteri;

    var lunghezza=Math.round((Math.random()*differenzaCaratteri)+minimoCaratteri);

    var incremento=0;
    var password="";

    while(incremento<lunghezza){
       password+=elencoCaratteri.charAt(Math.round(Math.random()*elencoCaratteri.length));
       incremento++;
    }
    document.getElementById(elem).value=password;

}

function NonImplementata() {
  alert("funzione non attiva ");
  return false;
}

function GetEstensione(nome) {
  var extension = "";
  var lunghezza = nome.length;
  if (lunghezza>4)  {
     extension =     nome.slice(-4);
     extension =  extension.toUpperCase();
  }
  return extension;
}


function EstraiTag(StIn,TagIn)
{
    var X;
    X="";
    var Ptr1=StIn.indexOf('<'  + TagIn + '>');
    var Ptr2=StIn.indexOf('</' + TagIn + '>');
    if (Ptr2>Ptr1 && Ptr1>-1)
        X=StIn.substring(Ptr1+2+TagIn.length,Ptr2);
    return X;
} 
    
function GetNumberAsFloat(N) 
{
    return parseFloat(N.replace(",","."));
} 

function round(value, decimals) 
{
    return Number(Math.round(value+'e'+decimals)+'e-'+decimals);
}

function SetData(F)
{
   var d = new Date();
    
   var gg = d.getDate();
    if (gg<=9)
       gg="0"+gg;

   var mm = d.getMonth()+1;
    if (mm<=9)
       mm="0"+mm;
    
    var aa=d.getFullYear();
    v=gg+"/"+mm+"/"+aa;
   x=ImpostaValoreDi(F,v);

    
}

function SaveWithOper(Op)
{

   locName=ValoreDi("DescLoaded");
    if (Op=="INS")
        yy=ImpostaValoreDi("DescLoaded","0");

    xx=ElaboraControlli();
    
   yy=ImpostaValoreDi("DescLoaded",locName);
    
     if (xx==false)
       return false;
        
    ImpostaValoreDi("Oper",Op);
    document.Fdati.submit();
}

function SalvaSingoloEdAttiva(Op,Id,conferma,callFun,argFun1,argFun2)
{
    locName=ValoreDi("DescLoaded");
    yy=ImpostaValoreDi("DescLoaded",Id);

    xx=ElaboraControlli();
    
    yy=ImpostaValoreDi("DescLoaded",locName);
    
     if (xx==false)
       return false;

    if (callFun && typeof callFun === "function") {
        
        if (argFun1 && argFun2) {
            yy = callFun(argFun1,argFun2);
        } else if (argFun1) {
            yy = callFun(argFun1);
        } else {
            yy = callFun();
        }
    
    }
    
    if (conferma==true) {
        var yy = confirm("Confermi modifica dati ?");
        if (yy==false)
        return false;   
    }    
    yy=AttivaFunzione(Op,Id);
}

function AttivaFunzione(Funzione,Id)
{
    xx=ImpostaValoreDi("ItemToRemove",Id);
    xx=ImpostaValoreDi("Oper",Funzione);
    document.Fdati.submit();
}

function RemoveItem(Id)
{
    var yy = confirm("Confermi cancellazione dati ?")
    
     if (yy==false)
       return false;

    ImpostaValoreDi("ItemToRemove",Id);
    ImpostaValoreDi("Oper","RemoveItem");
    document.Fdati.submit();
}



function IsRadioChecked(Campo)
{

    var el,c,Controllo;
    Controllo=0;
    
    el=document.getElementById(Campo+"0").value;
    
    for (var i=1; i<=el; i++) 
    {
        c=Campo+i;
        xx=ImpostaColoreFocus(c,"","white");    
        if (IsChecked(c)==true)
        {
           Controllo=1;
        }
    }

    if (Controllo==1)
       return true;
       
    for (var i=1; i<=el; i++) 
    {
        c=Campo+i;
        xx=ImpostaColoreFocus(c,"","yellow");    
    }

}
function TestoDiOption(Campo)
{
    var v;
    if (ng5==true)
    {
        v=document.getElementById(Campo);
    }    
    else
    {
        v=document.all.item(Campo);
    }

    return v.options[v.selectedIndex].text;
}


function ValoreDi(Campo)
{
    try
    {
        var v;
        if (ng5==true)
        {
            v=document.getElementById(Campo).value;
        }    
        else
        {
            v=document.all.item(Campo).value;
        }
        return v;
    }
    catch(e)
    {
        alert('ValoreDi: ' + e.message + ' ' + Campo);
    }
}

function ImpostaValoreDi(Campo,v)
{
    try
    {
        if (ng5==true)
        {
            document.getElementById(Campo).value=v;
        }    
        else
        {
            document.all.item(Campo).value=v;
        }
        return true;
        
    }
    catch(e)
    {
        alert('ValoreDi: ' + e.message + ' ' + Campo);
    }
    
}
function DisabilitaAbilita(Campo,Esito)
{
    if (ng5==true)
    {
        document.getElementById(Campo).disabled=Esito;
    }
    else
    {
        document.all.item(Campo).disabled=Esito;
    }

}

function GetDisabilitato(Campo)
{
    if (ng5==true)
    {
        return document.getElementById(Campo).disabled;
    }
    else
    {
        return document.all.item(Campo).disabled;
    }

}

function IsChecked(Campo)
{
    try
    {
        if (ng5==true)
        {
            return document.getElementById(Campo).checked;
        }
        else
        {
            return document.all.item(Campo).checked;
        }
    }
    catch(e)
    {
        alert('Ischechecked: ' + e.message + ' ' + Campo);
    }

}

function GetBgColor(Campo)
{
    if (ng5==true)
    {
        return document.getElementById(Campo).style.backgroundColor;
    }
    else
    {
        return document.all.item(Campo).style.backgroundColor;
    }

}

function SetBgColor(Campo,Colore)
{
    if (ng5==true)
    {
        document.getElementById(Campo).style.backgroundColor=Colore;
    }
    else
    {
        document.all.item(Campo).style.backgroundColor=Colore;
    }

}


function ImpostaColoreFocus(Campo,Fuoco,Colore)
{
    try
    {
        if (ng5==true)
        {
            if (Fuoco=="S")
            {
                document.getElementById(Campo).focus();
            }

            if (!Colore=="")
            {
                document.getElementById(Campo).style.backgroundColor=Colore;
            }
        }
        else
        {
            if (Fuoco=="S")
                document.all.item("GG_"+c).focus();
            if (Colore!="")
                document.all.item("GG_"+c).style.backgroundColor=Colore;
        }
    }
    catch(e)
    {
        alert('ValoreDi: ' + e.message + ' ' + Campo);
    }
}

function GetColoreFocus(Campo)
{
    var d;
    try
    {
        if (ng5==true)
        {
            return    document.getElementById(Campo).style.backgroundColor;
        }
        else
        {
            return    document.all.item("GG_"+c).style.backgroundColor=Colore;
        }
    }
    catch(d)
    {
        alert('ValoreDi: ' + d.message + ' ' + Campo);
    }
}

function IsEmail(email)
{
    if (!email.match(/^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})$/))
    {
        return false;
    }
    return true;
}

function IsIBanIt(IbanIt)
{
    /*check validità iban italiano*/
    //alert(!IbanIt.match(/^IT\d{2}[ ][a-zA-Z]\d{3}[ ]\d{4}[ ]\d{4}[ ]\d{4}[ ]\d{4}[ ]\d{3}|IT\d{2}[a-zA-Z]\d{22}$/));
    //alert(!IbanIt.match(/^IT\d{2}[a-zA-Z]\d{3}\d{4}\d{4}\d{4}\d{4}[a-zA-Z]\d{2}|IT\d{2}[a-zA-Z]\d{22}$/));
    //                      IT 15 K        056     9603    2090    0001    0396 X57
    //                      IT 19 P        030     7502    200C    C850    0563 659
    if (!IbanIt.match(/^IT\d{2}[ ]?[a-zA-Z]\d{3}[ ]?\d{4}[ ]?\d{4}[ ]?\d{4}[ ]?\d{4}[ ]?\d{3}|IT\d{2}[a-zA-Z]\d{22}|IT\d{2}[ ]?[a-zA-Z]\d{3}[ ]?\d{4}[ ]?\d{4}[ ]?\d{4}[ ]?\d{4}[ ]?[a-zA-Z]\d{2}|IT\d{2}[a-zA-Z]\d{22}|IT\d{2}[ ]?[a-zA-Z]\d{3}[ ]?\d{4}[ ]?\d{3}[ ]?[a-zA-Z]?[a-zA-Z]?\d{3}[ ]?\d{4}[ ]?\d{3}$/))
    {
        return false;
    }
    return true;
}

function IsIBanW(IbanW)
{
    /*check validità iban generico*/
    if (!IbanW.match(/^[a-zA-Z]{2}[0-9]{2}[a-zA-Z0-9]{4}[0-9]{7}([a-zA-Z0-9]?){0,16}$/))
    {
        return false;
    }
    return true;
}

function IsBIC(BIC)
{
    /*check validità BIC */
    if (!BIC.match(/^([a-zA-Z]{4}[a-zA-Z]{2}[a-zA-Z0-9]{2}([a-zA-Z0-9]{3})?)$/))
    {
        return false;
    }
    return true;
}

function VerificaData(gginput,mminput,aaaainput)

 {
   
   var dataverifica;
   var dataoutput;

   var ggoutput;   
   var mmoutput; 
   var aaaaoutput;

     

// viene utilizzato un oggetto data per la verifica costruendola in base
// ai parametri ricavati in input (l'anno è sottratto di 1 perchè la funzione data
// considera i mesi partendo da zero)

   //alert("DatiIn:"+aaaainput+mminput+gginput); 
   dataverifica =new Date(aaaainput,mminput-1,gginput)

   
   
// dalla data creata viene ricavato l'anno, il mese ed il giorno (il mese è incrementato
// di 1 per restituire il mese nel formato standard)

   aaaaoutput=dataverifica.getFullYear().toString(); 
   mmoutput=(dataverifica.getMonth()+1).toString(); 
   ggoutput=dataverifica.getDate().toString();

// poichè le funzioni getMonth e getDate restituiscono valori numerici
// occorre aggiungere lo zero per i mesi e gli anni unitari (es "1/1/2005")
// il controllo è effettuato sulla lunghezza della stringa (se < 2 aggiungi lo zero)


   if (mmoutput.length<2)
     mmoutput="0"+mmoutput;

   if (mminput.length<2)
     mminput="0"+mminput;
     
   if (ggoutput.length<2)
     ggoutput="0"+ggoutput;

   if (gginput.length<2)
     gginput="0"+gginput;

// La data in output è costituita dai dati ricavati precedentemente separati
// dal carattere "/" dopo aver aggiunto 


   dataoutput=ggoutput+"/"+mmoutput+"/"+aaaaoutput;


// viene confrontata la data in input con quella in output
// se non coincidono vuol dire che la data non è corretta


   if (gginput!=ggoutput || mminput!=mmoutput || aaaainput!=aaaaoutput)
     return false;
     
   else
     return true; 
   
 }

 function ltrim(str) { 
    for(var k = 0; k < str.length && isWhitespace(str.charAt(k)); k++);
    return str.substring(k, str.length);
}
function rtrim(str) {
    for(var j=str.length-1; j>=0 && isWhitespace(str.charAt(j)) ; j--) ;
    return str.substring(0,j+1);
}
function trim(str) {
    return ltrim(rtrim(str));
}
function isWhitespace(charToCheck) {
    var whitespaceChars = " \t\n\r\f";
    return (whitespaceChars.indexOf(charToCheck) != -1);
}

function IsNumber(numero) {
var x;
    
    if (String(numero).indexOf(".") != (-1))
    return false;
    x = new String(numero);
    x = x.replace(",",".");
    
    if (isNaN(x)==true)
    return false;
    
    return true;
}

function IsInteger(numero) {
    if (isNaN(numero)==true)
    return false;
            
    if (parseInt(numero)!=numero)
    return false;
}

function CheckTipo(Valore,Tipo) {
var Dati = new Array(),g,m,a;

    RetValue="";
    if (Tipo.length==3 && Tipo.substring(0,2)!="ML")
    {
        if (Tipo.substr(2,1)=='O' || Tipo.substr(2,1)=='P' || Tipo.substr(2,1)=='Z')
        {
        g=trim(Valore);
        if (g.length==0)
            return false;
        }
        Tchk=Tipo.substr(2,1);
        Tipo=Tipo.substr(0,2);    
        
        if (Tipo=="IN" || Tipo=="FL")
        {    
           
            if (IsNumber(Valore)==true)
            {    
                Num4 = GetNumberAsFloat(Valore);

                if (Num4==0 && Tchk=="O")
                    return false;
                if (Num4<=0 && Tchk=="P")
                    return false;
                if (Num4<0 && Tchk=="Z")
                    return false;
                if (Num4<0 && Tchk=="Q")
                    return false;
                if (Num4>100 && Tchk=="Q")
                    return false;
            }
        }
    }
    if (Tipo.substr(0,2)=="ML")
    {
        Num4 = GetNumberAsFloat(Tipo.substr(2));
        g=trim(Valore);
        if (g.length>Num4)
            return false;
        else    
            return true;
    }
    
    if (Tipo=="FL")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        else    
            return IsNumber(Valore);
    }

    if (Tipo=="LI")
    {
        g=trim(Valore);
        if (g=="-1")
            return false;
        else
            return true;
    }

    if (Tipo=="EMO")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        Tipo="EM";
    }

    if (Tipo=="EM")
    {
        g=trim(Valore);
        if (g.length==0)
            return true;
        else
            return IsEmail(Valore);
    }
    
    /* IBAN Italiano */
    if (Tipo=="IBO")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        Tipo="IB";
    }

    if (Tipo=="IB")
    {
        g=trim(Valore);
        
        /*elimino spazi interni*/
        var ixiban=1;
        
        while (ixiban > 0) 
        {
            Valore=Valore.replace(" ","");
            ixiban=Valore.indexOf(" ");
        }
        
        if (g.length==0)
            return true;
        else
            return IsIBanIt(Valore);
    }
    
    /* IBAN Generico */
    if (Tipo=="WBO")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        Tipo="WB";
    }

    if (Tipo=="WB")
    {
        g=trim(Valore);
        
        /*elimino spazi interni*/
        var ixiban=1;
        
        while (ixiban > 0) 
        {
            Valore=Valore.replace(" ","");
            ixiban=Valore.indexOf(" ");
        }
        
        if (g.length==0)
            return true;
        else
            return IsIbanW(Valore);
    }
    
    
    /* BIC Banca */
    if (Tipo=="BIO")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        Tipo="BI";
    }

    if (Tipo=="BI")
    {
        g=trim(Valore);
        
        /*elimino spazi interni*/
        var ixiban=1;
        
        while (ixiban > 0) 
        {
            Valore=Valore.replace(" ","");
            ixiban=Valore.indexOf(" ");
        }
        
        if (g.length==0)
            return true;
        else
            return IsBIC(Valore);
    }
    
    if (Tipo=="TE")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        else
            return true;
    }

    if (Tipo=="CFO")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        Tipo="CF";
    }
    
    if (Tipo=="CF")
    {
    
        g=trim(Valore);
        if (g.length==0)
            return true;
        else
        {    
            if (ControllaCF(g)=="")
                return true;
            return false;    
        }    
    }
    if (Tipo=="PIO")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        Tipo="PI";
    }

    
    if (Tipo=="PI")
    {
        g=trim(Valore);
        if (g.length==0)
            return true;
        else
        {    
            if (ControllaPIVA(g)=="")
                return true;
            return false;    
        }    
    }

    if (Tipo=="PGO")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        Tipo="PG";
    }

    
    if (Tipo=="PG")
    {
        g=trim(Valore);
        if (g.length==0)
            return true;
        else
        {    
            if (ControllaPIVAGiur(g)=="")
                return true;
            return false;    
        }    
    }

    if (Tipo=="CA")
    {
        g=trim(Valore);
        if (g.length==0)
            return true;
            
        if (g.length!=5)
            return false;
        
        var RegEx=/[0-9]{5}/
        if (!RegEx.test(g))
           return false;

        return true;
    }

    if (Tipo=="CL")
    {
        g=trim(Valore);
        if (g.length==0)
            return true;
        
        var RegEx=/^(\+39){0,1}[0-9]{3}[0-9]{6,7}$/
        if (!RegEx.test(g))
           return false;

        return true;
    }
    
    if (Tipo=="IN")
    return IsInteger(Valore);
    
    /* formato atteso gg/mm/aaaa */
    if (Tipo=="DTO")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        Tipo="DT"
    }
    
    if (Tipo=="DT")
    {
        g=trim(Valore);
        if (g.length==0)
            return true;
        if (Valore.indexOf("/")<=0)
        {
           if (g.length==8) 
           {
              Valore=Valore.substring(0,2) + "/" + Valore.substring(2,4) + "/" + Valore.substring(4,8);
              RetValue=Valore;
           }
        }
        
        
        var RegEx=/^\d{2}\/\d{2}\/\d{4}$/
        if (!RegEx.test(Valore))
           return false;
           
        Dati = Valore.split("/");
        
        g=Dati[0];
        m=Dati[1];
        a=Dati[2];

        return VerificaData(g,m,a);
    }

    if (Tipo=="DA")
    {
        
        Dati = Valore.split("/");
        
        if (Dati.length==3)  
        {
            g=Dati[0];
            m=Dati[1];
            a=Dati[2];
        }
        else
        {
            g=Valore.substring(6,8);
            m=Valore.substring(4,6);
            a=Valore.substring(0,4);
            
        }
        return VerificaData(g,m,a);
    }

    if (Tipo=="AN")
    {
        g=trim(Valore);
        if (g.length==0)
            return false;
        else
        {
            var espressione = /[^a-zA-Z0-9]/;
            if (espressione.test(g))
            {
                return false;
            }
            else
                return true;
        }        
    }
    
}
function ElaboraControlli() 
{   
    return ElaboraControlliD("S");
}
function ElaboraControlliNoMsg() 
{
    return ElaboraControlliD("N");
}



function ElaboraControlliD(FlagShow) 
{
         var s,t,f,val,tipo,c,ra,f1,f2,f1s,f2s,rn,Flag;
         var Id = new Array();
         var Nomi = new Array();
         var Dati = new Array();
         var Num1 = new Number();
         var Num2 = new Number();         
         var Num3 = new Number();
         var Num4 = new Number();         

         Flag=0;
         s=ValoreDi("DescLoaded");
         Id = s.split(";");
            
         for (i=0;i<Id.length;i++) 
         {
            t=Id[i];
            
            if (t.length!=0) 
            {
            
                t =ValoreDi("NameLoaded");
                ra=ValoreDi("NameRangeD");
                rn=ValoreDi("NameRangeN");

                Nomi = t.split(";");        
                for (j=0;j<Nomi.length;j++) 
                {
                    f=trim(Nomi[j]);

                    if(f.length!=0)
                    {
                        Dati = f.split(",");
                        c=Dati[0] + Id[i];

                        tipo=Dati[1];

                        if (tipo=="DA")
                        {
                            val=""
                            val=val+ValoreDi("AA_"+c);
                            val=val+ValoreDi("MM_"+c);
                            val=val+ValoreDi("GG_"+c);
                            //alert(val);
                            if (GetColoreFocus("GG_"+c)=="yellow")
                            {
                               xx=ImpostaColoreFocus("GG_"+c,"","white");    
                            }
                            
                            if (CheckTipo(val,tipo)==false)
                            {
                               xx=ImpostaColoreFocus("GG_"+c,"S","yellow");    
                               Flag=1;
                            }
                        }

                        if (CheckTipoValido(tipo)==true)
                        {

                            val=ValoreDi(c);
                            if (GetColoreFocus(c)=="yellow")
                            {
                               xx=ImpostaColoreFocus(c,"","white");    
                            }
                            
                            
                            if (CheckTipo(val,tipo)==false)
                            {
                               xx=ImpostaColoreFocus(c,"S","yellow");        
                               Flag=1;
                            }
                            else
                            {
                                if (tipo=="DT"  && RetValue!=ValoreDi(c) && RetValue!="" )
                                   xx=ImpostaValoreDi(c,RetValue);
                                if (tipo=="DTO" && RetValue!=ValoreDi(c) && RetValue!="" )
                                   xx=ImpostaValoreDi(c,RetValue);
                            }
                        }

                    }
                }
                
                if (Flag==1)
                {
                    if (FlagShow=="S")
                        alert("Dati non validi");
                   return false;
                }

                Nomi = ra.split(";");        
                for (j=0;j<Nomi.length;j+=2) 
                {
                    f1=trim(Nomi[j]);
                    
                    if(f1.length!=0)
                    {
                        f2=trim(Nomi[j+1]);
                        c=f1 + Id[i];
                        f1s=""
                        f1s=f1s+ValoreDi("AA_"+c);
                        f1s=f1s+ValoreDi("MM_"+c);
                        f1s=f1s+ValoreDi("GG_"+c);

                        c=f2 + Id[i];
                        f2s=""
                        f2s=f2s+ValoreDi("AA_"+c);
                        f2s=f2s+ValoreDi("MM_"+c);
                        f2s=f2s+ValoreDi("GG_"+c);
                    
                        if (f2s < f1s)
                        {
                            xx=ImpostaColoreFocus("GG_"+c,"S","");
                            alert("Range non valido");
                            return false;
                        }
                    }
                }

                Nomi = rn.split(";");    
                if (Nomi.length>3)
                {

                    for (j=0;j<Nomi.length;j+=4) 
                    {
                        
                        f1=trim(Nomi[j+0]);

                        c=f1 + Id[i];
                        f1s=ValoreDi(c);
                        Num1 = GetNumberAsFloat(f1s);

                        f2=trim(Nomi[j+1]);

                        c=f2 + Id[i];
                        f1s=ValoreDi(c);
                        Num2 = GetNumberAsFloat(f1s);

                        f1s=trim(Nomi[j+2]);
                        Num3 = GetNumberAsFloat(f1s);

                        f1s=trim(Nomi[j+3]);
                        Num4 = GetNumberAsFloat(f1s);

                        if (Num1 > Num2)
                        {
                            f1=trim(Nomi[j+0]);
                            c=f1 + Id[i];
                            xx=ImpostaColoreFocus(c,"S","");
                            alert("Range non valido :" + Num1 + ">" + Num2);
                            return false;
                        }
                
                        if (Num3>-1)
                        {
                            if (Num1 < Num3)
                            {
                                f1=trim(Nomi[j+0]);
                                c=f1 + Id[i];
                                xx=ImpostaColoreFocus(c,"S","");        
                                alert("Range non valido :" + Num1 + " minore del minimo previsto " + Num3);
                                return false;
                            }
                        }    

                        if (Num4>-1)
                        {
                            if (Num2 > Num4)
                            {
                                f1=trim(Nomi[j+1]);
                                c=f1 + Id[i];
                                xx=ImpostaColoreFocus(c,"S","");
                                alert("Range non valido :" + Num2 + " maggiore del massimo previsto " + Num4);
                                return false;
                            }
                        }    
                    }
                }
            }
         }
         
                                    
         return true;

 }

function ResetCampo(Campo)
{
    if (GetColoreFocus(Campo)=="yellow")
    {
       xx=ImpostaColoreFocus(Campo,"","white");    
    }
}
 
function ControllaCampo(Campo,tipo)
{
    if (GetColoreFocus(Campo)=="yellow")
    {
       xx=ImpostaColoreFocus(Campo,"","white");    
    }
                            
    val=ValoreDi(Campo);    
    if (CheckTipo(val,tipo)==false)
    {
        xx=ImpostaColoreFocus(Campo,"","yellow");
        return false;
   }    
    if (tipo=="DT"  && RetValue!=val && RetValue!="" )
        xx=ImpostaValoreDi(Campo,RetValue);
    if (tipo=="DTO" && RetValue!=val && RetValue!="" )
        xx=ImpostaValoreDi(Campo,RetValue);
    return true;
} 
 
 
function CheckNew() 
{
    var Msg="";
    
    return CheckNewAll("S",Msg) ;
} 

function CheckNewNoMsg(Msg) 
{
    return CheckNewAll("N",Msg) ;
} 

function CheckNewAll(FlagMsg,Msg) 
{
    
         var t,f,val,tipo,c,ra,rn,flag,j;
         var Id,f1,f2;
         var Nomi = new Array();
         var Dati = new Array();
         var Num1 = new Number();
         var Num2 = new Number();         
         var Num3 = new Number();
         var Num4 = new Number();         
         Id = "_new";
        
        
        t =ValoreDi("NameLoaded");
        ra=ValoreDi("NameRangeD");
        rn=ValoreDi("NameRangeN");

        Nomi = t.split(";");        

        flag=0;        
        for (j=0;j<Nomi.length;j++) 
        {
            f=trim(Nomi[j]);

            if(f.length!=0)
            {
                
                Dati = f.split(",");
                c=Dati[0] + Id;
                                            
                tipo=Dati[1];

                if (tipo=="DA")
                {
                    val=""
                    val=val+ValoreDi("AA_"+c);
                    val=val+ValoreDi("MM_"+c);
                    val=val+ValoreDi("GG_"+c);
                    //alert(val);
                    if (GetColoreFocus("GG_"+c)=="yellow")
                    {
                        xx=ImpostaColoreFocus("GG_"+c,"S","white");
                       
                    }
                    if (CheckTipo(val,tipo)==false)
                    {
                       xx=ImpostaColoreFocus("GG_"+c,"S","yellow");
                       flag=1;
                    }
                }

                if (tipo=="RA")
                {
                    if (IsRadioChecked(c)==false)
                    {
                       flag=1;
                    }
                }
                if (CheckTipoValido(tipo)==true)
            
                {
                    val=ValoreDi(c);
                    if (GetColoreFocus(c)=="yellow")
                    {
                       xx=ImpostaColoreFocus(c,"N","white");
                    }

                    if (CheckTipo(val,tipo)==false)
                    {
                        xx=ImpostaColoreFocus(c,"S","yellow");
                        flag=1;
                    }
                    else
                    {
                    
                    if (tipo=="DT"  && RetValue!=ValoreDi(c) && RetValue!="" )
                       xx=ImpostaValoreDi(c,RetValue);
                    if (tipo=="DTO" && RetValue!=ValoreDi(c) && RetValue!="" )
                       xx=ImpostaValoreDi(c,RetValue);
                    }
                    
                }

            }
        }

        if (flag==1)
        {
           if (FlagMsg=="N")
           {
              Msg="Dati non validi";
           }
           if (FlagMsg=="S")
           {
              alert("Dati non validi");
           }
           return false;
        }

        Nomi = ra.split(";");        
        for (j=0;j<Nomi.length;j+=2) 
        {

            f1=trim(Nomi[j]);
            if(f1.length!=0)
            {
            
                f2=trim(Nomi[j+1]);

                c=f1 + Id;
                f1s=""
                f1s=f1s+ValoreDi("AA_"+c);
                f1s=f1s+ValoreDi("MM_"+c);
                f1s=f1s+ValoreDi("GG_"+c);

                c=f2 + Id;
                f2s=""
                f2s=f2s+ValoreDi("AA_"+c);
                f2s=f2s+ValoreDi("MM_"+c);
                f2s=f2s+ValoreDi("GG_"+c);
                if (f2s < f1s)
                {
                   xx=ImpostaColoreFocus("GG_"+c,"S","yellow");
                   flag=1;

                }
            }
        }

        if (flag==1)
        {
           if (FlagMsg=="N")
           {
              Msg="Range Date non valido";
           }
           if (FlagMsg=="S")
           {
              alert("Range Date non valido");
           }
    
           return false;
        }
        
        Nomi = rn.split(";");        
        if (Nomi.length>3)
        {
            for (j=0;j<Nomi.length;j+=4) 
            {
                f1=trim(Nomi[j+0]);
                
                c=f1 + Id;
                f1s=ValoreDi(c);
                Num1 = GetNumberAsFloat(f1s);

                f2=trim(Nomi[j+1]);
                c=f2 + Id;
                f1s=ValoreDi(c);
                Num2 = GetNumberAsFloat(f1s);

                f1s=trim(Nomi[j+2]);
                Num3 = GetNumberAsFloat(f1s);

                f1s=trim(Nomi[j+3]);
                Num4 = GetNumberAsFloat(f1s);

                
                if (Num1 > Num2)
                {
                    f1=trim(Nomi[j+0]);
                    c=f1 + Id;
                    xx=ImpostaColoreFocus(c,"S","");
                    
                    if (FlagMsg=="N")
                    {
                      Msg="Range non valido :" + Num1 + ">" + Num2;
                    }
                    if (FlagMsg=="S")
                    {
                      alert("Range non valido :" + Num1 + ">" + Num2);
                    }

                    return false;
                }
                
                if (Num3>-1)
                {
                    if (Num1 < Num3)
                    {
                        f1=trim(Nomi[j+0]);
                        c=f1 + Id;
                        xx=ImpostaColoreFocus(c,"S","");
                        
                        if (FlagMsg=="N")
                        {
                          Msg="Range non valido :" + Num1 + " minore del minimo previsto " + Num3;
                        }
                        if (FlagMsg=="S")
                        {
                          alert("Range non valido :" + Num1 + " minore del minimo previsto " + Num3);
                        }

                        return false;
                    }
                }    

                if (Num4>-1)
                {
                    if (Num2 > Num4)
                    {
                        f1=trim(Nomi[j+1]);
                        c=f1 + Id;
                        xx=ImpostaColoreFocus(c,"S","");
                        
                        if (FlagMsg=="N")
                        {
                          Msg="Range non valido :" + Num2 + " maggiore del massimo previsto " + Num4;
                        }
                        if (FlagMsg=="S")
                        {
                          alert("Range non valido :" + Num2 + " maggiore del massimo previsto " + Num4);
                        }
                        return false;
                    }
                }    
                
            }
        }
  return true;

 }
 
 function Registra()
{
    if (ElaboraControlli()==true)
    {
        xx=ImpostaValoreDi("Oper","UPD");
        document.Fdati.submit();
    }
}

function ModificaNew()
{
    if (CheckNew()==true)
    {
        xx=ImpostaValoreDi("Oper","M");
        document.Fdati.submit();
    }
}

 function Cancella()
{
    xx=ImpostaValoreDi("Oper","D");
    document.Fdati.submit();
}

 function Sottometti()
{
    xx=ImpostaValoreDi("Oper","");
    document.Fdati.submit();
}

function RegistraNew()
{
    if (CheckNew()==true)
    {
        xx=ImpostaValoreDi("Oper","INS");
        document.Fdati.submit();
    }
}
function Gestisci(Id)
{
        document.Fdati.action=ValoreDi("PageToCall");
        xx=ImpostaValoreDi("IdParm",Id);
        document.Fdati.submit();
}

function Paginazione(Id)
{
        xx=ImpostaValoreDi("Oper","P");
        xx=ImpostaValoreDi("Pagina",Id);
        document.Fdati.submit();
}

function CheckCF(Valore,vuoto)
{
    g=trim(Valore);
    if (g.length==0)
    {
        if (vuoto==0)
        return false;
    }
            
    if (ControllaCF(g)=='')
        return true;
    return false;    
}

function ControllaCF(cf)
{
    var validi, i, s, set1, set2, setpari, setdisp;
    
    if( cf == '' )  return '';
    cf = cf.toUpperCase();
    
    if( cf.length != 16 )
        return "Lunghezza codice fiscale non corretta:\n"
        +"il codice fiscale dovrebbe essere lungo\n"
        +"16 caratteri.\n";
    validi = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    for( i = 0; i < 16; i++ ){
        if( validi.indexOf( cf.charAt(i) ) == -1 )
            return "Il codice fiscale contiene un carattere non valido `" +
                cf.charAt(i) +
                "'.\nI caratteri validi sono le lettere e le cifre.\n";
    }
    set1 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    set2 = "ABCDEFGHIJABCDEFGHIJKLMNOPQRSTUVWXYZ";
    setpari = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    setdisp = "BAKPLCQDREVOSFTGUHMINJWZYX";
    s = 0;
    for( i = 1; i <= 13; i += 2 )
        s += setpari.indexOf( set2.charAt( set1.indexOf( cf.charAt(i) )));
    for( i = 0; i <= 14; i += 2 )
        s += setdisp.indexOf( set2.charAt( set1.indexOf( cf.charAt(i) )));
    if( s%26 != cf.charCodeAt(15)-'A'.charCodeAt(0) )
        return "Codice fiscale non corretto:\n"+
            "il codice di controllo non corrisponde.\n";
    return "";
}

function CheckPI(Valore,vuoto)
{
    g=trim(Valore);
    if (g.length==0)
    {
        if (vuoto==0)
        return false;
    }
            
    if (ControllaPIVA(g)=='')
        return true;
    return false;    
}

function ControllaPIVA(pi)
{
    var i,s,c,validi;
    if( pi == '' )  return '';
    if( pi.length != 11 )
        return "Lunghezza partita IVA non \n" +
            "corretta: la partita IVA dovrebbe essere lunga\n" +
            "11 caratteri.\n";
    validi = "0123456789";
    for( i = 0; i < 11; i++ ){
        if( validi.indexOf( pi.charAt(i) ) == -1 )
            return "La partita IVA contiene un carattere non valido `" +
                pi.charAt(i) + "'.\nI caratteri validi sono le cifre.\n";
    }
    s = 0;
    for( i = 0; i <= 9; i += 2 )
        s += pi.charCodeAt(i) - '0'.charCodeAt(0);
    for( i = 1; i <= 9; i += 2 ){
        c = 2*( pi.charCodeAt(i) - '0'.charCodeAt(0) );
        if( c > 9 )  c = c - 9;
        s += c;
    }
    if( ( 10 - s%10 )%10 != pi.charCodeAt(10) - '0'.charCodeAt(0) )
        return "Partita IVA non valida:\n" +
            "il codice di controllo non corrisponde.\n";
    return '';
}

function CheckPIGiur(Valore,vuoto)
{
    g=trim(Valore);
    if (g.length==0)
    {
        if (vuoto==0)
        return false;
    }
            
    if (ControllaPIVAGiur(g)=='')
        return true;
    return false;    
}

function ControllaPIVAGiur(pi)
{
    var i,s,c,validi;
    if( pi == '' )  return '';
    if( pi.length != 11 )
        return "Lunghezza partita IVA non \n" +
            "corretta: la partita IVA dovrebbe essere lunga\n" +
            "11 caratteri.\n";
    validi = "0123456789";
    for( i = 0; i < 11; i++ ){
        if( validi.indexOf( pi.charAt(i) ) == -1 )
            return "La partita IVA contiene un carattere non valido `" +
                pi.charAt(i) + "'.\nI caratteri validi sono le cifre.\n";
    }
    return '';
}

function ShowHideArea(strAreaId,strImg)
{
    var oCollAreaDiv = document.getElementById(strAreaId);
    var oCollAreaImg = document.getElementById(strImg);    
    if (oCollAreaDiv.style.display == 'none')
    {
        oCollAreaDiv.style.display = 'block';
        oCollAreaImg.src = imgMinus;
    }
    else
    {        
        oCollAreaDiv.style.display = 'none';
        oCollAreaImg.src = imgPlus;    
    }

}
function ShowHideAreaNoImg(strAreaId)
{
    var oCollAreaDiv = document.getElementById(strAreaId);

    if (oCollAreaDiv.style.display == 'none')
    {
        oCollAreaDiv.style.display = 'block';
    }
    else
    {        
        oCollAreaDiv.style.display = 'none';
    }

}

function CheckRangeDT(minDate) 
{
         var s,t,f,val,ra,f1,f2,f1s,f2s,rn,j;
         var Id = new Array();
         var Nomi = new Array();
         var Dati = new Array();
 
         s=ValoreDi("DescLoaded");
                  
         Id = s.split(";");
                                     
         for (i=0;i<Id.length;i++) 
         {
            t=Id[i];
            if (t.length!=0) 
            {
            
                ra=ValoreDi("NameRangeDT");
                
                Nomi = ra.split(";");        
                for (j=0;j<Nomi.length;j+=2) 
                {
                    f1=trim(Nomi[j]);
                    
                    if(f1.length!=0)
                    {
                        xx=ImpostaColoreFocus(f1+t,"S","");
                        xx=ImpostaColoreFocus(trim(Nomi[j+1])+t,"S","");

                        f2=ValoreDi(f1+t);
                        f1s=""
                        f1s=f1s+f2.substring(6,10);
                        f1s=f1s+f2.substring(3,5);
                        f1s=f1s+f2.substring(0,2);

                        f2 =ValoreDi(trim(Nomi[j+1])+t);
                        f2s=""
                        f2s=f2s+f2.substring(6,10);
                        f2s=f2s+f2.substring(3,5);
                        f2s=f2s+f2.substring(0,2);

                        if (f2s < f1s)
                        {
                            xx=ImpostaColoreFocus(f1+t,"S","yellow");
                            xx=ImpostaColoreFocus(trim(Nomi[j+1])+t,"S","yellow");
                            alert("Range non valido");
                            return false;
                        }
                        if (f1s < minDate)
                        {
                            xx=ImpostaColoreFocus(f1+t,"S","yellow");
                            alert("Data minore del minimo previsto");
                            return false;
                        }                        
                    }
                }
            }
         }
         
         
         return true;

 }
 
  function CheckTipoValido(T)
 {
    // 3 carattere O = obbligatorio - Z = Positivo e zero  - P = Positivo - Q = Percentuale
     if (T=="IN" || T=="INO" || T=="INP" || T=="INZ" || T=="INQ")
        return true;
     if (T=="FL" || T=="FLO" || T=="FLP" || T=="FLZ" || T=="FLQ")
        return true;
    if (T=="TE" || T=="LI"  || T=="EM"  || T=="EMO")
       return true;
    if (T=="CF" || T=="CFO" || T=="PI"  || T=="PIO" || T=="PG"  || T=="PGO")
        return true;
    if (T=="DT" || T=="DTO" || T=="CA"  || T=="CAO") 
        return true;
    if (T=="AN" || T=="ANO") 
        return true;
    if (T=="CL" || T=="CLO") 
        return true;
    /*IBAN Italiano*/
    if (T=="IB" || T=="IBO") 
        return true;
    /*IBAN generico*/
    if (T=="WB" || T=="WBO") 
        return true;
    /*Codice BIC Banca*/
    if (T=="BI" || T=="BIO") 
        return true;
    /*ML  CONTROLLO LUNGHEZZA MAX LEN*/
    if (T.substring(0,2)=="ML" ) 
        return true;    
    return false;
 }



function CheckRangeDate(Dt1,Dt2) 
{
    var f1s,f2s;
    f2=Dt1;
    f1s=DateToString(Dt1);
    f2s=DateToString(Dt2);

    
    if (f2s < f1s)
        return false;
    return true;

 }
 
function DateToString(Dt)
{
    var f2,f1s;
    f2=Dt;
    f1s=""
    if(f2.length==10)
    {
        f1s=f1s+f2.substr(6,4);
        f1s=f1s+f2.substr(3,2);
        f1s=f1s+f2.substr(0,2);
   }
    else
    {
        f1s=f1s+f2.substr(4,4);
        f1s=f1s+f2.substr(2,2);
        f1s=f1s+f2.substr(0,2);
   }
    return f1s;
 }

function ismaxlength(obj)
{
    var mlength=obj.getAttribute? parseInt(obj.getAttribute("maxlength")) : ""
    if (obj.getAttribute && obj.value.length>mlength)
    obj.value=obj.value.substring(0,mlength)
}

function estraiVocali (str) {
  return str.replace(/[^AEIOU]/gi, '');
}

function estraiConsonanti (str) {
  return str.replace(/[^BCDFGHJKLMNPQRSTVWXYZ]/gi, '');
}
 
//-->