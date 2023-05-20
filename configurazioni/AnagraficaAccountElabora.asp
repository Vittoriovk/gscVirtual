   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   l_Id = "0"
   
   ao_lbd = "Cognome/Rag.Sociale"       'descrizione label 
   ao_nid = "Cognome" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"Cognome")       'valore di default
   ao_Plh = "|placeholder=Cognome/Rag.Sociale"              'placeholder
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Nome/Rag.Sociale"         'descrizione label 
   ao_nid = "Nome" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"Nome")        'valore di default
   ao_Plh = "|placeholder=Nome/Rag.Sociale"      'placeholder 
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   ao_lbd = "Tipo Account"             'descrizione label 
   ao_nid = "IdTipoAccount" & l_Id              'nome ed id
   ao_val = GetDiz(DizDatabase,"IdTipoAccount") 'valore di default	
   if ucase(ao_val)="SUPERV" then
      ao_Tex = "select * from TipoAccount Where IdTipoAccount='SuperV' order By DescTipoAccount"
   elseif ucase(ao_val)="ADMIN" then	  
      ao_Tex = "select * from TipoAccount Where IdTipoAccount='Admin'  order By DescTipoAccount"
   elseif ucase(ao_val)="BACKO" then	
      ao_Tex = "select * from TipoAccount Where IdTipoAccount='BackO' order By DescTipoAccount"
   end if 
   ao_ids = "IdTipoAccount"			  'valore della select 
   ao_des = "DescTipoAccount"         'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "0"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "UserId"                  'descrizione label 
   ao_nid = "UserId" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"UserId")
   ao_Plh = "|placeholder=UserId"      'placeholder 
   ao_Att = ""
   if cdbl(v_IdAccount)>0 then 
	  ao_Att=ao_Att & "|attribute=readonly"
   end if 
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Password"                 'descrizione label 
   ao_nid = "Password" & l_Id            'nome ed id
   ao_val = "|value=" & decripta(GetDiz(DizDatabase,"Password"))        'valore di default
   ao_Plh = "|placeholder=Password"      'placeholder 
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		
 
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
   ao_lbd = "Indirizzo"                 'descrizione label 
   ao_nid = "Indirizzo1" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"Indirizzo1")        'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->	
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Indirizzo"                 'descrizione label 
   ao_nid = "Indirizzo2" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"Indirizzo2")       'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->	

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   ao_lbd = "Cap"                      'descrizione label 
   ao_lbs = "col-2"                    'size della label	
   ao_3ld = " "                        'descrizione terzo elemento
   ao_3ls = "col-7"                    'size terzo elemento					   
   ao_div = "col-3"
   ao_nid = "Cap" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"cap")       'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->	

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Comune"                 'descrizione label 
   ao_nid = "Comune" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"Comune")        'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->	

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Provincia"                 'descrizione label 
   ao_nid = "Provincia" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"Provincia")        'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->	

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "Settore"                 'descrizione label 
   ao_nid = "Settore" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"Settore")        'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->	

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "email principale"                 'descrizione label 
   ao_nid = "email1" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"email1")        'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->	
				   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "email alternativa"                 'descrizione label 
   ao_nid = "email2" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"email2")        'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->					   

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   ao_lbd = "recapito telefonico"      'descrizione label 
   ao_nid = "Telefono" & l_Id            'nome ed id
   ao_val = "|value=" & GetDiz(DizDatabase,"Telefono")        'valore di default
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->
   <%if SoloLettura=false then%>
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;localSubmit('submit');S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div></div>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
		</div>
   <%end if %>

   
   