   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
   <%
  
   LeggiDati=false
   if Cdbl(IdUpload)>0 then
      err.clear 
      LeggiDati=true
      Set RsRec = Server.CreateObject("ADODB.Recordset")
      MySql = "" 
      MySql = MySql & " Select * from Upload Where IdUpload = " & IdUpload

      RsRec.CursorLocation = 3
      RsRec.Open MySql, ConnMsde 

      If Err.number<>0 then	
       	 LeggiDati=false
      elseIf RsRec.EOF then	
         LeggiDati=false
		 RsRec.close 
      End if
   end if   
 
   NameLoaded= ""
   NameLoaded= NameLoaded & "IdTipoDocumento,LI" 
   NameLoaded= NameLoaded & ";DescBreve,TE"  
   if FlagDescEstesa="S" then 
      NameLoaded= NameLoaded & ";DescEstesa,TE"  
   end if 
   If ShowValidoDal then 
      NameLoaded= NameLoaded & ";ValidoDal,DTO"  
   end if 
   if LeggiDati=false then 
      NameLoaded= NameLoaded & ";FileIn,TE"
   end if  
   DescLoaded="0"
   
   l_Id = "0"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->
   <%
   if LeggiDati then 
       ValoC = RsRec("IdTipoDocumento")   
   else
      ValoC = Request("IdTipoDocumento" & l_Id) 
   end if   
   ao_lbd = "Tipo Documento"                       'descrizione label 
   ao_nid = "IdTipoDocumento" & l_Id              'nome ed id
   ao_val = ValoC 'valore di default
   
   ao_Tex = "SELECT * From Documento Where IdDocumentoInterno='' order By DescDocumento"
   'response.write ao_tex
   ao_ids = "IdDocumento"				  'valore della select 
   ao_des = "DescDocumento"              'valore del testo da mostrare 
   ao_cla = ""                        'azzero classe
   ao_Eve = ""                        'azzero evento
   ao_Att = "1"                       'indica se deve mettere vuoto 
   ao_Plh = ""                        'indica cosa mettere in caso di vuoto
   ao_Cla = "class='form-control form-control-sm'"
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddList.asp"-->

   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			      
   <%
   if LeggiDati then 
       ValoC = RsRec("DescBreve")   
   else
      ValoC = Request("DescBreve" & l_Id) 
   end if
   ao_lbd = "Descrizione Breve"       'descrizione label 
   ao_nid = "DescBreve" & l_Id            'nome ed id
   ao_val = "|value=" & ValoC       'valore di default
   ao_Plh = "|placeholder=Descr.Breve"              'placeholder
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		

   <%if FlagDescEstesa="S" then %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			      
   <%
   if LeggiDati then 
       ValoC = RsRec("DescEstesa")   
   else
      ValoC = Request("DescEstesa" & l_Id) 
   end if
   ao_lbd = "Descrizione estesa"       'descrizione label 
   ao_nid = "DescEstesa" & l_Id            'nome ed id
   ao_val = "|value=" & ValoC       'valore di default
   ao_Plh = "|placeholder=Descr.Estesa"              'placeholder
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAdd.asp"-->		   
   <%else%>
      <input type="hidden" name="DescEstesa<%=l_Id%>" id="DescEstesa<%=l_Id%>" value="">
   <%end if %>
   
   
   <% If ShowValidoDal then%>
   
   <!--#include virtual="/gscVirtual/include/ParmAddObjSetFormControl.asp"-->			
   <%
   if LeggiDati then 
      ValoC = RsRec("ValidoDal")   
   else
      ValoC = Request("ValidoDal" & l_Id) 
   end if   
   ao_lbd = "Valido Dal"         'descrizione label
   ao_3ls = "col-6"                       'size terzo elemento	
   ao_div = "col-4"	   
   ao_nid = "ValidoDal" & l_Id            'nome ed id
   ao_val = ""       'valore di default
   if len(ValoC)<>8 then 
      ValoC=DtoS()
   end if 
   ValoC=Stod(ValoC)
   ao_val = ValoC 
   %>
   <!--#include virtual="/gscVirtual/include/ParmAddObjAddDate.asp"--> 
   
   
   <% end if  %>

   <%if LeggiDati=true then%>
   
   <div class="row" >
	   <div class="col-2">
		  <p class="font-weight-bold">File Caricato</p>
	   </div>
	   <div class = "col-8">
	   <input value="<%=RsRec("NomeDocumento")%>" type="text" READONLY class="form-control"  >
	   
	   </div>
   
   <div class="col-2">
      <p class="font-weight-bold"> </p>
   </div>  
   </div> 

   <%end if %>
   
   
<div class="row" >

   <div class="col-2">
      <p class="font-weight-bold">File Da Caricare</p>
   </div>
   <div class = "col-8">
   
  <div class="custom-file">
    <input type="file" id="FileIn0" name="FileIn0"  aria-describedby="inputGroupFileAddon01">
   </div>
     </div>
   <div class="col-2">
      <p class="font-weight-bold"> </p>
   </div>

</div> 
        
   
	
   <%if SoloLettura=false then%>
		<div class="row"><div class="mx-auto">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;localSubmit('submit','0');S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->			
		</div></div>
		<div class="row">
			<div class="col">
				&nbsp;
			</div>
		</div>
   <%end if %>
   
   