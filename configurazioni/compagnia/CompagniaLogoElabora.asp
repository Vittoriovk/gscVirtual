   <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
   <%
   DescLoaded="0"
   l_Id = "0"
   %>


   <%if logoPiccolo<>"" then
        logoPiccolo = replace((VirtualPath & DirectoryUpload) & logoPiccolo,"//","/")
   %>
   
   <div class="row" >
	   <div class="col-2">
  <p class="font-weight-bold">Logo Piccolo</p>
	   </div>
	   <div class = "col-10">
	   <img src="<%=(logoPiccolo) %>" class="img-fluid" alt="Logo Piccolo">
   </div>
   <div class="col-2">
      <p class="font-weight-bold"> </p>
   </div>   
   </div> 

   <%end if %>
   
   
<div class="row" >

   <div class="col-2">
      <p class="font-weight-bold">Logo Piccolo</p>
   </div>
   <div class = "col-8">
   
  <div class="form-group">
    <input type="file" class="form-control-file border" id="FileP0" name="FileP0" >
  </div>   

     </div>
   <div class="col-2">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;localSubmit1();S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->   
      <p class="font-weight-bold"> </p>
   </div>

</div> 
       

  <%if logoGrande<>"" then
        logoGrande = replace((VirtualPath & DirectoryUpload) & logoGrande,"//","/")
   %>
   
   <div class="row" >
	   <div class="col-2">
  <p class="font-weight-bold">Logo Grande</p>
	   </div>
	   <div class = "col-10">
	   <img src="<%=(logoGrande) %>" class="img-fluid" alt="Logo Piccolo">
   </div>
   <div class="col-2">
      <p class="font-weight-bold"> </p>
   </div>   
   </div> 

   <%end if %>
   
   
<div class="row" >

   <div class="col-2">
      <p class="font-weight-bold">Logo Grande</p>
   </div>
   <div class = "col-8">
   
  <div class="form-group">
    <input type="file" class="form-control-file border" id="FileG0" name="FileG0" >
  </div>   

     </div>
   <div class="col-2">
		<%RiferimentoA="center;#;;2;save;Registra; Registra;localSubmit2();S"%>
		<!--#include virtual="/gscVirtual/include/Anchor.asp"-->   
      <p class="font-weight-bold"> </p>
   </div>

</div> 	   
 
   
   