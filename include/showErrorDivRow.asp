
<% if MsgErrore<>"" then %> 
<div class="row  center " >
   <div class="col-12 bg-danger text-white">
      <p class="font-weight-bold"><%=server.htmlencode(MsgErrore)%></p>
   </div>
</div> 
<%end if %>