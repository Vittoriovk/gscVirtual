<%
  NomePagina="popup.asp"
  titolo="popup Ramo"
  default_check_profile="SuperV"
%>
<!--#include virtual="/gscVirtual/include/includeStdCheck.asp"-->
<%
  livelloPagina="00"
%>
<!--#include virtual="/gscVirtual/include/utility.asp"-->
<!DOCTYPE html>
<html>
  <head>
    <title><%= titolo %></title>
    <!-- Inserire qui eventuali metadati o fogli di stile -->
  </head>
  <script language="JavaScript">

    function localFun(Op,Id)
    {
        xx=ImpostaValoreDi("DescLoaded","0");
        xx=ElaboraControlli();
        
         if (xx==false)
           return false;
        
        ImpostaValoreDi("Oper","update");
        document.Fdati.submit();
    }
    
    </script>
<body>
<!-- Set Rs,MsgErrore,NameRangeD,NameRangeN,NameLoaded,DescLoaded,UsaPaginazione=false,SavePaginazione=false -->
<!--#include virtual="/gscVirtual/include/setupParm.asp"-->
<!--#include virtual="/gscVirtual/include/GetPagSize.asp"-->
    <script>
        function localSubmit(Op)
        {
        var xx;
            xx=false;
            if (Op=="submit")
               xx=ElaboraControlli();

             if (xx==false)
               return false;

            ImpostaValoreDi("Oper","update");
            document.Fdati.submit();
        }
    </script>
        <%
        NameLoaded= ""
        NameLoaded= NameLoaded & "DescRamo,TE"

        FirstLoad=(Request("CallingPage")<>NomePagina)
        IdRamo=0
        if FirstLoad then
          IdRamo   = "0" & Session("swap_IdRamo")
          if Cdbl(IdRamo)=0 then
            IdRamo = cdbl("0" & getValueOfDic(Pagedic,"IdRamo"))
          end if
          OperTabella   = Session("swap_OperTabella")
          PaginaReturn    = getCurrentValueFor("PaginaReturn")
        else
          IdRamo   = "0" & getValueOfDic(Pagedic,"IdRamo")
          OperTabella   = getValueOfDic(Pagedic,"OperTabella")
          PaginaReturn  = getValueOfDic(Pagedic,"PaginaReturn")
        end if
        IdRamo = cdbl(IdRamo)
        'inizio elaborazione pagina

        Dim DizDatabase
        Set DizDatabase = CreateObject("Scripting.Dictionary")

        xx=SetDiz(DizDatabase,"IdRamo",0)
        xx=SetDiz(DizDatabase,"DescRamo","")

        'recupero i dati
        if cdbl(IdRamo)>0 then
          MySql = ""
          MySql = MySql & " Select * From  Ramo "
          MySql = MySql & " Where IdRamo=" & IdRamo
          xx=GetInfoRecordset(DizDatabase,MySql)
        end if

        'inserisco il fornitore
        descD  = Request("DescRamo0")

        if Oper=ucase("update") and OperTabella="CALL_INS" then

          Session("TimeStamp")=TimePage
          KK="0"
          MyQ = ""
          MyQ = MyQ & " INSERT INTO Ramo (DescRamo,IdAnagRamo) "
          MyQ = MyQ & " values ('" & apici(descD) & "','" & apici(request("IdAnagRamo0")) & "')"

          ConnMsde.execute MyQ
          If Err.Number <> 0 Then
            MsgErrore = ErroreDb(Err.description)
          else
            response.redirect virtualpath & PaginaReturn
          End If
        end if

        'registro i dati della pagina
        xx=setValueOfDic(Pagedic,"IdRamo"  ,IdRamo)
        xx=setValueOfDic(Pagedic,"OperTabella"  ,OperTabella)
        xx=setValueOfDic(Pagedic,"PaginaReturn" ,PaginaReturn)
        xx=setCurrent(NomePagina,livelloPagina)

        DescLoaded="0"
      %>

      <%
        'xx=DumpDic(SessionDic,NomePagina)
      %>
  <div>
    <dialog class="finestra-dial" id="myFirstDialog"
        style="width:30%; border-radius:0.5em;  border: solid 0.05em black;">
        <div class="row">
            <div class="col-lg-12  stretch-card">
                <div class="card">
                    <div class="card-body text-center">
                        <div class="icons-list2 pb-4 pr-2">
                            <div class="float-right">
                                <a id="hide" style="cursor: pointer;"><i class="mdi mdi-close-circle-outline" onmouseover="mouseover(this)" onmouseout="mouseout(this)"></i></a>
                            </div>
                        </div>
                        <h2>Creazione ramo</h2>
                    </div>
                    <div class="float-left d-block">
                        <form class="form w-75 bar-loader" method="post" id="update-form">
                            <div class="form-group">
                                <%
                                kk="DescRamo"
                                xx=ShowLabel("Descrizione Ramo")
                                NameLoaded= NameLoaded & kk & ",TE;"
                                %>
                                <input type="text" class="form-control" Id="<%=KK%>0" name="<%=KK%>0" value="<%=GetDiz(DizDatabase,"DescRamo") %>" >
                            </div>
                            <%if false then %>
                            <div class="row">
                                <div class="col-6">
                                  <div class="form-group ">
                                    <%
                                    kk="IdAnagRamo"
                                    xx=ShowLabel("Ramo di riferimento")
                                    NameLoaded= NameLoaded & kk & ",LI;"
                                    q = ""
                                    q = q & " select * from AnagRamo "
                                    q = q & " where IdAnagRamo not in (select IdAnagRamoPadre from AnagRamo)"
                                    q = q & " order by descAnagRamo"
                                    stdClass="class='form-control form-control-sm'"
                                    response.write ListaDbChangeCompleta(q,"IdAnagRamo0",GetDiz(DizDatabase,"IdAnagRamo") ,"IdAnagRamo","DescAnagRamo" ,tt,"","","","","",stdClass)
                                    %>
                                  </div>
				                        </div>
                              </div>
			                        <%end if %>
                              <br>
                              <!--#include virtual="/gscVirtual/include/showErrorDivRow.asp"-->
                              <%if SoloLettura=false then%>
                              <div class="mx-auto">
                                  <%RiferimentoA="center;#;;2;save;Registra; Registra;localFun('submit','0');S"%>
                                  <!--#include virtual="/gscVirtual/include/Anchor.asp"-->
                              </div>
                            </div>
                              <%end if %>
			                        <!--#include virtual="/gscVirtual/include/CampiHidden.asp"-->
                          </form>
                          <button id="submit" class="add btn btn-primary todo-list-add-btn float-right mt-4 mb-2 mr-5">Aggiungi</button>
                      </div>
                  </div>
              </div>
          </div>
      </dialog>
  </div>

  <script>
    var myDialog = document.querySelector('#myFirstDialog');
    myDialog.addEventListener('click', function(event) {
    if(event.target === this) {
        this.close();
    }
  });

function closeDialog() {
  var myDialog = document.querySelector('#myFirstDialog');
  if (myDialog.open) {
    myDialog.close();
  }
}

</script>
<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scriptsAll.asp"-->

</body>
</html>
