<!--#include virtual="/gscVirtual/common/function.asp"-->
<!--#include virtual="/gscVirtual/common/functionNew.asp"-->
<!--#include virtual="/gscVirtual/common/connDb.asp"-->
<!--#include virtual="/gscVirtual/common/parametri.asp"-->

<!DOCTYPE html>
<html lang="en">

<head>
	<!-- Required meta tags -->
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
	<title>Skydash Admin</title>
	<!-- plugins:css -->
	<link rel="stylesheet" href="vendors/feather/feather.css">
	<link rel="stylesheet" href="vendors/ti-icons/css/themify-icons.css">
	<link rel="stylesheet" href="vendors/css/vendor.bundle.base.css">
	<!-- endinject -->
	<!-- Plugin css for this page -->
	<!-- End plugin css for this page -->
	<!-- inject:css -->
	<link rel="stylesheet" href="css/vertical-layout-light/style.css">
	<!-- endinject -->
	<link rel="shortcut icon" href="images/favicon.png" />
  </head>	
	
<script language="JavaScript">
function Entra()
{
	g=document.getElementById("email").value;
	if (g.length==0)
	{
		return false;
	}		
	g=document.getElementById("password").value;
	if (g.length==0)
	{
		return false;
	}		
	document.Fdati.submit();

}

function Recupera()
{
	
	g=document.getElementById("email").value;
	if (g.length==0)
	{
		alert("Inserire l'indirizzo mail per il recupero !")
		return false;
	}		
	
	ImpostaValoreDi("Oper","Recupera");
	alert("qui");
	document.Fdati.submit();

}

</script>

<body class="text-center">
	<div class="container-scroller">
		<div class="container-fluid page-body-wrapper full-page-wrapper">
		  	<div class="content-wrapper d-flex align-items-center auth px-0">
				<div class="row w-100 mx-0">
			  		<div class="col-lg-4 mx-auto">
						<div class="auth-form-light text-left py-5 px-4 px-sm-5">
							<div class="brand-logo text-center">
								<img src="img/logo.png" alt="logo">
							</div>   
							<form class="form-signin" name="Fdati" id="loginform" action="HomeAreaRiservata.asp" method="POST">
								<input type = "hidden" name="Oper" id="Oper" value = "">	
								<div class="row row-cols-12">&nbsp;</div>
								<%
								IdAzienda= cdbl("0" & Request("IdAzienda"))
								if Cdbl(IdAzienda)=0 then 
								IdAzienda=1
								end if 
								esito = rtrim(request("esito"))
								if esito<>"" then 
								%>
								<div class = "card-small red">
									<div class = "row">
										<div class = "col s12">
											<h4 class="icon-site-color"><%=esito%></h4>
										</div>
									</div>
								</div>
								<%
								end if 
								%>	
								<div class="row p-b-30">
									<div class="col-12">
										<div class="input-group mb-3">
												<%
													Query="Select * from Azienda Order By DescAzienda"
													response.write ListaDbChangeCompleta (Query,"IdAzienda",IdAzienda,"IdAzienda","DescAzienda",1,"cambiaIcona()","IdAzienda","","","","class='form-control form-control-lg'")
													%>
										</div>		
										<div class="form-group">
											<input type="email" placeholder="Email" class="form-control form-control-lg" name="email" id="email"  value="<%=Request("email")%>" aria-label="Email" aria-describedby="basic-addon1" required="">
										</div>
										<div class="input-group mb-3">
											<input type="password" placeholder="Password" class="form-control form-control-lg" name="password" id="password" value="<%=Request("password")%>" aria-label="Password" aria-describedby="basic-addon2" required="">
										</div>
										<div class="row ">
											<div class="col-12">
												<div class="form-group">
													<div class="p-t-20">
														<button class="btn btn-info"    type="button" onclick="Recupera();"></i>Ricordami Password</button>
														<button class="btn btn-success float-right" type="button" onclick="Entra();">Login</button>
													</div>
												</div>
											</div>
										</div>
									</div>
								</div>
							</form>
						</div>
			  		</div>
				</div>
  		  	</div>
		</div>
	</div>
<!--  Scripts-->
<!--#include virtual="/gscVirtual/include/scripts.asp"-->  

</body>

</html>