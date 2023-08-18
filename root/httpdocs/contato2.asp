<!DOCTYPE html>
<html> <head>
<!-- Inicio Programação topo ------------------------------------>
<!--#include file="adm/Connections/dlfelix.asp" -->

<%	

	Set rec2 = Server.CreateObject("ADODB.Recordset")
	
	
	

x = 1 %>

<%
Dim atualiza
Dim atualiza_numRows

Set atualiza = Server.CreateObject("ADODB.Recordset")
atualiza.ActiveConnection = MM_dlfelix_STRING
atualiza.Source = "SELECT *  FROM home"
atualiza.CursorType = 0
atualiza.CursorLocation = 2
atualiza.LockType = 1
atualiza.Open()

atualiza_numRows = 0
%>
<%
Dim categoria
Dim categoria_numRows

Set categoria = Server.CreateObject("ADODB.Recordset")
categoria.ActiveConnection = MM_dlfelix_STRING
categoria.Source = "SELECT * FROM categorias where id = '"& request("categoria")&"'"
categoria.CursorType = 0
categoria.CursorLocation = 2
categoria.LockType = 1
categoria.Open()

categoria_numRows = 0
%>
<%
Dim atualiza1
Dim atualiza1_numRows

Set atualiza1 = Server.CreateObject("ADODB.Recordset")
atualiza1.ActiveConnection = MM_dlfelix_STRING
atualiza1.Source = "SELECT *  FROM empresa"
atualiza1.CursorType = 0
atualiza1.CursorLocation = 2
atualiza1.LockType = 1
atualiza1.Open()

atualiza1_numRows = 0
%>
<%
Dim atualiza2
Dim atualiza2_numRows

Set atualiza2 = Server.CreateObject("ADODB.Recordset")
atualiza2.ActiveConnection = MM_dlfelix_STRING
atualiza2.Source = "SELECT *  FROM pacotes"
atualiza2.CursorType = 0
atualiza2.CursorLocation = 2
atualiza2.LockType = 1
atualiza2.Open()

atualiza2_numRows = 0
%>


<!-------------------------------------->
<head>
	 
    <meta http-equiv="Content-Type" content="text/html; charset=latin1" />
    
	<meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0">
	<title><%=(atualiza1.Fields.Item("nomeempresa").Value)%></title>
	<link rel="stylesheet" href="css/style.css">
    <link rel="stylesheet" href="css/xmcountdown.css">
	<!-- favicon -->
	<link rel="icon" href="favicon.ico">
 
</head>
<body>



<!-- MENU TOPO ------>
<!--NAVIGATION-->
<div id="nav-wrap">
	  <nav id="main-nav">
			
		<ul>
				<li><a href="index.asp">Home</a></li>
				<li><a href="empresa.asp"> A POUSADA</a></li>
                <li><a href="suites.asp">Su&iacute;tes</a>
                <li><a href="reservas.asp"> Reservas</a></li>
				<li><a href="mapa.asp">Como Chegar</a></li>
				<li><a href="eventos.asp">Eventos</a></li>
				
				<li><a href="contato.asp">Contato</a></li>
			</ul>
      <a href="#" id="pull"> Menu&nbsp;&nbsp;</a></nav>
	</div>
	<!--/NAVIGATION-->	

	<!-- BANNER PROMOÇÕES ------------------------------------------------------------------>
	 
     <link rel="stylesheet" type="text/css" href="./banner_files/style.css" media="screen">
     <div id="apDiv1"><img src="images/logotipo.png"></div>
<header id="header"><div id="slideshow"><ul class="rslides rslides1">
  
  <%	
						rec2.Open "SELECT * FROM home ORDER BY link1",MM_dlfelix_STRING
				
		if not rec2.eof then %>      
        
<!-- inicio .slideshow item -->        
	<li id="rslides1_s1" style="float: none; position: absolute; opacity: 0; z-index: 1; display: list-item; 
    transition: opacity 500ms ease-in-out; background-image: url(<%=(atualiza.Fields.Item("img1").Value)%>);" name="rslides1_s1">
	<img src="<%=(atualiza.Fields.Item("img1").Value)%>" alt="">
	<div class="slideshow-caption">
	<h1 class="sombra"><%=(atualiza.Fields.Item("titulo1").Value)%>

	<h3><%=(atualiza.Fields.Item("resumo1").Value)%></h3>  		
	</div>
	</li>

<!-- inicio .slideshow item -->        
	<li id="rslides1_s1" style="float: none; position: absolute; opacity: 0; z-index: 1; display: list-item; 
    transition: opacity 500ms ease-in-out; background-image: url(<%=(atualiza.Fields.Item("img2").Value)%>);" name="rslides1_s1">
	<img src="<%=(atualiza.Fields.Item("img2").Value)%>" alt="">
	<div class="slideshow-caption">
	<h1 class="sombra"><%=(atualiza.Fields.Item("titulo2").Value)%>

	<h3><%=(atualiza.Fields.Item("resumo2").Value)%></h3>  		
	</div>
	</li>
    
<!-- inicio .slideshow item -->        
	<li id="rslides1_s1" style="float: none; position: absolute; opacity: 0; z-index: 1; display: list-item; 
    transition: opacity 500ms ease-in-out; background-image: url(<%=(atualiza.Fields.Item("img3").Value)%>);" name="rslides1_s1">
	<img src="<%=(atualiza.Fields.Item("img3").Value)%>" alt="">
	<div class="slideshow-caption">
	<h1 class="sombra"><%=(atualiza.Fields.Item("titulo3").Value)%>

	<h3><%=(atualiza.Fields.Item("resumo3").Value)%></h3>  		
	</div>
	</li>
    

 <!-- end .slideshow-caption -->		  
         
          	  
		</ul> 	
	</div>
</header>
 
<!--/HEADER-->
<!--TICKETS-->
<div id="tickets-wrap">
		<section id="tickets">
			<p class="pre-title" style="font-size:36px; color:#090;"><!--/TICKET ITEM-->e-mail enviado<br>
			  <br>
		    com sucesso!!!</p>
			<p class="pre-title" style="font-size:36px; color:#090;">&nbsp;</p>
			<p class="pre-title" style="font-size:20px; color:#090;">Em Breve entraremos em contato. Obrigado!</p>
  </section>
  </div>
	<!--/TICKETS-->
	

<!--MAP-->
	<section id="map"> </section>
	<!--/MAP-->

	<!--FOOTER-->
	<footer>
		
        
		<div id="footer-bottom-wrap">
			<section id="footer-bottom">
<h6><span>Pousada My Power  &copy;</span> 2015 Todos os direitos reservados / AOL - Sistemhost - Printsystem</h6>
				<ul class="social-links small">
					<li class="fb"><a href="https://www.facebook.com/Pousada-My-Power-Bertioga-1622682654672705/"></a></li>
					<li class="twt"><a href="#"></a></li>
	<!--/FOOTER-->

<!--jQuery-->
<script src="js/jquery-1.11.1.min.js"></script>
<!--bxSlider-->
<script src="js/jquery.bxslider.min.js"></script>
<!--XMCountdown-->
<script src="js/jquery.xmcountdown.min.js"></script>
<!--Tweet-->
<script src="js/twitter/jquery.tweet.min.js"></script>
<!--XMAccordion-->
<script src="js/jquery.xmaccordion.min.js"></script>
<!--XMTab-->
<script src="js/jquery.xmtab.min.js"></script>
<!--imgLiquid-->
<script src="js/imgLiquid-min.js"></script>
<!--Menu-->
<script src="js/menu.js"></script>
<!--Main-->
<script src="js/main.js"></script>
<!-- Contact -->
<script src="js/contact.js"></script>
<!--Google Maps API-->
<script src="https://maps.googleapis.com/maps/api/js"></script>
<!--Google Maps Config-->
<script src="js/gmaps.js"></script>
<script type="text/javascript" src="./banner_files/jquery-2.1.0.min68b3.js"></script>
<script type="text/javascript" src="./banner_files/jquery.prettyPhoto.js"></script>
<script type="text/javascript" src="./banner_files/jquery.tools.min.js"></script>
<script type="text/javascript" src="./banner_files/owl.carousel.min.js"></script>
<script type="text/javascript" src="./banner_files/jquery.nav.js"></script>
<script type="text/javascript" src="./banner_files/jquery.appear.js"></script>
<script type="text/javascript" src="./banner_files/responsiveslides.min.js"></script>
<script type="text/javascript" src="./banner_files/custom.js"></script>

<!-- fim da Pagina------------->
<%	rec2.close
	 %>
</span>
</div>
</div>
<div align="left">
<%	end if

%>
<%
categoria.Close()
Set categoria = Nothing
%>

<!-- ------------->


</body>
</html>
