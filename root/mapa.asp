<!DOCTYPE html>
<html> <head>
<!-- Google Tag Manager -->
<script>(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
'https://www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
})(window,document,'script','dataLayer','GTM-T9ND2HR');</script>
<!-- End Google Tag Manager -->

<!-- Inicio Programa��o topo ------------------------------------>
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

<!-- Google Tag Manager (noscript) -->
<noscript><iframe src="https://www.googletagmanager.com/ns.html?id=GTM-T9ND2HR"
height="0" width="0" style="display:none;visibility:hidden"></iframe></noscript>
<!-- End Google Tag Manager�(noscript)�-->

<!-- MENU TOPO ------>
<!--NAVIGATION-->
<div id="nav-wrap">
	  <nav id="main-nav">
			
		<ul>
				<li><a href="index.asp">Home</a></li>
				<li><a href="empresa.asp"> A POUSADA</a></li>
                <li><a href="suites.asp">Su&iacute;tes</a>
                <li><a href="reservas.asp"> Pacotes</a></li>
				<li><a href="mapa.asp">Como Chegar</a></li>
				
				<li><a href="contato.asp">Contato</a></li>
                <li><a href="#"></a></li>
                <li><a href="#"></a></li>
			</ul>
      <a href="#" id="pull"> Menu&nbsp;&nbsp;</a></nav>
	</div>
	<!--/NAVIGATION-->		

	<!-- BANNER PROMO��ES ------------------------------------------------------------------>
	 
     <link rel="stylesheet" type="text/css" href="./banner_files/style.css" media="screen">
     <div id="apDiv1"><img src="images/logotipo.png" class="sombra"></div>
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
<!--NEWS-->
	<div id="news-wrap">
		<section id="guests">
			<p class="pre-title"><%=(atualiza1.Fields.Item("nomeempresa").Value)%></p>
			<h2>Confira <span>como</span> <span>Chegar</span></h2>
			<hr>
			<ul class="guests-items">
			            
<!--MAPA-->

                 <% if not rec2.eof then %>
                 
                  <div class="aaa" >
                  
                  <iframe src=" <%=(atualiza1.Fields.Item("maps").Value)%> " width="100%" height="450" frameborder="0" style="border:0" allowfullscreen></iframe>                      
                      

                  <% end if %>  
<!--/MAPA-->
	
    </ul>
			<div class="cleaner"></div>
		</section>
	</div>
    <!--SPONSORS-->
	<div id="sponsors-wrap">
		<section id="sponsors">
			<ul class="sponsors-items">
				<li>
					<!-- <img src="images/sponsor-01.png" alt="sponsor1">
					<img src="images/sponsor-02.png" alt="sponsor2">
					<img src="images/sponsor-03.png" alt="sponsor3">
				</li>
				<li>
					<img src="images/sponsor-04.png" alt="sponsor4">
					<img src="images/sponsor-05.png" alt="sponsor5">
					<img src="images/sponsor-06.png" alt="sponsor6"> 
                    -->
				</li>
			</ul>
			<div class="cleaner"></div>
		</section>
	</div>
	<!--/SPONSORS-->
	<!--CONTACT-->
	<div id="contact-wrap">
	  <table width="100%" border="0" cellspacing="30">
	    <tr>
	      <td align="center"><p class="pre-title">Pousada My Power - Bertioga</p>
	        <h2>Fale conosco<span></span></h2>
	        <hr>
	        <table width="600" border="0" cellspacing="10">
	          <tr>
	            <td width="60%" rowspan="3"><form action="formmail.asp" method="post" name="form1" id="form1"class="form-full-width contact-form">
	              <input type="hidden" name="_redirect" value="contato2.asp" />
	              <input type="hidden" name="_subject" value="CONTATO SITE" />
	              <input type="hidden" name="_recipients" value="mypower@pousadamypower.com.br" />
	              <div class="row">
	                <div class="col-xs-12 col-sm-12">
	                  <div class="form-group"></div>
                    </div>
	                <div class="col-xs-12 col-sm-12">
	                  <div class="form-group">
	                    <input placeholder="NOME" type="text" id="contact-subject1" name="Nome:" />
                      </div>
                    </div>
	                <div class="col-xs-12 col-sm-12">
	                  <div class="form-group">
	                    <input placeholder="TELEFONE" type="text" id="contact-subject2" name="Telefone:" />
                      </div>
                    </div>
	                <div class="col-xs-12 col-sm-12">
	                  <div class="form-group">
	                    <input placeholder="E-MAIL" type="text" id="contact-subject3" name="E-Mail" />
                      </div>
                    </div>
	                <div class="col-xs-12 col-sm-12">
	                  <div class="form-group">
	                    <textarea placeholder=" MENSAGEM" id="contact-message" name="Mensagem" ></textarea>
                      </div>
                    </div>
	                <div class="col-xs-12 col-sm-12 text-left">
	                  <div class="wrap-main">
	                    <input type="submit" name="submit" id="submit" value="enviar" class="button ruby" style="float:right; width:100%; padding:5px;">
                      </div>
                    </div>
                  </div>
	              </form></td>
	            <td valign="top"><img src="images/about-item01.png" alt="ctc-option"></td>
	            <td valign="top"><h6>Local</h6>
	              <p><%=(atualiza1.Fields.Item("endereco").Value)%></p>
	              <p></td>
              </tr>
	          <tr>
	            <td valign="top"><img src="images/contact-item02.png" alt="ctc-option"></td>
	            <td valign="top"><h6>Telefone</h6>
	              <p><%=(atualiza1.Fields.Item("telefone1").Value)%></p>
	              <p><%=(atualiza1.Fields.Item("telefone2").Value)%>                  
	              <p><%=(atualiza1.Fields.Item("telefone3").Value)%></td>
              </tr>
	          <tr>
	            <td width="9%" valign="top"><ul class="contact-items">
	              <img src="images/contact-item03.png" alt="ctc-option">
	              </ul></td>
	            <td width="31%" valign="top"><h6> Email</h6>
	              <p> <%=(atualiza1.Fields.Item("email").Value)%> </td>
              </tr>
            </table></td>
        </tr>
      </table>
	</div>
	<!--/CONTACT-->
	<!--FOOTER-->
	<footer>
		
        
		<div id="footer-bottom-wrap">
			<section id="footer-bottom">
<h6><span>Pousada My Power  &copy;</span> 2015 Todos os direitos reservados / AOL - Sistemhost - Printsystem</h6>
				<ul class="social-links small">
					<a class="soboo soboo-simple" data-hover="counter" data-model="circle" data-orientation="horizontal" data-services="twitter facebook favoritos googleplus  google email mais" data-userid="15963"  id="vertical"></a>
        <script type='text/javascript' src="http://www.seomaster.com.br/js/soboo2.js"></script>
		</div>
	</footer>
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
