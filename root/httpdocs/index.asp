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
			<p class="pre-title">Pousada My Power Bertioga</p>
			<h2>Lazer <span>e na</span> <span>My Power</span></h2>
			<hr>
			<ul class="guests-items">
				
                <%	
						
				
		if not rec2.eof then %>
                
                <!--GUEST ITEM-->
				<li>
					<div class="circle">
						<figure class="imgLiquidFill">
							<img src="<%=(atualiza.Fields.Item("img4").Value)%>" alt="guest-item">
						</figure>
						<div class="fill"></div>
						<ul class="social-links medium">
						<li class="fb"><a href="#"></a></li>
						<%	end if


 %>
                            
							
						</ul>
					</div>
					<h5><%=(atualiza.Fields.Item("titulo4").Value)%></h5>
					<h6><%=(atualiza.Fields.Item("resumo4").Value)%></h6>
					<a href="lazer.asp?categoria=<%=(atualiza.Fields.Item("link4").Value)%>" class="button ruby">Confira!</a>
			  </li>
				<!--/GUEST ITEM-->

				<li>
					<div class="circle">
						<figure class="imgLiquidFill">
							<img src="<%=(atualiza.Fields.Item("img5").Value)%>" alt="guest-item">
						</figure>
						<div class="fill"></div>
						<ul class="social-links medium">
							<li class="fb"><a href="#"></a></li>
							
							
						</ul>
					</div>
					<h5><%=(atualiza.Fields.Item("titulo5").Value)%></h5>
					<h6><%=(atualiza.Fields.Item("resumo5").Value)%></h6>
					<a href="lazer.asp?categoria=<%=(atualiza.Fields.Item("link5").Value)%>" class="button ruby">Confira!</a>
				</li>
				<!--/GUEST ITEM-->

				<!--GUEST ITEM-->
				<li>
					<div class="circle">
						<figure class="imgLiquidFill">
							<img src="<%=(atualiza.Fields.Item("img6").Value)%>" alt="guest-item">
						</figure>
						<div class="fill"></div>
						<ul class="social-links medium">
							<li class="fb"><a href="#"></a></li>
							
							
						</ul>
					</div>
					<h5><%=(atualiza.Fields.Item("titulo6").Value)%></h5>
					<h6><%=(atualiza.Fields.Item("resumo6").Value)%></h6>
					<a href="lazer.asp?categoria=<%=(atualiza.Fields.Item("link6").Value)%>" class="button ruby">Confira!</a>
				</li>
				<!--/GUEST ITEM-->

				<!--GUEST ITEM-->
				<li>
					<div class="circle">
						<figure class="imgLiquidFill">
							<img src="<%=(atualiza.Fields.Item("img7").Value)%>" alt="guest-item">
						</figure>
						<div class="fill"></div>
						<ul class="social-links medium">
							<li class="fb"><a href="#"></a></li>
							
							
						</ul>
					</div>
					<h5><%=(atualiza.Fields.Item("titulo7").Value)%></h5>
					<h6><%=(atualiza.Fields.Item("resumo7").Value)%></h6>
					<a href="lazer.asp?categoria=<%=(atualiza.Fields.Item("link7").Value)%>" class="button ruby">Confira!</a>
				</li>
				<!--/GUEST ITEM-->

			
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

	<!--BLOG NEWS-->
	<div id="blog-news-wrap">
		<section id="blog-news">
			<p class="pre-title">Venha conhecer nossa pousada</p>
			<h2>Nossas   <span>Acomoda&ccedil;&otilde;es em destaque</span></h2>
			<hr>
			<!--POSTS-->
			<ul class="posts small">
				<!--POST-->
				<li>
					<figure class="imgLiquidFill"><img src="<%=(atualiza.Fields.Item("img8").Value)%>" alt="post-img"></figure>
					<article>
					  <h5><span class="pre-title"><%=(atualiza.Fields.Item("titulo8").Value)%></span></h5>
					  <br>

						
					  <p><%=(atualiza.Fields.Item("resumo8").Value)%></p>
						<a href="produtos.asp?categoria=<%=(atualiza.Fields.Item("link8").Value)%>" class="button ruby">Confira!</a>
					</article>
				</li>
				<!--/POST-->

				<!--POST-->
				<li>
					<figure class="imgLiquidFill"><img src="<%=(atualiza.Fields.Item("img9").Value)%>" alt="post-img">
					</figure>
					<article>
					  <h5><span class="pre-title"><%=(atualiza.Fields.Item("titulo9").Value)%></span></h5>
					  <br>
                      <p><%=(atualiza.Fields.Item("resumo9").Value)%></p>
						<a href="produtos.asp?categoria=<%=(atualiza.Fields.Item("link9").Value)%>" class="button ruby">Confira!</a>
					</article>
				</li>
				<!--/POST-->

				<!--POST-->
				<li>
					<figure class="imgLiquidFill"><img src="<%=(atualiza.Fields.Item("img10").Value)%>" alt="post-img"></figure>
					<article>
					  <h5><span class="pre-title"><%=(atualiza.Fields.Item("titulo10").Value)%></span></h5>
					  <br>
                      <p><%=(atualiza.Fields.Item("resumo10").Value)%></p>
						<a href="produtos.asp?categoria=<%=(atualiza.Fields.Item("link10").Value)%>" class="button ruby">Confira!</a>
					</article>
				</li>
				<!--/POST-->
			</ul>
			<!--/POSTS-->
		  <div class="cleaner"></div>
		</section>
	</div>
	<!--/BLOG NEWS-->
<!--TICKETS-->
	<div id="tickets-wrap">
		<section id="tickets">
			<p class="pre-title">O melhor pre&ccedil;o da regi&atilde;o</p>
			<h2>APROVEITE <span>NOSSAS	</span> <span>PROMO&Ccedil;&Otilde;ES</span></h2>
			<hr>
			<ul class="ticket-items">
				<!--TICKET ITEM-->
				<li class="small">
					<!--CORNERS-->
					<div class="corner lt"></div>
					<div class="corner rt"></div>
					<div class="corner lb"></div>
					<div class="corner rb"></div>
					<!--/CORNERS-->

					<!--HEXAGONS-->
					<div class="hexa outer ruby">
						<div></div>
						<div></div>
						<div></div>
					</div>
					<div class="hexa inner dark-blue">
						<div>2</div>
						<div></div>
						<div></div>
					</div>
					<!--/HEXAGONS-->
				  <h3><%=(atualiza2.Fields.Item("pactitulo2").Value)%><span></span></h3>
					<h class="titulo5"><%=(atualiza2.Fields.Item("pacpreco2").Value)%></h>
					<img src="images/ribbon-small01.png" alt="ribbon-s">
					<img src="images/ribbon-medium01.png" alt="ribbon-m">
					<article>
						
						<p><%=(atualiza2.Fields.Item("pacresumo2").Value)%></p>
				  </article>
					<a href="pacotes.asp?categoria=<%=(atualiza2.Fields.Item("paclink2").Value)%>" class="button ruby">Confira!</a>
				</li>
				<!--/TICKET ITEM-->

				<!--TICKET ITEM-->
				<li class="medium">
					<!--CORNERS-->
					<div class="corner lt"></div>
					<div class="corner rt"></div>
					<div class="corner lb"></div>
					<div class="corner rb"></div>
					<!--/CORNERS-->

					<!--HEXAGONS-->
					<div class="hexa outer dark-blue">
						<div></div>
						<div></div>
						<div></div>
					</div>
					<div class="hexa inner ruby">
						<div>1</div>
						<div></div>
						<div></div>
					</div>
					<!--/HEXAGONS-->

				  <h3><%=(atualiza2.Fields.Item("pactitulo1").Value)%><span></span></h3>
					<h class="titulo6"><%=(atualiza2.Fields.Item("pacpreco1").Value)%></h>
					<img src="images/ribbon-medium02.png" alt="ribbon-m">
					<img src="images/ribbon-small02.png" alt="ribbon-s">
					<article>
						<p><%=(atualiza2.Fields.Item("pacresumo1").Value)%></p>
				  </article>
					<a href="pacotes.asp?categoria=<%=(atualiza2.Fields.Item("paclink1").Value)%>" class="button ruby">Confira!</a>
				</li>
				<!--TICKET ITEM-->

				<!--TICKET ITEM-->
				<li class="small">
					<!--CORNERS-->
					<div class="corner lt"></div>
					<div class="corner rt"></div>
					<div class="corner lb"></div>
					<div class="corner rb"></div>
					<!--/CORNERS-->

					<!--HEXAGONS-->
					<div class="hexa outer ruby">
						<div></div>
						<div></div>
						<div></div>
					</div>
					<div class="hexa inner dark-blue">
						<div>3</div>
						<div></div>
						<div></div>
					</div>
					<!--/HEXAGONS-->
					<h3><%=(atualiza2.Fields.Item("pactitulo3").Value)%><span></span></h3>
					<h class="titulo5"><%=(atualiza2.Fields.Item("pacpreco3").Value)%></h>
					<img src="images/ribbon-medium01.png" alt="ribbon-m">
					<img src="images/ribbon-small01.png" alt="ribbon-s">
					<article>
						<p><%=(atualiza2.Fields.Item("pacresumo3").Value)%></p>
					</article>
					<a href="pacotes.asp?categoria=<%=(atualiza2.Fields.Item("paclink3").Value)%>" class="button ruby">Confira!</a>
				</li>
				<!--TICKET ITEM-->
			</ul>
			<div class="cleaner"></div>
		</section>
	</div>
	<!--/TICKETS-->
	<!--CONTACT-->
	<div id="contact-wrap">
	  <table width="100%" border="0" cellspacing="30">
	    <tr>
	      <td align="center"><p class="pre-title">Pousada bertioga</p>
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
			</section>
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
