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
categoria.Source = "SELECT * FROM pgpacotes where id = '"& request("categoria")&"'"
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
<%
Dim menu
Dim menu_numRows

Set menu = Server.CreateObject("ADODB.Recordset")
menu.ActiveConnection = MM_dlfelix_STRING
menu.Source = "SELECT * FROM categoria order by titulo asc"
menu.CursorType = 0
menu.CursorLocation = 2
menu.LockType = 1
menu.Open()

menu_numRows = 0
%>

<%
Dim Repeat23__numRows
Dim Repeat23__index

Repeat23__numRows = 100
Repeat23__index = 0
menu_numRows = menu_numRows + Repeat23__numRows
%>


<!-------------------------------------->
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=latin1	" />
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
                <li><a href="suites.asp">Su�tes</a>
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
    

 <!-- end .slideshow-caption -->

<!-- inicio .slideshow item -->        
	<li id="rslides1_s1" style="float: none; position: absolute; opacity: 0; z-index: 1; display: list-item; 
    transition: opacity 500ms ease-in-out; background-image: url(<%=(atualiza.Fields.Item("img2").Value)%>);" name="rslides1_s1">
	<img src="<%=(atualiza.Fields.Item("img2").Value)%>" alt="">
	<div class="slideshow-caption">
	<h1 class="sombra"><%=(atualiza.Fields.Item("titulo2").Value)%>

	<h3><%=(atualiza.Fields.Item("resumo2").Value)%></h3>  		
	</div>
	</li>
    

 <!-- end .slideshow-caption -->
    
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
			<p class="pre-title">Pousada My Power</p>
			<h2>Confira abaixo<span> o nosso </span> <span><%=(categoria.Fields.Item("pactitulo").Value)%></span></h2>
			<hr>
			<ul class="guests-items">
				
               
                
                <!--GUEST ITEM-->
				
				 	
                     <p class="pct-preco">2 Pessoas sem Ar Condicionado - <%=(categoria.Fields.Item("2psap").Value)%></p>
              <p class="pct-obs"><%=(categoria.Fields.Item("2psa").Value)%></p>
              <p class="pct-preco">..................................................................</p>
              
                     <p class="pct-preco">2 Pessoas com Ar Condicionado - <%=(categoria.Fields.Item("2pcap").Value)%></p>
                     <p class="pct-obs"><%=(categoria.Fields.Item("2pca").Value)%></p>
                     <p class="pct-preco">..................................................................</p>
                     <p class="pct-preco">3 Pessoas sem Ar Condicionado - <%=(categoria.Fields.Item("3psap").Value)%></p>
              <p class="pct-obs"><%=(categoria.Fields.Item("3psa").Value)%></p>
              <p class="pct-preco">..................................................................</p>
              
                     <p class="pct-preco">3 Pessoas com Ar Condicionado - <%=(categoria.Fields.Item("3pcap").Value)%></p>
                     <p class="pct-obs"><%=(categoria.Fields.Item("3pca").Value)%></p>
                     <p class="pct-preco">..................................................................</p>
                     <p class="pct-preco">4 Pessoas sem Ar Condicionado - <%=(categoria.Fields.Item("4psap").Value)%></p>
              <p class="pct-obs"><%=(categoria.Fields.Item("4psa").Value)%></p>
              <p class="pct-preco">..................................................................</p>
              
                     <p class="pct-preco">4 Pessoas com Ar Condicionado - <%=(categoria.Fields.Item("4pcap").Value)%></p>
                     <p class="pct-obs"><%=(categoria.Fields.Item("4pca").Value)%></p>
                     <p class="pct-preco">..................................................................</p>
                     <p class="pct-preco">6 Pessoas com Ar Condicionado - <%=(categoria.Fields.Item("6pcap").Value)%></p>
                     <p class="pct-obs"><%=(categoria.Fields.Item("6pca").Value)%></p>
                     <p class="pct-preco">..................................................................</p>
                     <p class="pct-preco">10 Pessoas com Ar Condicionado - <%=(categoria.Fields.Item("10pcap").Value)%></p>
                     <p class="pct-obs"><%=(categoria.Fields.Item("10pca").Value)%></p>
                     <p class="pct-preco">..................................................................</p>
                     <p class="pct-preco">16 Pessoas com Ar Condicionado - <%=(categoria.Fields.Item("16pcap").Value)%></p>
                     <p class="pct-obs"><%=(categoria.Fields.Item("16pca").Value)%></p>
                     <p class="pct-preco">..................................................................</p>
                     <p class="pct-preco">Promo&Ccedil;&Atilde;O 1 - <%=(categoria.Fields.Item("promocaop1").Value)%></p>
                     <p class="pct-obs"><%=(categoria.Fields.Item("promocao1").Value)%></p>
                     <p class="pct-preco">..................................................................</p>
                     <p class="pct-preco">Promo&Ccedil;&Atilde;o 2 - <%=(categoria.Fields.Item("promocaop2").Value)%></p>
                     <p class="pct-obs"><%=(categoria.Fields.Item("promocao2").Value)%></p>
                     <p class="pct-preco">..................................................................</p>
                    <p class="pct-obs">Observa&Ccedil;&Otilde;es e requisitos Gerais do Pacote:</p>
                     <p class="pct-obs"><%=(categoria.Fields.Item("obs1").Value)%></p>
                     
              
                   

                            
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

	
		  <div class="cleaner"></div>
		</section>
	</div>

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
<h6><span>Pousada My Power - Bertioga -</span> AOL - Sistemhost - Printsystem</h6>
				<ul class="social-links small">
					<li class="fb"><a href="#"></a></li>
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
