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
Dim orkut
Dim orkut_numRows

Set orkut = Server.CreateObject("ADODB.Recordset")
orkut.ActiveConnection = MM_dlfelix_STRING
orkut.Source = "SELECT * FROM fotos where categoria = '"& request("categoria")&"'"
orkut.CursorType = 0
orkut.CursorLocation = 2
orkut.LockType = 1
orkut.Open()

orkut_numRows = 0
%>
<%
Dim HLooper1__numRows
HLooper1__numRows = 2
Dim HLooper1__index
HLooper1__index = 0
orkut_numRows = orkut_numRows + HLooper1__numRows
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 2
Repeat1__index = 0
orkut_numRows = orkut_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
orkut_total = orkut.RecordCount

' set the number of rows displayed on this page
If (orkut_numRows < 0) Then
  orkut_numRows = orkut_total
Elseif (orkut_numRows = 0) Then
  orkut_numRows = 1
End If

' set the first and last displayed record
orkut_first = 1
orkut_last  = orkut_first + orkut_numRows - 1

' if we have the correct record count, check the other stats
If (orkut_total <> -1) Then
  If (orkut_first > orkut_total) Then orkut_first = orkut_total
  If (orkut_last > orkut_total) Then orkut_last = orkut_total
  If (orkut_numRows > orkut_total) Then orkut_numRows = orkut_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (orkut_total = -1) Then

  ' count the total records by iterating through the recordset
  orkut_total=0
  While (Not orkut.EOF)
    orkut_total = orkut_total + 1
    orkut.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (orkut.CursorType > 0) Then
    orkut.MoveFirst
  Else
    orkut.Requery
  End If

  ' set the number of rows displayed on this page
  If (orkut_numRows < 0 Or orkut_numRows > orkut_total) Then
    orkut_numRows = orkut_total
  End If

  ' set the first and last displayed record
  orkut_first = 1
  orkut_last = orkut_first + orkut_numRows - 1
  If (orkut_first > orkut_total) Then orkut_first = orkut_total
  If (orkut_last > orkut_total) Then orkut_last = orkut_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = orkut
MM_rsCount   = orkut_total
MM_size      = orkut_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
orkut_first = MM_offset + 1
orkut_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (orkut_first > MM_rsCount) Then orkut_first = MM_rsCount
  If (orkut_last > MM_rsCount) Then orkut_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_dlfelix_STRING
Recordset1.Source = "SELECT * FROM categoria where id = '"& request("categoria")&"'"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
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
			<h2>Confira abaixo <span><%=(Recordset1.Fields.Item("titulo").Value)%></span></h2>
			<hr>
			<ul class="guests-items">
				
               
                
                <!--GUEST ITEM-->
                <tr>
            <td class="laranja"><p class="pct-obs"><%=(recordset1.Fields.Item("texto").Value)%></p></td>
          </tr><tr>
            <td class="laranja"><div align="left">&nbsp;</div></td>
          </tr><tr>
            <td class="laranja"><div align="left">&nbsp;</div></td>
          </tr>
                
                <td width="417" valign="top" class="style37"><div align="left">
                <table width="191" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
                  <% If Not orkut.EOF Or Not orkut.BOF Then %>
                  <%
startrw = 0
endrw = HLooper1__index
numberColumns = 4
numrows = 3
while((numrows <> 0) AND (Not orkut.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
                  <tr>
                    <%
While ((startrw <= endrw) AND (Not orkut.EOF))
%>
                    
                    <td><img src="<%=(orkut.Fields.Item("img").Value)%>" width="200" height="150" border="0" class="borda" />
                        <table border="0" width="100%" cellspacing="0" cellpadding="0">
                          <tr>
                            <td><%=(orkut.Fields.Item("codigo").Value)%></td>
                          </tr>
                      </table></td>
                    <td width="10">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                    <%
	startrw = startrw + 1
	orkut.MoveNext()
	Wend
	%>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                  </tr>
                  <%
 numrows=numrows-1
 Wend
 %>
                  <% End If ' end Not orkut.EOF Or NOT orkut.BOF %>
                </table>
            </div></td>
          </tr>
          <tr>
            <td class="laranja"></td>
          </tr>
          
          <tr>
            <td class="laranja"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
                <tr>
                  <td height="18"><font size="1">P&aacute;gina:&nbsp; <span class="vinho">
                    <%
TM_counter = 0
For i = 1 to orkut_total Step MM_size
TM_counter = TM_counter + 1
TM_PageEndCount = i + MM_size - 1
if TM_PageEndCount > orkut_total Then TM_PageEndCount = orkut_total
if i <> MM_offset + 1 then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(TM_counter & "</a>")
else
Response.Write("<b>" & TM_counter & "</b>")
End if
if(TM_PageEndCount <> orkut_total) then Response.Write("&nbsp;|&nbsp;")
next
 %>
                  </a></span></font></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><div align="center" class="vermelho"><a href="javascript:history.go(-1);" class="vermelho">&laquo; voltar</a></div>
                <div align="center" class="laranja"></div>
              <div align="right"></div></td>
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
