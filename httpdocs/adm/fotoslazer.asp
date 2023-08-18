<%@LANGUAGE="VBSCRIPT"%>
<% session.lcid=1046 %>

<!--#include file="Connections/dlfelix.asp" -->

		 <%	
	Set Recordset1 = Server.CreateObject("ADODB.Recordset")
	
    Recordset1.ActiveConnection = MM_dlfelix_STRING
Recordset1.Source = "SELECT *  FROM fotoslazer where categoria='"&request("id")&"' ORDER BY id desc"
	Recordset1.CursorLocation = 3
	Recordset1.CursorType = 1
	Recordset1.PageSize = 50
	Recordset1.CacheSize = 50
	Recordset1.Open()

	Dim contadorvideos

	If request("pagina") <> "" Then
		Recordset1.AbsolutePage = request("pagina")
		pagina2 = request("pagina")
	Else
		If NOT Recordset1.EOF Then Recordset1.AbsolutePage = 1
		pagina = 1
	End If %>
	
	

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 50
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>


<%
Dim cat__MMColParam
cat__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  cat__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim cat
Dim cat_numRows

Set cat = Server.CreateObject("ADODB.Recordset")
cat.ActiveConnection = MM_dlfelix_STRING
cat.Source = "SELECT *  FROM lazer  WHERE id = " + Replace(cat__MMColParam, "'", "''") + ""
cat.CursorType = 0
cat.CursorLocation = 2
cat.LockType = 1
cat.Open()

cat_numRows = 0
%>

<!--  FIM PROGRAMAÇÃO TOPO-->

<!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml" class="">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=latin1" />
		<title>MENU ADM</title>
 <!-- Mobile Specific Metas -->
		<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
 <!-- Basic Page Needs  -->
		 
 <!-- CSS -->		
		<link href='http://fonts.googleapis.com/css?family=Lato:100,300,400,700,900' rel='stylesheet' type='text/css'>
        <link href="css/icones.css" rel="stylesheet" type="text/css">
		<link rel="stylesheet" type="text/css" href="css/bootstrap.min.css" />
		<link rel="stylesheet" type="text/css" href="css/style.css" />
		<link rel="shortcut icon" href="images/favicon.ico">
 <!-- Modernizr -->
		 
	    </head><body>
		<div class="sp-body">	<div class="body">
 <!-- Navigation Menu -->
		<div id="navigation" class="navbar dark navbar-default navbar-fixed-top " role="navigation">
		<div class="navbar-inner">
		<div class="menubar btn"><z class="icon-casa"></z></div>
		<div class="logo-alt">
		<a href="#"><img src="images/logo-mini.png" alt="logo-mini"></a>
		</div>	</div>	</div>
				<div class="vmenu-wrapper">
				<nav class="vmenu dark">
				<div id="rt-menu-wrapper">
							<!-- Sidebar Logo -->
 <div class="logo-box">	<img src="images/logo.png" alt="logo" width="132" height="205" id="logo-image"> </div>
 
				<!--#include file="menu-esquerdo.html"-->	</div>	</nav>		
				<i class="buton font-p"> &nbsp;&nbsp;&nbsp; Em caso de d&uacute;vidas, entre em contato com o Suporte clicando no botão SUPORTE SITE<span class="  font-m" ><strong></strong></span></i></div>
				<!-- /Navigation Menu -->
				<div class="site-overlay"></div>
				<!-- Main SECTION -->	
				<div class="section">
	<div class="page-wrapper">
	<div class="container">
							  <!-- page title --><!--/ page title -->
 <aol class="box">
 <div> 
<br>
<strong class="text-primary fa-2x azul">  &Xi; Fotos</strong>
<br>
<br>
<strong class="text-primary fa-2x verde">  <a href="listar_lazer.asp" class="verde"> &lt;&lt; Voltar</a></strong>
<br>
<br>
  

</div><br>
<form method="POST" action="<%=MM_editAction%>" name="form1">

	<input type="hidden" name="img5" size="19"><input type="hidden" name="img4" size="19"><input type="hidden" name="img3" size="19"><input type="hidden" name="img2" size="19"><input type="hidden" name="img1" size="19">
												<input type="hidden" name="data" size="19" value="<%=date%>">

	
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    


<body bgcolor="#E6E6E6">
<table height="27%" cellspacing="0" cellpadding="0" width="100%" border="0">
  <tr>
    <td valign="top" align="middle"><table border="0" style="border-collapse: collapse" width="741" cellpadding="0">
      <tr>
        <td width="739" align="right"></td>
      </tr>
    </table>
        <table cellspacing="0" cellpadding="0" width="775" align="center" bgcolor="" background="bg_pag.gif" border="0">
          <tr>
            <td width="775"><div align="center">
                <table cellspacing="50" cellpadding="0" width="100%" align="center" border="0">
                  <tr>
                    <td><table border="0" style="border-collapse: collapse" width="579" cellpadding="0">
                      <tr>
                        <td width="579"><span style="text-transform: uppercase"><b><font face="Verdana"><font color="#DF8600"></font> <u class="botoes"> <a href="incluir_fotolazer.asp?id=<%=request("id")%>&id1=<%=request("id1")%>&vsession=<%=request("id")%>"> <font color="#fffffff">INSERIR FOTO</font></a></u></font></b></span></td>
                      </tr>
                      <tr>
                        <td width="579">&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="579"><table border="0" style="border-collapse: collapse" width="100%">
                          <tr>
                            <td><font color="#868686">
                              <%

			
							Response.Write "<b>" & Recordset1.RecordCount & "</b> Fotos - Mostrando p&aacute;gina <b>" &_
							pagina & "</b> de <b>" & Recordset1.PageCount & "</b>"
%>
                            </font></td>
                            <td width="120"><% strProcura =	"acao=procurar"

			
							


			If pagina > 1 Then
				Response.Write "<a href=""fotoslazer.asp?" & strProcura & "&pagina=1""><img src=""imagens/First.gif"" border=""0"" alt=""Primeira""></a>&nbsp;" & Chr(13)
				Response.Write "<a href=""fotoslazer.asp?" & strProcura & "&pagina=" & (pagina - 1) & """><img src=""imagens/Previous.gif"" border=""0"" alt=""Anterior""></a>&nbsp;" & Chr(13)
			Else
				Response.Write "<img src=""imagens/First_off.gif"" border=""0"" alt=""Primeira"">&nbsp;" & Chr(13)
				Response.Write "<img src=""imagens/Previous_off.gif"" border=""0"" alt=""Anterior"">&nbsp;" & Chr(13)
			End If

			If CInt(pagina) <> CInt(Recordset1.PageCount) Then 
				Response.Write "<a href=""fotoslazer.asp?" & strProcura & "&pagina=" & (pagina + 1) & """><img src=""imagens/Next.gif"" border=""0"" alt=""Pr&oacute;xima""></a>&nbsp;" & Chr(13)
				Response.Write "<a href=""fotoslazer.asp?" & strProcura & "&pagina=" & Recordset1.PageCount & """><img src=""imagens/Last.gif"" border=""0"" alt=""&Uacute;ltima""></a>&nbsp;" & Chr(13)
			Else
				Response.Write "<img src=""imagens/Next_off.gif"" border=""0"" alt=""Pr&oacute;xima"">&nbsp;" & Chr(13)
				Response.Write "<img src=""imagens/Last_off.gif"" border=""0"" alt=""&Uacute;ltima"">&nbsp;" & Chr(13)
			End If %></td>
                          </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td width="579">&nbsp;</td>
                      </tr>
                      <tr>
                        <td background="bg_pontos.gif" width="579">&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="579"><table border="0" style="border-collapse: collapse" width="100%">
                          <tr>
                            <td>&nbsp;</td>
                            <td width="65" align="center">&nbsp;</td>
                            <td width="82" align="center">&nbsp;</td>
                          </tr>
                        </table></td>
                      </tr>
                      <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
                      <tr>
                        <td width="579"><table border="0" style="border-collapse: collapse" width="579" cellpadding="6">
                          <tr>
                            <td width="367"><p><b> <center><img src=<%=(Recordset1.Fields.Item("img").Value)%> width="262" height="219"> </center> </b></p>
                              <p><b><span class="verde">| Detalhes:&nbsp;&nbsp;</span> <span class="digit"><%=(Recordset1.Fields.Item("codigo").Value)%></span></b></p></td>
                            <td width="47" align="center"><p><a href="editar_fotolazer.asp?id3=<%=(Recordset1.Fields.Item("id").Value)%>&id=<%=request("id")%>&id1=<%=request("id1")%>"> <font color="#666666">Editar</font></a></p>
                              <p><a href="editar_fotolazer.asp?id3=<%=(Recordset1.Fields.Item("id").Value)%>&id=<%=request("id")%>&id1=<%=request("id1")%>"><img src="editar.gif" width="32" height="31" border="0"></a></p></td>
                            <td width="61" align="center"><p><a onClick="javascript:if (!confirm('Voc&ecirc; tem certeza que deseja excluir?')){return false;}" href="del_fotolazer.asp?del=<%=(Recordset1.Fields.Item("id").Value)%>&id=<%=request("id")%>&id1=<%=request("id1")%>"> <font color="#666666">Excluir</font></a></p>
                              <p><a onClick="javascript:if (!confirm('Voc&ecirc; tem certeza que deseja excluir?')){return false;}" href="del_fotolazer.asp?del=<%=(Recordset1.Fields.Item("id").Value)%>&id=<%=request("id")%>&id1=<%=request("id1")%>"><img height="30" alt="del" src="excluir.gif" width="30" border="0"></a></p></td>
                          </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td background="bg_pontos.gif" width="579">&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="579">&nbsp;</td>
                      </tr>
                      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
                      <tr>
                        <td width="579">&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="579"><table border="0" style="border-collapse: collapse" width="100%">
                          <tr>
                            <td><font color="#868686">
                              <%

			
							Response.Write "<b>" & Recordset1.RecordCount & "</b> Fotos - Mostrando p&aacute;gina <b>" &_
							pagina & "</b> de <b>" & Recordset1.PageCount & "</b>"
%>
                            </font></td>
                            <td width="120"><% strProcura =	"acao=procurar"

			
							


			If pagina > 1 Then
				Response.Write "<a href=""fotoslazer.asp?" & strProcura & "&pagina=1""><img src=""imagens/First.gif"" border=""0"" alt=""Primeira""></a>&nbsp;" & Chr(13)
				Response.Write "<a href=""fotoslazer.asp?" & strProcura & "&pagina=" & (pagina - 1) & """><img src=""imagens/Previous.gif"" border=""0"" alt=""Anterior""></a>&nbsp;" & Chr(13)
			Else
				Response.Write "<img src=""imagens/First_off.gif"" border=""0"" alt=""Primeira"">&nbsp;" & Chr(13)
				Response.Write "<img src=""imagens/Previous_off.gif"" border=""0"" alt=""Anterior"">&nbsp;" & Chr(13)
			End If

			If CInt(pagina) <> CInt(Recordset1.PageCount) Then 
				Response.Write "<a href=""fotoslazer.asp?" & strProcura & "&pagina=" & (pagina + 1) & """><img src=""imagens/Next.gif"" border=""0"" alt=""Pr&oacute;xima""></a>&nbsp;" & Chr(13)
				Response.Write "<a href=""fotoslazer.asp?" & strProcura & "&pagina=" & Recordset1.PageCount & """><img src=""imagens/Last.gif"" border=""0"" alt=""&Uacute;ltima""></a>&nbsp;" & Chr(13)
			Else
				Response.Write "<img src=""imagens/Next_off.gif"" border=""0"" alt=""Pr&oacute;xima"">&nbsp;" & Chr(13)
				Response.Write "<img src=""imagens/Last_off.gif"" border=""0"" alt=""&Uacute;ltima"">&nbsp;" & Chr(13)
			End If %></td>
                          </tr>
                        </table></td>
                      </tr>
                    </table></td>
                  </tr>
                </table>
            </div></td>
          </tr>
        </table>
      <table width="59%" border="0">
          <tr>
            <td></td>
          </tr>
      </table></td>
  </tr>
</table>

	<!-- FORMULARIOS ---------------------------------------------> 
  <!-- FIM FORMULARIO -->
</aol>	  
 
	</div>
				</div>
				<!-- END MainSection -->	
			</div>	</div>
	<!-- Scripts -->
	<script type="text/javascript" src="js/jquery.min.js"></script>
	<script type="text/javascript" src="js/jquery.fullPage.min.js"></script>
	<script type="text/javascript" src="js/jquery.dcjqaccordion.2.7.min.js"></script>
	<script type="text/javascript" src="js/main.js"></script>
	</body>
</html>


<!-- FECHA PROGRAMAÇÃO ------------------------------------>
  <%
cat.Close()
Set cat = Nothing
%>