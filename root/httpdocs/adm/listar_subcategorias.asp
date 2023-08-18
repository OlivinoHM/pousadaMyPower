
<%@LANGUAGE="VBSCRIPT"%>
<% session.lcid=1046 %>

<!--#include file="Connections/dlfelix.asp" -->

		 <%	
	Set Recordset1 = Server.CreateObject("ADODB.Recordset")
	
    Recordset1.ActiveConnection = MM_dlfelix_STRING
Recordset1.Source = "SELECT categorias.nome as ok, categorias.id as idcat, subcategorias.categoria, subcategorias.subcategoria,subcategorias.id FROM subcategorias INNER JOIN categorias ON (subcategorias.categoria = categorias.id)  ORDER BY categorias.nome asc"
	Recordset1.CursorLocation = 3
	Recordset1.CursorType = 1
	Recordset1.PageSize = 1000
	Recordset1.CacheSize = 1000
	Recordset1.Open()

	Dim contadorvideos

	If request("pagina") <> "" Then
		Recordset1.AbsolutePage = request("pagina")
		pagina = request("pagina")
	Else
		If NOT Recordset1.EOF Then Recordset1.AbsolutePage = 1
		pagina = 1
	End If %>
	
	

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 1000
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>

<!--  FIM PROGRAMAÃ‡ÃƒO TOPO-->

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
				<i class="buton font-p"> &nbsp;&nbsp;&nbsp; Em caso de dÃºvidas, entre em contato com o Suporte clicando no botÃ£o SUPORTE SITE<span class="  font-m" ><strong></strong></span></i></div>
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
<strong class="text-primary fa-2x azul">  &Xi; Listar Pacotes</strong>
<strong class="text-primary fa-2x right">  <a href="incluir_categoria.asp" class="verde">   &Xi; Incluir Pacotes&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; </a></strong><br>
<br>
  

</div><br>
<br>
 <aol class="box">


<!--- conteudo programado ------->


<input type="hidden" name="img5" size="19"><input type="hidden" name="img4" size="19"><input type="hidden" name="img3" size="19"><input type="hidden" name="img2" size="19"><input type="hidden" name="img1" size="19">
												<input type="hidden" name="data" size="19" value="<%=date%>">
												<table border="0" style="border-collapse: collapse" width="100%" cellpadding="0">
												  <tr>
												    <td width="579"><table border="0" style="border-collapse: collapse" width="100%">
												      <tr>
												        <td width="244" align="left">Categoria</td>
												        <td width="281" align="left"><p>Sub-Categoria</td>
												        <td width="45" align="center"><font color="#666666">Fotos</font></td>
												        <td width="36" align="center"><font color="#666666">Editar</font></td>
												        <td width="41" align="center"><font color="#666666">Excluir</font></td>
											          </tr>
												      </table></td>
											      </tr>
												  <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
												  <tr>
												    <td width="579"><div class=" cantoredondo"><table border="0" style="border-collapse: collapse" width="100%" cellpadding="6">
												      <tr>
												        <td width="146" height="44" align="left"><b><font color="#868686"><%=(Recordset1.Fields.Item("ok").Value)%></font></b></td>
												        <td width="222" align="left"><b><%=(Recordset1.Fields.Item("subcategoria").Value)%></b></td>
												        <td align="center"><a onClick="javascript:if (!confirm('Voc&ecirc; tem certeza que deseja excluir?')){return false;}" href="del_subcategoria.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>"> <div class=" icon-fechar font-m right pading-botao" style="display:block; float:right"></div></a><a href="editar_subcategoria.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>&id1=<%=(Recordset1.Fields.Item("idcat").Value)%>"><div class=" icon-tag font-m right pading-botao" style="display:block; float:right"></div></a><a href="fotos.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>&id1=<%=(Recordset1.Fields.Item("idcat").Value)%>"> <div class=" icon-foto font-m right pading-botao" style="display:block; float:right"></div></a></td>
											          </tr>
												      </table></div></td>
											      </tr>
												  <tr>
												    <td width="579" ><img src=" " width="1" height="1"></td>
											      </tr>
												  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>


 </aol>

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


<!-- FECHA PROGRAMAÃ‡ÃƒO ------------------------------------>
  <%
cat.Close()
Set cat = Nothing
%>