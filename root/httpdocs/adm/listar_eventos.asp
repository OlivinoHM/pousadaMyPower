<%@LANGUAGE="VBSCRIPT"%>
<% session.lcid=1046 %>

<!--#include file="Connections/dlfelix.asp" -->

		 <%	
	Set Recordset1 = Server.CreateObject("ADODB.Recordset")
	
    Recordset1.ActiveConnection = MM_dlfelix_STRING
Recordset1.Source = "SELECT * FROM eventos ORDER BY titulo asc"
	Recordset1.CursorLocation = 3
	Recordset1.CursorType = 1
	Recordset1.PageSize = 10
	Recordset1.CacheSize = 10
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

Repeat1__numRows = 10
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>

<!--  FIM PROGRAMAÇÃO TOPO-->

<!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml" class="">
		<head>
		<meta http-equiv="Content-Type" content="text/html; charset=latin1" />

		<title>MENU ADM</title>
 <!-- Mobile Specific Metas -->
		<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
 <!-- Basic Page Needs  -->
		<meta charset="utf-8">
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
				<i class="buton font-p"> &nbsp;&nbsp;&nbsp; Em caso de d&uacute;vidas, entre em contato com o Suporte clicando no bot&atilde;o SUPORTE SITE<span class="  font-m" ><strong></strong></span></i></div>
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
<strong class="text-primary fa-2x azul">  &Xi; Listar Eventos</strong>
<strong class="text-primary fa-2x right">  <a href="incluir_eventos.asp" class="verde">   &Xi; Incluir Eventos</a></strong>

<br>
 
  

</div><br>
<br>
	<form method="POST" action="<%=MM_editAction%>" name="form1">

	<input type="hidden" name="img5" size="19"><input type="hidden" name="img4" size="19"><input type="hidden" name="img3" size="19"><input type="hidden" name="img2" size="19"><input type="hidden" name="img1" size="19">
												<input type="hidden" name="data" size="19" value="<%=date%>">
												<table border="0" style="border-collapse: collapse" width="90%" cellpadding="0">
												      <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
												  <tr>
												        <td width="30" align="center">&nbsp;</td>
											          </tr>
                                                      <tr>
												        <td width="100%">
                                                        
                                                        
                          <div class="cantoredondo">
                            <table border="0" style="border-collapse: collapse" width="100%" cellpadding="6">
                              <tr>
                                <td width="30" align="center"><a href="editar_fotoevento.asp?id3=<%=(Recordset1.Fields.Item("id").Value)%>&id=<%=request("id")%>&id1=<%=request("id1")%>">Clique para atualizar ou incluir a foto da capa.<img src="<%=(Recordset1.Fields.Item("img").Value)%>" name=img width="184" height="120"></a></td>
                                <td width="611"><b><font color="#868686"></font></b>
                                  <table width="100%" height="130" border="0" cellpadding="0" cellspacing="2">
                                    <tr>
                                      <td><b>Titulo: <%=(Recordset1.Fields.Item("titulo").Value)%></b></td>
                                    </tr>
                                  </table></td>
                                <td width="49" align="center">&nbsp;</td>
                                <td width="49" align="center"><a href="editar_eventos.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>"> <div class=" icon-tag font-m right pading-botao" style="display:block; float:right"></div></a></td>
                                <td width="30" align="center"><a onClick="javascript:if (!confirm('Voc&ecirc; tem certeza que deseja excluir?')){return false;}" href="del_eventos.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>"> <div class=" icon-fechar font-m right pading-botao" style="display:block; float:right"></div></a></td>
                                
                              </tr>
                            </table>
                          </div>                              
                                                        
                                                        
                                                        
                                                        </td>
											          </tr>
												      
												      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
	        </table>
												<input type="hidden" name="MM_insert" value="form1">
	</form>

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
 