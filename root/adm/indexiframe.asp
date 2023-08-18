<!DOCTYPE html>
<html> <head>
<!-- Inicio Programação topo ------------------------------------>
<!--#include file="../adm/Connections/dlfelix.asp" -->

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




<iframe src="http://sistemhost.com.br/mypower/adm/index.asp" width="100%" height="1000" marginwidth="0" marginheight="0" align="center" border:0; " hspace="0" vspace="0" frameborder="0" scrolling="no"></iframe>


</body>
</html>
