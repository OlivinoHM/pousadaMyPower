<%@LANGUAGE="VBSCRIPT"%>
<% session.lcid=1046 %>

<!--#include file="Connections/dlfelix.asp" -->

<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="index.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>

<%
Dim cat1
Dim cat1_numRows

Set cat1 = Server.CreateObject("ADODB.Recordset")
cat1.ActiveConnection = MM_dlfelix_STRING
cat1.Source = "SELECT *  FROM subcategorias ORDER BY nome ASC"
cat1.CursorType = 0
cat1.CursorLocation = 2
cat1.LockType = 1
cat1.Open()

cat1_numRows = 0
%>
 
<%
Dim cat
Dim cat_numRows

Set cat = Server.CreateObject("ADODB.Recordset")
cat.ActiveConnection = MM_dlfelix_STRING
cat.Source = "SELECT *  FROM subcategorias ORDER BY nome ASC"
cat.CursorType = 0
cat.CursorLocation = 2
cat.LockType = 1
cat.Open()

cat_numRows = 0
%>

<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_dlfelix_STRING
  MM_editTable = "home"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "edit_home1.asp"
  MM_fieldsStr  = "img1|value|titulo1|value|resumo1|value|link1|value|img2|value|titulo2|value|resumo2|value|link2|value|img3|value|titulo3|value|resumo3|value|link3|value|img4|value|titulo4|value|resumo4|value|link4|value|img5|value|titulo5|value|resumo5|value|link5|value|img6|value|titulo6|value|resumo6|value|link6|value|img7|value|titulo7|value|resumo7|value|link7|value"
  MM_columnsStr = "img1|',none,''|titulo1|',none,''|resumo1|',none,''|link1|',none,''|img2|',none,''|titulo2|',none,''|resumo2|',none,''|link2|',none,''|img3|',none,''|titulo3|',none,''|resumo3|',none,''|link3|',none,''|img4|',none,''|titulo4|',none,''|resumo4|',none,''|link4|',none,''|img5|',none,''|titulo5|',none,''|resumo5|',none,''|link5|',none,''|img6|',none,''|titulo6|',none,''|resumo6|',none,''|link6|',none,''|img7|',none,''|titulo7|',none,''|resumo7|',none,''|link7|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>

<%
Dim atualiza__MMColParam
atualiza__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  atualiza__MMColParam = Request.QueryString("id")
End If
%>
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



<!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml" class="">
		<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
				<i class="buton font-p"> &nbsp;&nbsp;&nbsp; Em caso de dúvidas, entre em contato com o Suporte clicando no botão SUPORTE SITE<span class="  font-m" ><strong></strong></span></i></div>
				<!-- /Navigation Menu -->
				<div class="site-overlay"></div>
				<!-- Main SECTION -->	
				<div class="section">
	<div class="page-wrapper">
	<table width="1000" border="0" cellspacing="10">
  <tr>
    <td colspan="4" class="azul">&larr; PARA INICIAR O GETENCIAMENTO DO SITE CLIQUE NO MENU AO LADO</td>
    </tr>
  <tr>
    <td colspan="4">&nbsp;</td>
    </tr>
  <tr>
    <td width="11%"><span class="icon-home icone font-m" ><span class="font-p "><br>
                <br>
                Editar 
      </span><span class="font-p "></span><span class="font-p ">inicial</span></span></td>
    <td width="35%" align="left" valign="top"> Editar funções e conteúdo da página inicial</td>
    <td width="11%"><span class="icon-sub-titulo icone font-m" ><span class="font-p"><br>
            <br>
Listar<br>
Pacote</span></span></td>
    <td width="43%" align="left" valign="top">Listar, cadastrar e alterar pacotes</td>
  </tr>
  <tr>
    <td><span class=" icon-briefcase icone font-m" ><span class="font-p "><br>
        <br>
        Editar 
        Dados</span></span></td>
    <td align="left" valign="top">Editar informações da empresa</td>
    <td><span class=" icon-sair icone font-m" ><span class="font-p"><br>
            <br>
Sair<br>
Painel</span></span></td>
    <td align="left" valign="top">Sair do Painel administrativo</td>
  </tr>
  <tr>
    <td><span class="icon-bubble icone font-m" ><span class="font-p"><br>
                <br>
              Listar<br>
                Evento</span></span></td>
    <td align="left" valign="top">Listar, cadastrar e alterar eventos</td>
    <td><span class=" icon-html5 icone font-m" ><span class="font-p"><br>
           <br>
              Editar<br>
          Banner</span></span></td>
    <td align="left" valign="top">Alterar foto e informações do banner principal do site</td>
    
  </tr>
  <tr>
    <td><span class="icon-tag icone font-m" ><span class="font-p"><br>
            <br>
          Listar 
          Suite</span></span></td>
    <td align="left" valign="top">Listar, cadastrar e alterar suítes</td>
    <td><span class=" icon-estatistica2 icone font-m" ><a href="http://www.printsystem.com.br/"><span class="font-p"><br>
                      <br>
                Suporte<br>
                      Site</span></a></span>
    <td align="left" valign="top">Para obter contato com o programador</td>
  </tr>
  <tr>
    <td><span class=" icon-newspaper icone font-m" ><span class="font-p"><br>
            <br>
Listar<br>
Lazer</span></span></td>
    <td align="left" valign="top">Listar, cadastrar e alterar informações de lazer</td>
    <td>&nbsp;</td>
    <td align="left" valign="top">&nbsp;</td>
  </tr>
</table>

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
<%
atualiza.Close()
Set atualiza = Nothing
%>
<%
cat.Close()
Set cat = Nothing
%>
<%
cat1.Close()
Set cat1 = Nothing
%>