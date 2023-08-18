<%@LANGUAGE="VBSCRIPT"%>
<% session.lcid=1046 %>

<!--#include file="Connections/dlfelix.asp" -->

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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_dlfelix_STRING
  MM_editTable = "eventos"
  MM_editRedirectUrl = "listar_eventos.asp"
  MM_fieldsStr  = "titulo|value|data1|value|resumo|value|texto|value|link|value"
  MM_columnsStr = "titulo|',none,''|data1|',none,''|resumo|',none,''|texto|',none,''|link|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
Dim cat
Dim cat_numRows

Set cat = Server.CreateObject("ADODB.Recordset")
cat.ActiveConnection = MM_dlfelix_STRING
cat.Source = "SELECT *  FROM categorias  ORDER BY nome ASC"
cat.CursorType = 0
cat.CursorLocation = 2
cat.LockType = 1
cat.Open()

cat_numRows = 0
%>
<%
Dim cat1
Dim cat1_numRows

Set cat1 = Server.CreateObject("ADODB.Recordset")
cat1.ActiveConnection = MM_dlfelix_STRING
cat1.Source = "SELECT *  FROM categoria ORDER BY nome ASC"
cat1.CursorType = 0
cat1.CursorLocation = 2
cat1.LockType = 1
cat1.Open()

cat1_numRows = 0
%>
<%
Dim cat2
Dim cat2_numRows

Set cat2 = Server.CreateObject("ADODB.Recordset")
cat2.ActiveConnection = MM_dlfelix_STRING
cat2.Source = "SELECT *  FROM pgpacotes ORDER BY pactitulo ASC"
cat2.CursorType = 0
cat2.CursorLocation = 2
cat2.LockType = 1
cat2.Open()

cat2_numRows = 0
%>
<%
Dim cat3
Dim cat3_numRows

Set cat3 = Server.CreateObject("ADODB.Recordset")
cat3.ActiveConnection = MM_dlfelix_STRING
cat3.Source = "SELECT *  FROM eventos ORDER BY titulo ASC"
cat3.CursorType = 0
cat3.CursorLocation = 2
cat3.LockType = 1
cat3.Open()

cat3_numRows = 0
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
				<i class="buton font-p"> &nbsp;&nbsp;&nbsp; Em caso de dúvidas, entre em contato com o Suporte clicando no botão SUPORTE SITE<span class="  font-m" ><strong></strong></span></i></div>
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
<strong class="text-primary fa-2x azul">  &Xi; Incluir Eventos</strong>
<strong class="text-primary fa-2x right">  <a href="incluir_eventos.asp" class="verde">&nbsp; &nbsp; &nbsp;&nbsp; &nbsp; </a></strong><br>
<br>
  

</div><br>
<br>
 <form method="POST" action="<%=MM_editAction%>" name="form1">

	<input type="hidden" name="img5" size="19"><input type="hidden" name="img4" size="19"><input type="hidden" name="img3" size="19"><input type="hidden" name="img2" size="19"><input type="hidden" name="img1" size="19">
												<input type="hidden" name="data" size="19" value="<%=date%>">
												<table border="0" style="border-collapse: collapse" width="90%" cellpadding="0">
												  <tr>
												    <td width="579" height="22"><p align="right">Nome Eventos:&nbsp;&nbsp; </td>
												    <td width="455" height="22"><input type="text" name="titulo" size="38"></td>
											      </tr>
												  <tr>
												    <td width="579"><p align="right">Data do evento:&nbsp;&nbsp; </td>
												    <td width="455"><input type="text" name="data1" size="38"></td>
											      </tr>
                                                  <tr>
												    <td width="579"><p align="right">Resumo do evento:&nbsp;&nbsp; </td>
												    <td width="455"><textarea name="resumo" cols="38" rows="2"></textarea></td>
											      </tr>
                                                  <tr>
												    <td width="579"><p align="right">Texto do evento:&nbsp;&nbsp; </td>
												    <td width="455"><textarea name="texto" cols="38" rows="5"></textarea></td>
											      </tr>
                                                  <tr>
												    <td width="579"><p align="right">Link:&nbsp;&nbsp; </td>
												    <td width="455"><select size="1" name="link">
												      <option value="<%=(cat1.Fields.Item("id").Value)%>">&lt;%=(cat1.Fields.Item(&quot;titulo&quot;).Value)%&gt;</option>
												      <option>------------------------------</option>
												      <option value="<%=(cat2.Fields.Item("id").Value)%>">&lt;%=(cat2.Fields.Item(&quot;pactitulo&quot;).Value)%&gt;</option>
												      <option>------------------------------</option>
												      <option value="<%=(cat3.Fields.Item("id").Value)%>">&lt;%=(cat3.Fields.Item(&quot;titulo&quot;).Value)%&gt;</option>
												      <%
While (NOT cat.EOF)
%>
												      <%
  cat.MoveNext()
Wend
If (cat.CursorType > 0) Then
  cat.MoveFirst
Else
  cat.Requery
End If
%>
											        </select></td>
											      </tr>
												  <tr>
												    <td width="579" colspan="2">&nbsp;</td>
											      </tr>
												  <tr>
												    <td width="579" colspan="2">&nbsp;&nbsp;
												      <input name="B1" type="submit" class="botaoverde" value="Gravar e Visualizar"></td>
											      </tr>
												  <tr>
												    <td width="579" colspan="2">&nbsp;</td>
											      </tr>
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
  <%
cat.Close()
Set cat = Nothing
%>