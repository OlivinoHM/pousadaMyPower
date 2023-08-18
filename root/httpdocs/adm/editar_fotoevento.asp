
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
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_dlfelix_STRING
  MM_editTable = "eventos"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "listar_eventos.asp"
  MM_fieldsStr  = "img|value"
  MM_columnsStr = "img|',none,''"

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
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("id3") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("id3")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_dlfelix_STRING
Recordset1.Source = "SELECT *  FROM eventos  WHERE id = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>

<!--  FIM PROGRAMAÇÃO TOPO-->

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
	<div class="container">
							  <!-- page title --><!--/ page title -->
 <aol class="box">
 <div> 
<br>
<strong class="text-primary fa-2x azul">  &Xi; Editar Foto Evento</strong><br>
<br>
  

</div><br>
<br>
<form method="POST" action="<%=MM_editAction%>" name="form1">

	<table height="27%" cellSpacing="0" cellPadding="0" width="100%" border="0">
		<tr>
			<td vAlign="top" align="middle">
			<table border="0" style="border-collapse: collapse" width="741" cellpadding="0">
				<tr>
					<td width="739" align="right"></td>
				</tr>
			</table>
			<table cellSpacing="0" cellPadding="0" width="775" align="center" bgColor="" background="bg_pag.gif" border="0">
				<tr>
					<td width="775">&nbsp;</td>
				</tr>
				<tr>
					<td width="100%">
					<div align="center">
						<table cellSpacing="0" cellPadding="0" width="728" align="center" border="0">
							<tr>
								<td vAlign="top" bgColor="" width="579">
								<table border="0" style="border-collapse: collapse" width="100%" cellpadding="0">
									<tr>
										<td colspan="4">&nbsp;</td>
									</tr>
									<tr>
										<td width="209">
										<p align="right">Imagem:&nbsp;&nbsp; </td>
										<td width="39">
												<input type="hidden" name="img" size="4" value="<%=(Recordset1.Fields.Item("img").Value)%>"><input type="hidden" name="categoria" size="10" value="<%=(Recordset1.Fields.Item("nome").Value)%>">
                          </td>
										<td width="295">
												<img src="<%=(Recordset1.Fields.Item("img").Value)%>" name=img width="279" height="194"></td>
										<td width="175">
												<table border="0" width="86%" bgcolor="">
													<tr>
												<td width="170" bgColor="">
												<table border="0" style="border-collapse: collapse" width="100%" cellspacing="1" cellpadding="0" bgcolor="#000000">
													<tr>
														<td width="12">
														<a href="javascript:MM_openBrWindow('listar_eventos.asp?id=33&amp;id1=9','','width=815,height=400,scrollbars=yes')">
														<img alt="Inserir/Listar" src="listar1.gif" border="0" width="16" height="16"></a></td>
														<td bgcolor="#000000">
														<table border="0" style="border-collapse: collapse" width="100%">
															<tr>
																<td>
												<b>
												<font face="Verdana" color="#006699" size="1">
												<a onClick="window.open('imgeventos.asp?img=des01&id=1','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=190'); return false;" href="#">
												<font color="#FFFFFF">
												<span style="TEXT-DECORATION: none">
												Upload &gt;&gt;</span></font></a></font></b></td>
															</tr>
														</table>
														</td>
													</tr>
												</table>
												<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px">&nbsp;
												</p>
												<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px"></td>
													</tr>
												</table>
										</td>
									</tr>
									<tr>
										<td width="209" height="25">
										<p align="right">&nbsp;&nbsp; </td>
										<td height="25" colspan="3">&nbsp;</td>
									</tr>
									<tr>
										<td colspan="4">&nbsp;</td>
									</tr>
									<tr>
										<td colspan="4">&nbsp;</td>
									</tr>
									<tr>
										<td colspan="4">&nbsp;
										</td>
									</tr>
									<tr>
										<td colspan="4">&nbsp;</td>
									</tr>
									<tr>
										<td colspan="4">
										&nbsp;&nbsp;
										<center><input type="submit" value="Gravar e Visualizar" name="B1"></center></td>
									</tr>
									<tr>
										<td colspan="4">&nbsp;</td>
									</tr>
									</table>
							  <p></td>
							</tr>
						</table>
					</div>
					</td>
				</tr>
				<tr>
					<td width="775">&nbsp;</td>
				</tr>
			</table>
			<table width="59%" border="0">
				<tr>
					<td></td>
				</tr>
			</table>
			</td>
		</tr>
	</table>

        <input type="hidden" name="MM_update" value="form1">
<input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("id").Value %>">
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
Recordset1.Close()
Set Recordset1 = Nothing
%>