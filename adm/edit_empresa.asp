
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
  MM_editTable = "empresa"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "edit_empresa1.asp"
  MM_fieldsStr  = "img1|value|nomeempresa|value|sobreempresa|value|endereco|value|telefone1|value|telefone2|value|telefone3|value|email|value|site|value|face|value|maps|value|titulodicas|value|textodicas|value|img2|value|linkimg2|value"
  MM_columnsStr = "img1|',none,''|nomeempresa|',none,''|sobreempresa|',none,''|endereco|',none,''|telefone1|',none,''|telefone2|',none,''|telefone3|',none,''|email|',none,''|site|',none,''|face|',none,''|maps|',none,''|titulodicas|',none,''|textodicas|',none,''|img2|',none,''|linkimg2|',none,''"

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
atualiza.Source = "SELECT *  FROM empresa"
atualiza.CursorType = 0
atualiza.CursorLocation = 2
atualiza.LockType = 1
atualiza.Open()

atualiza_numRows = 0
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
<strong class="text-primary fa-2x azul">  &Xi; Página Dados da Empresa</strong><br>
<br>
  

</div><br>
<br>
<form method="POST" action="<%=MM_editAction%>" name="form1"><table height="27%" cellSpacing="0" cellPadding="0" width="100%" border="0">
		<tr>
			<td vAlign="top" align="middle"><div align="center">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td background=""><div align="left">
                      <table width="99%" border="0" cellspacing="0" cellpadding="10">
                        <tr>
                          <td><div align="center">
                            <table border="0" style="border-collapse: collapse" width="728" cellpadding="0">
                              <tr>
                                
                              </tr>
                              <tr>
                                <td width="728"><table border="0" width="100%" cellpadding="6">
                                    <tr>
                                      <% if request("edit") = "ok" then %>
                                      <td><table border="0" width="100%">
                                          <tr>
                                            <td width="54"><img border="0" src="ok.png" width="50" height="50"></td>
                                            <td> <span class="style27"> <font face="Arial">Atualiza&ccedil;&atilde;o efetuada com 
                                              sucesso!</font></span></td>
                                          </tr>
                                        </table>
                                          <%end if%>                                      </td>
                                    </tr>
                                </table></td>
                              </tr>

                              <tr>
                                <td width="728"><table cellSpacing="0" cellPadding="0" width="98%" border="0">
                                    <tr>
                                      <td vAlign="top" width="275"><table cellSpacing="0" cellPadding="0" width="90%" border="0">
                                        <tr>
                                            <td>&nbsp;</td>
                                          </tr>
                                          <tr>
                                            <td><table border="0" width="111%">
                                              <tr>
                                                <td width="110" class="style22">&nbsp;</td>
                                                <td width="560" align="right"><span class="style14"><b><font face="Verdana" size="1">&nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=1','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=200'); return false;" href="#" style="text-decoration: none;background: #F60; padding:5px; border-radius:5px;"><font color="#000000">Cadastrar 
                                                    Imagem</font></a></font></b> <font face="Arial" size="1">
                                                    <input name="img1" type="hidden" value="<%=(atualiza.Fields.Item("img1").Value)%>" size="19">
                                                </font></span></td>
                                              </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td class="style10">&nbsp;</td>
                                          </tr>
                                          <tr>
                                            <td><table border="0" width="111%">
                                              <tr>
                                                <td width="52" class="style22">&nbsp;</td>
                                                <td><span class="style10"><center><img src="<%=(atualiza.Fields.Item("img1").Value)%>" alt="" name=img1 height="221" border="0"></center></span></td>
                                              </tr>
                                              <tr>
                                                <td colspan="2" align="center" class="titulo"> <hr></td>
                                                </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Nome da Empresa:</div></td>
                                                <td><input type="text" name="nomeempresa" size="70" value="<%=(atualiza.Fields.Item("nomeempresa").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td class="style10"><div align="right">Sobre a Empresa:</div></td>
                                                <td><textarea name="sobreempresa" cols="70" rows="10"><%=(atualiza.Fields.Item("sobreempresa").Value)%></textarea></td>
                                              </tr>
                                              <tr>
                                                <td class="style11"><div align="right"></div></td>
                                                <td>&nbsp;</td>
                                              </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Endereço:</div></td>
                                                <td><textarea rows="4" name="endereco" cols="70"><%=(atualiza.Fields.Item("endereco").Value)%></textarea></td>
                                              </tr>
                                              <tr>
                                                <td class="style11"><div align="right"></div></td>
                                                <td>&nbsp;</td>
                                              </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Telefone1:</div></td>
                                                <td><input type="text" name="telefone1" size="70" value="<%=(atualiza.Fields.Item("telefone1").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Telefone2:</div></td>
                                                <td><input type="text" name="telefone2" size="70" value="<%=(atualiza.Fields.Item("telefone2").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Telefone3:</div></td>
                                                <td><input type="text" name="telefone3" size="70" value="<%=(atualiza.Fields.Item("telefone3").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td class="style11"><div align="right"></div></td>
                                                <td>&nbsp;</td>
                                              </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Email:</div></td>
                                                <td><input type="text" name="email" size="70" value="<%=(atualiza.Fields.Item("email").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td colspan="2" align="center" class="style11"><span class="titulo"> <hr> </span></td>
                                                </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Site: <font color="#000000">www.</font></div></td>
                                                <td><input type="text" name="site" size="70" value="<%=(atualiza.Fields.Item("site").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td class="style11"><div align="right"></div></td>
                                                <td>&nbsp;</td>
                                              </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Pagina Facebook:</div></td>
                                                <td><input type="text" name="face" size="70" value="<%=(atualiza.Fields.Item("face").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td class="style22">&nbsp;</td>
                                                <td>&nbsp;</td>
                                              </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Endereço Google Maps (HTTP):</div></td>
                                                <td><textarea rows="6" name="maps" cols="70"><%=(atualiza.Fields.Item("maps").Value)%></textarea></td>
                                              </tr>
                                              
                                                            <!-----dicas 
                                              
                                              <tr>
                                                <td colspan="2" class="style11"><span class="titulo"> <hr></span></td>
                                              </tr>
                                              <tr>
                                                <td colspan="2" class="style11"><span style="text-transform: uppercase"><b><font face="Verdana">dicas</font></b></span></td>
                                                </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Titulo Dicas:</div></td>
                                                <td><input type="text" name="titulodicas" size="70" value="<%=(atualiza.Fields.Item("titulodicas").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Texto Dicas:</div></td>
                                                <td><textarea rows="6" name="textodicas" cols="70"><%=(atualiza.Fields.Item("textodicas").Value)%></textarea></td>
                                              </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Imagem Dicas:</div></td>
                                                <td><span class="style10"><br>
                                                  <img src="<%=(atualiza.Fields.Item("img2").Value)%>" alt="" name=img2 height="176" border="0"></span></td>
                                              </tr>
                                              <tr bgcolor="#FFFFFF">
                                                <td bgcolor="#FFFFFF">&nbsp;</td>
                                                <td align="right" class="style14"><b> <font face="Verdana" size="1"> &nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=2','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=200'); return false;" href="#" style="text-decoration: none;background: #F60; padding:5px; border-radius:5px;"><font color="#000000">Cadastrar 
                                                  Imagem</font></a></font></b> <font face="Arial" size="1"> (Tamanho em 
                                                    pixels: 358x260 
                                                    )
                                                    <input name="img2" type="hidden" value="<%=(atualiza.Fields.Item("img2").Value)%>" size="19">
                                                  </font></td>
                                              </tr>
                                              <tr>
                                                <td class="style22"><div align="right">Link da Foto:</div></td>
                                                <td><input type="text" name="linkimg2" size="70" value="<%=(atualiza.Fields.Item("linkimg2").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td class="style22">&nbsp;</td>
                                                <td>&nbsp;</td>
                                              </tr> ------>
                                          </table></td>
                                          </tr>
                                      </table></td>
                                    </tr>
                                  </table></td>
                              </tr>
                              <tr>
                                <td width="728"><p align="center">
                                    <input type="submit" value="Atualizar dados da Empresa" name="B1" style="color: #FFFFFF; border: 1px solid #CC3300; background-color: #CC3300">
                                </td>
                              </tr>
                            </table>
                          </div></td>
                        </tr>
                      </table>
                  </div></td>
                </tr>
              </table>
			  </div>
			  <table width="59%" border="0">
				<tr>
			
				</tr>
			</table>
			</td>
		</tr>
	</table>

<input type="hidden" name="MM_update" value="form1">
<input type="hidden" name="MM_recordId" value="<%= atualiza.Fields.Item("id").Value %>">
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
atualiza.Close()
Set atualiza = Nothing
%>