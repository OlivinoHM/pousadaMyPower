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
  MM_editTable = "home"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "edit_banner1.asp"
  MM_fieldsStr  = "img1|value|titulo1|value|resumo1|value|link1|value|img2|value|titulo2|value|resumo2|value|link2|value|img3|value|titulo3|value|resumo3|value|link3|value"
  MM_columnsStr = "img1|',none,''|titulo1|',none,''|resumo1|',none,''|link1|',none,''|img2|',none,''|titulo2|',none,''|resumo2|',none,''|link2|',none,''|img3|',none,''|titulo3|',none,''|resumo3|',none,''|link3|',none,''"

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
<%
Dim cat
Dim cat_numRows

Set cat = Server.CreateObject("ADODB.Recordset")
cat.ActiveConnection = MM_dlfelix_STRING
cat.Source = "SELECT *  FROM pgpacotes  ORDER BY pactitulo ASC"
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
cat2.Source = "SELECT *  FROM eventos ORDER BY titulo ASC"
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
cat3.Source = "SELECT *  FROM lazer ORDER BY titulo ASC"
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
<strong class="text-primary fa-2x azul">  &Xi; Editar Banner Rotativo</strong><br>
<br>
  

</div><br>
<br><form method="POST" action="<%=MM_editAction%>" name="form1"><table height="27%" cellSpacing="0" cellPadding="0" width="100%" border="0">
		<tr>
			<td vAlign="top" align="middle"><div align="center">
              <table width="800" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td><div align="left">
                      <table width="99%" border="0" cellspacing="0" cellpadding="10">
                        <tr>
                          <td><div align="center">
                            <table border="0" style="border-collapse: collapse" width="728" cellpadding="0">
                              
                              <tr>
                                <td width="100%"><table border="0" width="100%" cellpadding="6" bgcolor="">
                                    <tr>
                                      <% if request("edit") = "ok" then %>
                                      <td bgcolor=""><table border="0" width="100%">
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
                                            <td><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                                <tr>
                                                  <td width="11"><img border="0" src="edit_picture.png" width="22" height="22"></td>
                                                  <td class="style14"><b> <font face="Verdana" size="1">&nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=1','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=200'); return false;" href="#" style="text-decoration: none"><font color="#000000" class="b-alterar">Cadastrar 
                                                    Imagem</font></a> </font></b> <font face="Arial" size="1"> (Tamanho em 
                                                      pixels: 1280x400px 
                                                      )
                                                      <input name="img1" type="hidden" value="<%=(atualiza.Fields.Item("img1").Value)%>" size="19">
                                                                                      </font></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><img height="187" src="<%=(atualiza.Fields.Item("img1").Value)%>" align="center" width="592" border="0" name=img1></td>
                                          </tr>
                                          <tr>
                                            <td class="style10">&nbsp;</td>
                                          </tr>
                                          <tr>
                                            <td><table border="0" width="111%">
                                              <tr>
                                                <td width="52" class="style22"> Titulo:</td>
                                                <td><input type="text" name="titulo1" size="20" value="<%=(atualiza.Fields.Item("titulo1").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td width="52" class="style10"> Resumo:</td>
                                                <td><textarea rows="3" name="resumo1" cols="35"><%=(atualiza.Fields.Item("resumo1").Value)%></textarea></td>
                                              </tr>
                                              <tr>
                                                <td width="52" class="style11"> Link:</td>
                                                <td><select size="1" name="link1">
          <%
While (NOT cat3.EOF)
%>
          <option value="<%=(cat3.Fields.Item("id").Value)%>" <%If (Not isNull((cat3.Fields.Item("titulo").Value))) Then If (CStr(cat3.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("link1").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Lazer)&nbsp;-&nbsp;<%=(cat3.Fields.Item("titulo").Value)%></option>
          <%
  cat3.MoveNext()
Wend
If (cat3.CursorType > 0) Then
  cat3.MoveFirst
Else
  cat3.Requery
End If
%>

        </select></td>
                                              </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><img height="10" width="1"></td>
                                          </tr>
                                          <tr>
                                            <td><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                                <tr>
                                                  <td width="11"><img border="0" src="edit_picture.png" width="22" height="22"></td>
                                                  <td class="style14"><b> <font face="Verdana" size="1">&nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=2','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=200'); return false;" href="#" style="text-decoration: none"><font color="#000000" class="b-alterar">CadastrarImagem</font></a> </font></b> <font face="Arial" size="1"> (Tamanho em 
                                                    pixels: 1280x400px )
                                                    <input name="img2" type="hidden" id="img2" value="<%=(atualiza.Fields.Item("img2").Value)%>" size="19">
                                                  </font></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><img height="187" src="<%=(atualiza.Fields.Item("img2").Value)%>" width="592" border="0" name=img2></td>
                                          </tr>
                                          <tr>
                                            <td class="style10">&nbsp;</td>
                                          </tr>
                                          <tr>
                                            <td><table border="0" width="111%">
                                              <tr>
                                                <td width="52" class="style22"> Titulo:</td>
                                                <td><input type="text" name="titulo2" size="20" value="<%=(atualiza.Fields.Item("titulo2").Value)%>"></td>
                                              </tr>
                                              <tr>
                                                <td width="52" class="style10"> Resumo:</td>
                                                <td><textarea rows="3" name="resumo2" cols="35"><%=(atualiza.Fields.Item("resumo2").Value)%></textarea></td>
                                              </tr>
                                              <tr>
                                                <td width="52" class="style11"> Link:</td>
                                                <td><select size="1" name="link2">
                                                  <%
While (NOT cat3.EOF)
%>
                                                  <option value="<%=(cat3.Fields.Item("id").Value)%>" <%If (Not isNull((cat3.Fields.Item("titulo").Value))) Then If (CStr(cat3.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("link2").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Lazer)&nbsp;-&nbsp;<%=(cat3.Fields.Item("titulo").Value)%></option>
                                                  <%
  cat3.MoveNext()
Wend
If (cat3.CursorType > 0) Then
  cat3.MoveFirst
Else
  cat3.Requery
End If
%>
                                                </select></td>
                                              </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><img height="10" width="1"></td>
                                          </tr>
                                          <tr>
                                            <td height="10"><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                                <tr>
                                                  <td width="11"><img border="0" src="edit_picture.png" width="22" height="22"></td>
                                                  <td><b> <font face="Verdana" size="1">&nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=3','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=200'); return false;" href="#" style="text-decoration: none"><font color="#000000" class="b-alterar">Cadastrar 
                                                    Imagem</font></a> </font></b> <font face="Arial" size="1"> (Tamanho em 
                                                      pixels: 1280x400px )
                                                      <input name="img3" type="hidden" id="img3" value="<%=(atualiza.Fields.Item("img3").Value)%>" size="19">
                                                                                      </font></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          
                                          <tr>
                                            <td><table cellSpacing="0" cellPadding="0" width="321" border="0">
                                                <tr>
                                                  <td align="left" vAlign="top"><img height="187" src="<%=(atualiza.Fields.Item("img3").Value)%>" width="592" border="0" name=img3></td>
                                                </tr>
                                                <tr>
                                                  <td align="left" vAlign="top"><div align="justify">
                                                      <table border="0" width="111%">
                                                        <tr>
                                                          <td width="52" class="style22"> Titulo:</td>
                                                          <td><input type="text" name="titulo3" size="20" value="<%=(atualiza.Fields.Item("titulo3").Value)%>"></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="52" class="style10"> Resumo:</td>
                                                          <td><textarea rows="3" name="resumo3" cols="35"><%=(atualiza.Fields.Item("resumo3").Value)%></textarea></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="52" class="style11"> Link:</td>
                                                          <td><select size="1" name="link3">
                                                            <%
While (NOT cat3.EOF)
%>
                                                            <option value="<%=(cat3.Fields.Item("id").Value)%>" <%If (Not isNull((cat3.Fields.Item("titulo").Value))) Then If (CStr(cat3.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("link3").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Lazer)&nbsp;-&nbsp;<%=(cat3.Fields.Item("titulo").Value)%></option>
                                                            <%
  cat3.MoveNext()
Wend
If (cat3.CursorType > 0) Then
  cat3.MoveFirst
Else
  cat3.Requery
End If
%>
                                                          </select></td>
                                                        </tr>
                                                      </table>
                                                  </div></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                      </table></td>
                                      <td vAlign="top" width="47" rowSpan="2"><img height="1" width="20"></td>
                                      </tr>
                                  </table>                                  �</td>
                              </tr>
                              <tr>
                                <td width="728" bgcolor="#FFFFFF"><p align="center">
                                    <input type="submit" value="Atualizar Banner Rotativo" name="B1" style="color: #FFFFFF; border: 1px solid #CC3300; background-color: #CC3300">
                                </td>
                              </tr>
                            </table>
                          </div></td>
                        </tr>
                      </table>
                  </div></td>
                </tr>
               
               
                <tr>
                  <td>&nbsp;</td>
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


<!-- FECHA PROGRAMAÇÃO ------------------------------------><%
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