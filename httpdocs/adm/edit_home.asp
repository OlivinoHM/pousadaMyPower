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
cat1.Source = "SELECT *  FROM categoria ORDER BY nome ASC"
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
cat.Source = "SELECT *  FROM lazer ORDER BY nome ASC"
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
  MM_fieldsStr  = "img4|value|titulo4|value|resumo4|value|link4|value|img5|value|titulo5|value|resumo5|value|link5|value|img6|value|titulo6|value|resumo6|value|link6|value|img7|value|titulo7|value|resumo7|value|link7|value|img8|value|titulo8|value|resumo8|value|link8|value|img9|value|titulo9|value|resumo9|value|link9|value|img10|value|titulo10|value|resumo10|value|link10|value"
  MM_columnsStr = "img4|',none,''|titulo4|',none,''|resumo4|',none,''|link4|',none,''|img5|',none,''|titulo5|',none,''|resumo5|',none,''|link5|',none,''|img6|',none,''|titulo6|',none,''|resumo6|',none,''|link6|',none,''|img7|',none,''|titulo7|',none,''|resumo7|',none,''|link7|',none,''|img8|',none,''|titulo8|',none,''|resumo8|',none,''|link8|',none,''|img9|',none,''|titulo9|',none,''|resumo9|',none,''|link9|',none,''|img10|',none,''|titulo10|',none,''|resumo10|',none,''|link10|',none,''"

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
<strong class="text-primary fa-2x azul">  &Xi; Editar inicial</strong><br>
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
                                     <td><center><strong class="text-primary fa-2x verde">------</strong> </center></td></tr>
                                           <tr>
                                             <td><center>
                                             <strong class="text-primary fa-2x verde"><a href="edit_pacotes.asp">Clique aqui para editar pacotes em destaque</a></strong> 
                                             </center></td></tr>
                                             <tr>
                                             <td><center><strong class="text-primary fa-2x verde">------</strong> </center></td></tr>
                                         
                                          <tr>
                                             <td><center><strong class="text-primary fa-2x verde">------</strong> </center></td></tr>
                                           <tr>
                                             <td><center><strong class="text-primary fa-2x verde">Lazer em destaque</strong> </center></td></tr>
                                             <tr>
                                             <td><center><strong class="text-primary fa-2x verde">------</strong> </center></td></tr>
                                             
                                           <tr> <td height="10"><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                               
                                                <tr>
                                                  <td width="11"><img border="0" src="edit_picture.png" width="22" height="22"></td>
                                                  <td class="style14"><font face="Verdana" size="1"><b> &nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=4','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=900,height=350'); return false;" href="#" style="text-decoration: none"><font color="#000000" class="botoes">Cadastrar 
                                                    Imagem</font></a> </b></font> <font face="Arial" size="1"> (Tamanho em 
                                                      pixels: 96x96 )
                                                      <input name="img4" type="hidden" value="<%=(atualiza.Fields.Item("img4").Value)%>" size="19">
                                                  </font></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><table cellSpacing="0" cellPadding="0" width="100%" border="0">
                                                <tr>
                                                  <td vAlign="top" align="left" width="25%"><img height="176" src="<%=(atualiza.Fields.Item("img4").Value)%>" width="213" border="0" name=img4></td>
                                                  <td vAlign="top" align="left" width="68%"><div align="justify">
                                                      <table border="0" width="100%">
                                                      <tr>
                                                          <td width="44" class="style22"> Titulo:</td>
                                                          <td width="199"><input name="titulo4" type="text" value="<%=(atualiza.Fields.Item("titulo4").Value)%>" size="100"></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style10"> Resumo:</td>
                                                          <td><textarea rows="3" name="resumo4" cols="28"><%=(atualiza.Fields.Item("resumo4").Value)%></textarea></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style11"> Link:</td>
                                                          <td><select size="1" name="link4">
                                                            <%
While (NOT cat.EOF)
%>
                                                            <option value="<%=(cat.Fields.Item("id").Value)%>" <%If (Not isNull((cat.Fields.Item("titulo").Value))) Then If (CStr(cat.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("link4").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Lazer)&nbsp;-&nbsp;<%=(cat.Fields.Item("titulo").Value)%></option>
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
                                                      </table>
                                                  </div></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td height="10"><img height="10" width="1"></td>
                                          </tr>

                                        </table>
                                        <table width="86%" border="0" cellpadding="0" cellspacing="0">
                                          <tr> <td height="10"><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                               
                                                <tr>
                                                  <td width="11"><img border="0" src="edit_picture.png" width="22" height="22"></td>
                                                  <td class="style14"><font face="Verdana" size="1"><b> &nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=5','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=900,height=350'); return false;" href="#" style="text-decoration: none"><font color="#000000" class="botoes">Cadastrar 
                                                    Imagem</font></a> </b></font> <font face="Arial" size="1"> (Tamanho em 
                                                      pixels: 96x96 )
                                                      <input name="img5" type="hidden" value="<%=(atualiza.Fields.Item("img5").Value)%>" size="19">
                                                  </font></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><table cellSpacing="0" cellPadding="0" width="100%" border="0">
                                                <tr>
                                                  <td vAlign="top" align="left" width="25%"><img height="176" src="<%=(atualiza.Fields.Item("img5").Value)%>" width="213" border="0" name=img5></td>
                                                  <td vAlign="top" align="left" width="68%"><div align="justify">
                                                      <table border="0" width="100%">
                                                      <tr>
                                                          <td width="44" class="style22"> Titulo:</td>
                                                          <td width="199"><input name="titulo5" type="text" value="<%=(atualiza.Fields.Item("titulo5").Value)%>" size="100"></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style10"> Resumo:</td>
                                                          <td><textarea rows="3" name="resumo5" cols="28"><%=(atualiza.Fields.Item("resumo5").Value)%></textarea></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style11"> Link:</td>
                                                          <td><select size="1" name="link5">
                                                            <%
While (NOT cat.EOF)
%>
                                                            <option value="<%=(cat.Fields.Item("id").Value)%>" <%If (Not isNull((cat.Fields.Item("titulo").Value))) Then If (CStr(cat.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("link5").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Lazer)&nbsp;-&nbsp;<%=(cat.Fields.Item("titulo").Value)%></option>
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
                                                      </table>
                                                  </div></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td height="10"><img height="10" width="1"></td>
                                          </tr>

                                        </table>
                                        <table width="86%" border="0" cellpadding="0" cellspacing="0">
                                          <tr> <td height="10"><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                               
                                                <tr>
                                                  <td width="11"><img border="0" src="edit_picture.png" width="22" height="22"></td>
                                                  <td class="style14"><font face="Verdana" size="1"><b> &nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=6','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=900,height=350'); return false;" href="#" style="text-decoration: none"><font color="#000000" class="botoes">Cadastrar 
                                                    Imagem</font></a> </b></font> <font face="Arial" size="1"> (Tamanho em 
                                                      pixels: 96x96 )
                                                      <input name="img6" type="hidden" value="<%=(atualiza.Fields.Item("img6").Value)%>" size="19">
                                                  </font></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><table cellSpacing="0" cellPadding="0" width="100%" border="0">
                                                <tr>
                                                  <td vAlign="top" align="left" width="25%"><img height="176" src="<%=(atualiza.Fields.Item("img6").Value)%>" width="213" border="0" name=img6></td>
                                                  <td vAlign="top" align="left" width="68%"><div align="justify">
                                                      <table border="0" width="100%">
                                                      <tr>
                                                          <td width="44" class="style22"> Titulo:</td>
                                                          <td width="199"><input name="titulo6" type="text" value="<%=(atualiza.Fields.Item("titulo6").Value)%>" size="100"></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style10"> Resumo:</td>
                                                          <td><textarea rows="3" name="resumo6" cols="28"><%=(atualiza.Fields.Item("resumo6").Value)%></textarea></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style11"> Link:</td>
                                                          <td><select size="1" name="link6">
                                                            <%
While (NOT cat.EOF)
%>
                                                            <option value="<%=(cat.Fields.Item("id").Value)%>" <%If (Not isNull((cat.Fields.Item("titulo").Value))) Then If (CStr(cat.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("link6").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Lazer)&nbsp;-&nbsp;<%=(cat.Fields.Item("titulo").Value)%></option>
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
                                                      </table>
                                                  </div></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td height="10"><img height="10" width="1"></td>
                                          </tr>

                                        </table>
                                        <table width="86%" border="0" cellpadding="0" cellspacing="0">
                                          <tr> <td height="10"><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                               
                                                <tr>
                                                  <td width="11"><img border="0" src="edit_picture.png" width="22" height="22"></td>
                                                  <td class="style14"><font face="Verdana" size="1"><b> &nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=7','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=900,height=350'); return false;" href="#" style="text-decoration: none"><font color="#000000" class="botoes">Cadastrar 
                                                    Imagem</font></a> </b></font> <font face="Arial" size="1"> (Tamanho em 
                                                      pixels: 96x96 )
                                                      <input name="img7" type="hidden" value="<%=(atualiza.Fields.Item("img7").Value)%>" size="19">
                                                  </font></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><table cellSpacing="0" cellPadding="0" width="100%" border="0">
                                                <tr>
                                                  <td vAlign="top" align="left" width="25%"><img height="176" src="<%=(atualiza.Fields.Item("img7").Value)%>" width="213" border="0" name=img7></td>
                                                  <td vAlign="top" align="left" width="68%"><div align="justify">
                                                      <table border="0" width="100%">
                                                      <tr>
                                                          <td width="44" class="style22"> Titulo:</td>
                                                          <td width="199"><input name="titulo7" type="text" value="<%=(atualiza.Fields.Item("titulo7").Value)%>" size="100"></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style10"> Resumo:</td>
                                                          <td><textarea rows="3" name="resumo7" cols="28"><%=(atualiza.Fields.Item("resumo7").Value)%></textarea></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style11"> Link:</td>
                                                          <td><select size="1" name="link7">
                                                            <%
While (NOT cat.EOF)
%>
                                                            <option value="<%=(cat.Fields.Item("id").Value)%>" <%If (Not isNull((cat.Fields.Item("titulo").Value))) Then If (CStr(cat.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("link7").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Lazer)&nbsp;-&nbsp;<%=(cat.Fields.Item("titulo").Value)%></option>
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
                                                      </table>
                                                  </div></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td height="10"><img height="10" width="1"></td>
                                          </tr>

                                        </table>
                                        <table width="86%" border="0" cellpadding="0" cellspacing="0">
                                          
                                          <tr>
                                            
                                          <tr>
                                             <td><center><strong class="text-primary fa-2x verde">------</strong> </center></td></tr>
                                             <tr>
                                             <td><center><strong class="text-primary fa-2x verde">Su�tes em destaque</strong> </center></td></tr>
                                             <tr>
                                             <td><center><strong class="text-primary fa-2x verde">------</strong> </center></td></tr>
                                             
                                            <td height="10"><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                                <td width="11"><img border="0" src="edit_picture.png" width="22" height="22"></td>
                                                  <td><b> <font face="Verdana" size="1">&nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=8','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=200'); return false;" href="#" style="text-decoration: none"><font color="#000000" class="b-alterar">Cadastrar 
                                                    Imagem</font></a> </font></b> <font face="Arial" size="1"> (Tamanho em 
                                                      pixels: 96x93 )
                                                      <input name="img8" type="hidden" id="img8" value="<%=(atualiza.Fields.Item("img8").Value)%>" size="19">
                                                  </font></td>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><table cellSpacing="0" cellPadding="0" width="100%" border="0">
                                                <tr>
                                                  <td vAlign="top" align="left" width="25%"><img height="176" src="<%=(atualiza.Fields.Item("img8").Value)%>" width="213" border="0" name=img8></td>
                                                  <td vAlign="top" align="left" width="68%"><div align="justify">
                                                      <table border="0" width="100%">
                                                      <tr>
                                                          <td width="44" class="style22"> Titulo:</td>
                                                          <td width="199"><input type="text" name="titulo8" size="27" value="<%=(atualiza.Fields.Item("titulo8").Value)%>"></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style10"> Resumo:</td>
                                                          <td><textarea rows="3" name="resumo8" cols="32"><%=(atualiza.Fields.Item("resumo8").Value)%></textarea></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style11"> Link:</td>
                                                          <td><select size="1" name="link8">
                                                            <%
While (NOT cat1.EOF)
%>
                                                            <option value="<%=(cat1.Fields.Item("id").Value)%>" <%If (Not isNull((cat1.Fields.Item("titulo").Value))) Then If (CStr(cat1.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("link8").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Su&iacute;tes)&nbsp;-&nbsp;<%=(cat1.Fields.Item("titulo").Value)%></option>
                                                            <%
  cat1.MoveNext()
Wend
If (cat1.CursorType > 0) Then
  cat1.MoveFirst
Else
  cat1.Requery
End If
%>
                                                          </select></td>
                                                        </tr>
                                                      </table>
                                                  </div></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td height="10"><img height="10" width="1"></td>
                                          </tr>
                                          <tr>
                                            <td height="10"><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                                <tr>
                                                  <td width="11"><img border="0" src="edit_picture.png" width="22" height="22"></td>
                                                  <td class="style14"><b> <font face="Verdana" size="1">&nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=9','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=200'); return false;" href="#" style="text-decoration: none"><font color="#000000" class="b-alterar">Cadastrar 
                                                    Imagem</font></a> </font></b> <font face="Arial" size="1"> (Tamanho em 
                                                      pixels:96x96 )
                                                      <input name="img9" type="hidden" id="img9" value="<%=(atualiza.Fields.Item("img9").Value)%>" size="19">
                                                  </font></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><table cellSpacing="0" cellPadding="0" width="100%" border="0">
                                                <tr>
                                                  <td vAlign="top" align="left" width="25%"><img height="176" src="<%=(atualiza.Fields.Item("img9").Value)%>" width="213" border="0" name=img9></td>
                                                  <td vAlign="top" align="left" width="68%"><div align="justify">
                                                      <table border="0" width="100%">
                                                      <tr>
                                                          <td width="44" class="style22"> Titulo:</td>
                                                          <td><input type="text" name="titulo9" size="27" value="<%=(atualiza.Fields.Item("titulo9").Value)%>"></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style10"> Resumo:</td>
                                                          <td><textarea rows="3" name="resumo9" cols="28"><%=(atualiza.Fields.Item("resumo9").Value)%></textarea></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style11"> Link:</td>
                                                          <td><select size="1" name="link9">
                                                            <%
While (NOT cat1.EOF)
%>
                                                            <option value="<%=(cat1.Fields.Item("id").Value)%>" <%If (Not isNull((cat1.Fields.Item("titulo").Value))) Then If (CStr(cat1.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("link9").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Su&iacute;tes)&nbsp;-&nbsp;<%=(cat1.Fields.Item("titulo").Value)%></option>
                                                            <%
  cat1.MoveNext()
Wend
If (cat1.CursorType > 0) Then
  cat1.MoveFirst
Else
  cat1.Requery
End If
%>
                                                          </select></td>
                                                        </tr>
                                                      </table>
                                                  </div></td>
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
                                                  <td><b> <font face="Verdana" size="1">&nbsp;<a onClick="window.open('uploadimagemchamadaprincipal_nova.asp?img=10','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=200'); return false;" href="#" style="text-decoration: none"><font color="#000000" class="b-alterar">Cadastrar 
                                                    Imagem</font></a> </font></b> <font face="Arial" size="1"> (Tamanho em 
                                                      pixels: 96x96 )
                                                      <input name="img10" type="hidden" id="img10" value="<%=(atualiza.Fields.Item("img10").Value)%>" size="19">
                                                  </font></td>
                                                </tr>
                                            </table></td>
                                          </tr>
                                          <tr>
                                            <td><table cellSpacing="0" cellPadding="0" width="100%" border="0">
                                                <tr>
                                                  <td vAlign="top" align="left" width="32%"><img height="176" src="<%=(atualiza.Fields.Item("img10").Value)%>" width="213" border="0" name=img10></td>
                                                  <td vAlign="top" align="left" width="68%"><div align="justify">
                                                      <table border="0" width="100%">
                                                      <tr>
                                                          <td width="44" class="style22"> Titulo:</td>
                                                          <td><input type="text" name="titulo10" size="20" value="<%=(atualiza.Fields.Item("titulo10").Value)%>"></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style10"> Resumo:</td>
                                                          <td><textarea rows="3" name="resumo10" cols="28"><%=(atualiza.Fields.Item("resumo10").Value)%></textarea></td>
                                                        </tr>
                                                        <tr>
                                                          <td width="44" class="style11"> Link:</td>
                                                          <td><select size="1" name="link10">
                                                            <%
While (NOT cat1.EOF)
%>
                                                            <option value="<%=(cat1.Fields.Item("id").Value)%>" <%If (Not isNull((cat1.Fields.Item("titulo").Value))) Then If (CStr(cat1.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("link10").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Su&iacute;tes)&nbsp;-&nbsp;<%=(cat1.Fields.Item("titulo").Value)%></option>
                                                            <%
  cat1.MoveNext()
Wend
If (cat1.CursorType > 0) Then
  cat1.MoveFirst
Else
  cat1.Requery
End If
%>
                                                          </select></td>
                                                        </tr>
                                                        
                                                      </table>
                                                  </div></td>
                                                </tr>
                                            </table></td>
                                          </tr>

                                </table>                                        </td>
                              </tr>
                                  </table>                                  �</td>
                              </tr>
                              <tr>
                                <td width="728" bgcolor=""><p align="center">
                                    <input type="submit" value="Atualizar P&aacute;gina Inicial" name="B1" style="color: #FFFFFF; border: 1px solid #CC3300; background-color: #CC3300">
                                </td>
                              </tr>
                            </table>
                          </div></td>
                        </tr>
                      </table>
                  </div></td>
                </tr>
                
                
                
              </table>
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
</body>

</html>
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
<%
cat.Close()
Set cat = Nothing
%>
<%
cat1.Close()
Set cat1 = Nothing
%>


