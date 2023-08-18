
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
cat.Source = "SELECT *  FROM pgpacotes ORDER BY pactitulo ASC"
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
  MM_editTable = "pacotes"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "edit_pacotes1.asp"
  MM_fieldsStr  = "pactitulo1|value|pacpreco1|value|pacresumo1|value|paclink1|value|pactitulo2|value|pacpreco2|value|pacresumo2|value|paclink2|value|pactitulo3|value|pacpreco3|value|pacresumo3|value|paclink3|value"
  MM_columnsStr = "pactitulo1|',none,''|pacpreco1|',none,''|pacresumo1|',none,''|paclink1|',none,''|pactitulo2|',none,''|pacpreco2|',none,''|pacresumo2|',none,''|paclink2|',none,''|pactitulo3|',none,''|pacpreco3|',none,''|pacresumo3|',none,''|paclink3|',none,''"

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
atualiza.Source = "SELECT *  FROM pacotes"
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
<strong class="text-primary fa-2x azul">  &Xi; P&aacute;gina Pacotes em destaque</strong><br>
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
                          <td><div align="left">
                      <table width="99%" border="0" cellspacing="0" cellpadding="10">
                        <tr>
                          <td><div align="center">
                            <table border="0" style="border-collapse: collapse" width="728" cellpadding="0">
                              
                              <tr>
                                <td width="728"><table border="0" width="100%" cellpadding="6" bgcolor="">
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
                                      <td vAlign="top" width="275"><table cellSpacing="0" cellPadding="0" width="100%" border="0">
                                        <tr>
                                          <td><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                            <tr>
                                              <td width="11">&nbsp;</td>
                                              </tr>
                                            </table></td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          </tr>
                                        <tr>
                                          <td height="19" class="style27"><strong class="text-primary fa-2x verde"><center>Pacote 1</center></strong></td>
                                          <td class="style27">&nbsp;</td>
                                          <td class="style27">&nbsp;</td>
                                          </tr>
                                        <tr>
                                          <td class="style10">&nbsp;</td>
                                          <td class="style10">&nbsp;</td>
                                          <td class="style10">&nbsp;</td>
                                          </tr>
                                        <tr>
                                          <td><table border="0" width="100%">
                                            <tr>
                                              <td width="52" class="style22">Titulo Pacote:</td>
                                              <td><input type="text" name="pactitulo1" size="20" value="<%=(atualiza.Fields.Item("pactitulo1").Value)%>"></td>
                                            </tr>
                                            <tr>                                          
                                            <tr>
                                              <td width="52" class="style30">Pre&ccedil;o destaque:</td>
                                              <td><input type="text" name="pacpreco1" size="20" value="<%=(atualiza.Fields.Item("pacpreco1").Value)%>"></td>
                                            </tr>
  <td width="52" class="style10"> Resumo:</td>
    <td><textarea rows="5" name="pacresumo1" cols="35"><%=(atualiza.Fields.Item("pacresumo1").Value)%></textarea></td>
  </tr>
  <tr>
    <td width="52" class="style11"> Link:</td>
    <td><select size="1" name="paclink">
      <%
While (NOT cat.EOF)
%>
      <option value="<%=(cat.Fields.Item("id").Value)%>" <%If (Not isNull((cat.Fields.Item("pactitulo").Value))) Then If (CStr(cat.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("paclink1").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Pacote)&nbsp;-&nbsp;<%=(cat.Fields.Item("pactitulo").Value)%></option>
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
                                          </table></td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          </tr>
                                        <tr>
                                          <td><img height="10" width="1"></td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          </tr>
                                        <tr>
                                          <td><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                            
                                            </table></td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          </tr>
                                        <tr>
                                          <td class="style27"><strong class="text-primary fa-2x verde"><center>Pacote 2</center></strong></td>
                                          <td class="style27">&nbsp;</td>
                                          <td class="style27">&nbsp;</td>
                                          </tr>
                                        <tr>
                                          <td class="style10">&nbsp;</td>
                                          <td class="style10">&nbsp;</td>
                                          <td class="style10">&nbsp;</td>
                                          </tr>
                                        <tr>
                                          <td><table border="0" width="100%">
                                            <tr>
                                              <td width="52" class="style22">Titulo Pacote:</td>
                                              <td><input type="text" name="pactitulo2" size="20" value="<%=(atualiza.Fields.Item("pactitulo2").Value)%>"></td>
                                              </tr>
                                            <tr>
                                              <tr>
                                                <td width="52" class="style30">Pre&ccedil;o destaque:</td>
                                                <td><input type="text" name="pacpreco2" size="20" value="<%=(atualiza.Fields.Item("pacpreco2").Value)%>"></td>
                                                </tr>
                                            <td width="52" class="style10"> Resumo:</td>
                                              <td><textarea rows="5" name="pacresumo2" cols="35"><%=(atualiza.Fields.Item("pacresumo2").Value)%></textarea></td>
                                              </tr>
                                            <tr>
                                              <td width="52" class="style11"> Link:</td>
                                              <td><select size="1" name="paclink2">
                                                            <%
While (NOT cat.EOF)
%>
                                                            <option value="<%=(cat.Fields.Item("id").Value)%>" <%If (Not isNull((cat.Fields.Item("pactitulo").Value))) Then If (CStr(cat.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("paclink2").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Pacote)&nbsp;-&nbsp;<%=(cat.Fields.Item("pactitulo").Value)%></option>
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
                                            </table></td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          </tr>
                                        <tr>
                                          <td><img height="10" width="1"></td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          </tr>
                                        <tr>
                                          <td height="10"><table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="">
                                            <tr>
                                              <td width="11"><strong class="style27"><strong class="text-primary fa-2x verde"><center>Pacote 3</center></strong></strong></td>
                                              </tr>
                                            </table></td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          </tr>
                                        
                                        <tr>
                                          <td><table cellSpacing="0" cellPadding="0" width="100%" border="0">
                                            <tr>
                                              <td align="left" vAlign="top">&nbsp;</td>
                                              </tr>
                                            <tr>
                                              <td align="left" vAlign="top"><div align="justify">
                                                <table border="0" width="100%">
                                                  <tr>
                                                    <td width="52" class="style22"> Titulo Pacote:</td>
                                                    <td><input type="text" name="pactitulo3" size="20" value="<%=(atualiza.Fields.Item("pactitulo3").Value)%>"></td>
                                                    </tr>
                                                  <tr>
                                                    <tr>
                                                      <td width="52" class="style30"> Pre&ccedil;o destaque:</td>
                                                      <td><input type="text" name="pacpreco3" size="20" value="<%=(atualiza.Fields.Item("pacpreco3").Value)%>"></td>
                                                      </tr>
                                                  <td width="52" class="style10"> Resumo:</td>
                                                    <td><textarea rows="3" name="pacresumo3" cols="35"><%=(atualiza.Fields.Item("pacresumo3").Value)%></textarea></td>
                                                    </tr>
                                                  <tr>
                                                    <td width="52" class="style11"> Link:</td>
                                                    <td><select size="1" name="paclink3">
                                                            <%
While (NOT cat.EOF)
%>
                                                            <option value="<%=(cat.Fields.Item("id").Value)%>" <%If (Not isNull((cat.Fields.Item("pactitulo").Value))) Then If (CStr(cat.Fields.Item("id").Value) = CStr((atualiza.Fields.Item("paclink3").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>>(Pacote)&nbsp;-&nbsp;<%=(cat.Fields.Item("pactitulo").Value)%></option>
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
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          </tr>
                                      </table></td>
                                      </tr>
                                  </table></td>
                              </tr>
                              
                              <tr>
                              <tr>
                                              <td width="11">&nbsp;</td>
                                              </tr>
                                              <tr>
                                              <td width="11">&nbsp;</td>
                                              </tr>
                                <td width="100%" bgcolor=""><p align="center">
                                    <input type="submit" value="Atualizar Pacotes" name="B1" style="color: ; border: 1px solid #CC3300; background-color: #CC3300">
                                </td>
                              </tr>
                            </table>
                          </div></td>
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
<%
cat.Close()
Set cat = Nothing
%>
<%
cat1.Close()
Set cat1 = Nothing
%>