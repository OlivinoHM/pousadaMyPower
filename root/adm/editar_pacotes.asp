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
  MM_editTable = "pgpacotes"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "listar_pacotes.asp"
  MM_fieldsStr  = "pactitulo|value|datad1|value|horad1|value|dataa1|value|horaa1|value|datad2|value|horad2|value|dataa2|value|horaa2|value|datad3|value|horad3|value|dataa3|value|horaa3|value|2psap|value|2pcap|value|3psap|value|3pcap|value|4psap|value|4pcap|value|6pcap|value|10pcap|value|16pcap|value|promocao1|value|promocaop1|value|promocao2|value|promocaop2|value|obs1|value"
  MM_columnsStr = "pactitulo|',none,''|datad1|',none,''|horad1|',none,''|dataa1|',none,''|horaa1|',none,''|datad2|',none,''|horad2|',none,''|dataa2|',none,''|horaa2|',none,''|datad3|',none,''|horad3|',none,''|dataa3|',none,''|horaa3|',none,''|2psap|',none,''|2pcap|',none,''|3psap|',none,''|3pcap|',none,''|4psap|',none,''|4pcap|',none,''|6pcap|',none,'|10pcap|',none,''|16pcap|',none,''|promocao1|',none,''|promocaop1|',none,''|promocao2|',none,''|promocaop2|',none,''|obs1|',none,''"

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
atualiza.Source = "SELECT *  FROM pgpacotes  WHERE id = " + Replace(atualiza__MMColParam, "'", "''") + ""
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
<strong class="text-primary fa-2x azul">  &Xi; Editar Pacotes</strong><br>
 
  

</div><br>
<br>


	<form method="POST" action="<%=MM_editAction%>" name="form1">

	<input type="hidden" name="img5" size="19"><input type="hidden" name="img4" size="19"><input type="hidden" name="img3" size="19"><input type="hidden" name="img2" size="19"><input type="hidden" name="img1" size="19">
	 <input type="hidden" name="data" size="19" value="<%=date%>">

<tr>
                  <td width="579" colspan="4"><center>Nome do pacote</center></td>
            <tr>
                  <td width="579" colspan="4"><center><center><input name="pactitulo" type="text" value="<%=(atualiza.Fields.Item("pactitulo").Value)%>" size="30"></center></center></td>
            </tr>

     <table width="900" border="0" cellpadding="0" cellspacing="3" align="center">
       <tr>
         <td colspan="5"><center>
           <br></center>
           </td>
       </tr>
       <tr>
         <td width="24">&nbsp;</td>
         <td width="212">&nbsp;</td>
         <td width="212">&nbsp;</td>
         <td width="212">&nbsp;</td>
         <td width="212">&nbsp;</td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td>&nbsp;&nbsp;Data Entrada</td>
         <td>&nbsp;&nbsp;Hora entrada</td>
         <td>&nbsp;&nbsp;Data sa&iacute;da</td>
         <td>&nbsp;&nbsp;Hora sa&iacute;da</td>
       </tr>
       <tr>
         <td><center>OP&Ccedil;&Atilde;O 1</center></td>
         <td><input name="datad1" type="text" value="<%=(atualiza.Fields.Item("datad1").Value)%>" width="200"></td>
         <td><input name="horad1" type="text" value="<%=(atualiza.Fields.Item("horad1").Value)%>"  width="200"></td>
         <td><input name="dataa1" type="text" value="<%=(atualiza.Fields.Item("dataa1").Value)%>" width="200"></td>
         <td><input name="horaa1" type="text" value="<%=(atualiza.Fields.Item("horaa1").Value)%>" width="200"></td>
       </tr>
       <tr>
         <td><center>OP&Ccedil;&Atilde;O 2</center></td>
         <td><input name="datad2" type="text" value="<%=(atualiza.Fields.Item("datad2").Value)%>" size="38"></td>
         <td><input name="horad2" type="text" value="<%=(atualiza.Fields.Item("horad2").Value)%>" size="38"></td>
         <td><input name="dataa2" type="text" value="<%=(atualiza.Fields.Item("dataa2").Value)%>" size="38"></td>
         <td><input name="horaa2" type="text" value="<%=(atualiza.Fields.Item("horaa2").Value)%>" size="38"></td>
       </tr>
       <tr>
         <td><center>OP&Ccedil;&Atilde;O 3</center></td>
         <td><input name="datad3" type="text" value="<%=(atualiza.Fields.Item("datad3").Value)%>" size="38"></td>
         <td><input name="horad3" type="text" value="<%=(atualiza.Fields.Item("horad3").Value)%>" size="38"></td>
         <td><input name="dataa3" type="text" value="<%=(atualiza.Fields.Item("dataa3").Value)%>" size="38"></td>
         <td><input name="horaa3" type="text" value="<%=(atualiza.Fields.Item("horaa3").Value)%>" size="38"></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
       </tr>
       
     </table>
     <table border="0" style="border-collapse: collapse" width="900" cellpadding="0">
			  <tr>
			    <td width="579">&nbsp;</td>
			    <td width="579">&nbsp;</td>
		      </tr>
			  <tr>
			    <td width="579"><p align="center">2 Pessoas sem Ar:&nbsp;&nbsp; </td>
			    <td width="579"><p align="center">2 Pessoas com Ar:&nbsp;&nbsp; </td>
		      </tr>
                <tr>
                <td width="455"><center><input name="2psap" type="text" value="<%=(atualiza.Fields.Item("2psap").Value)%>" size="20"></center></td>
                <td width="455"><center>
                  <input name="2pcap" type="text" value="<%=(atualiza.Fields.Item("2pcap").Value)%>" size="20">
                </center></td>
              </tr>
			 
			   <tr>
			    <td width="579"><p align="center">&nbsp; </td>
			    <td width="579"><p align="center">&nbsp; </td>
		      </tr>
              <tr>
			    <td width="579"><p align="center">3 Pessoas sem Ar:&nbsp;&nbsp; </td>
			    <td width="579"><p align="center">3 Pessoas com Ar:&nbsp;&nbsp; </td>
		      </tr>
			<tr>    <td width="455"><center>
			  <input name="3psap" type="text" value="<%=(atualiza.Fields.Item("3psap").Value)%>" size="20">
			</center></td>
			  <td width="455"><center>
			    <input name="3pcap" type="text" value="<%=(atualiza.Fields.Item("3pcap").Value)%>" size="20">
			  </center></td>
			  </tr>
			  
			  <tr>
			    <td width="579"><p align="center">&nbsp; </td>
			    <td width="579"><p align="center">&nbsp; </td>
		      </tr>
              <tr>
			    <td width="579"><p align="center">4 Pessoas sem Ar:&nbsp;&nbsp; </td>
			    <td width="579"><p align="center">4 Pessoas com Ar:&nbsp;&nbsp; </td>
		      </tr>
			<tr>    <td width="455"><center>
			  <input name="4psap" type="text" value="<%=(atualiza.Fields.Item("4psap").Value)%>" size="20">
			</center></td>
			  <td width="455"><center>
			    <input name="4pcap" type="text" value="<%=(atualiza.Fields.Item("4pcap").Value)%>" size="20">
			  </center></td>
			  </tr>
			 <tr>
			    <td width="579"><p align="center">&nbsp; </td>
			    <td width="579"><p align="center">&nbsp; </td>
		      </tr>
              
			  <tr>
			    <td width="579"><p align="center">6 Pessoas com Ar:&nbsp;&nbsp; </td>
			    <td width="579"><p align="center">10 Pessoas com Ar:&nbsp;&nbsp; </td>
		      </tr>
			<tr>    <td width="455"><center><input name="6pcap" type="text" value="<%=(atualiza.Fields.Item("6pcap").Value)%>" size="20"></center></td>
			  <td width="455"><center>
			    <input name="10pcap" type="text" value="<%=(atualiza.Fields.Item("10pcap").Value)%>" size="20">
			    </center></td>
			  </tr>
			  </table>
              
              <table border="0" style="border-collapse: collapse" width="900" cellpadding="0" align="center">
			  <tr>
			    <td width="579"><p align="center">&nbsp; </td>
			    
		      </tr>
              <tr>
			    <td width="579"><p align="center">16 Pessoas sem Ar:&nbsp;&nbsp; </td>
			    </tr>
			<tr>    <td width="455"><center>
			  <table width="100%" border="0" cellpadding="0" cellspacing="0">
			    <tr>
			      <td width="35%">&nbsp;</td>
			      <td width="59%"><input name="16pcap" type="text" value="<%=(atualiza.Fields.Item("16pcap").Value)%>" size="20"></td>
			      <td width="6%">&nbsp;</td>
			      </tr>
			    </table></center></td>
			  
		      </tr>
			  </table>
              <table border="0" style="border-collapse: collapse" width="900" cellpadding="0" align="center">
			  <tr>
			    <td width="579"><p align="center">&nbsp;</td>
                <td width="579"><p align="center">&nbsp;</td>
			    
		      </tr>
              <tr>
			    <td width="579"><p align="center">&nbsp;</td>
                <td width="579"><p align="center">&nbsp;</td>
			    
		      </tr>
              <tr>
			    <td width="579"><p align="center">Promo&ccedil;&atilde;o 1 Descri&ccedil;&atilde;o:</td>
			    <td width="579"><p align="center">Promoc&atilde;o 1 Pre&ccedil;o:</td>
		      </tr>
			<tr>    <td width="455"><center><input name="promocao1" type="text" value="<%=(atualiza.Fields.Item("promocao1").Value)%>" size="20"></center></td>
			  <td width="455"><center>
			    <input name="promocaop1" type="text" value="<%=(atualiza.Fields.Item("promocaop1").Value)%>" size="20">
			    </center></td>
			  </tr>
              <tr>
			    <td width="579"><p align="center">&nbsp; </td>
			    <td width="579"><p align="center">&nbsp; </td>
		      </tr>
              <tr>
			    <td width="579"><p align="center">Promo&ccedil;&atilde;o 2 Descri&ccedil;&atilde;o:</td>
			    <td width="579"><p align="center">Promoc&atilde;o 2 Pre&ccedil;o:</td>
		      </tr>
			<tr>    <td width="455"><center><input name="promocao2" type="text" value="<%=(atualiza.Fields.Item("promocao2").Value)%>" size="20"></center></td>
			  <td width="455"><center>
			    <input name="promocaop2" type="text" value="<%=(atualiza.Fields.Item("promocaop2").Value)%>" size="20">
			    </center></td>
			  </tr>
		    </table>
              <table border="0" style="border-collapse: collapse" width="900" cellpadding="0" align="center">
			 
                  <tr>
                  <td width="579" colspan="4"><center>&nbsp;</center></td>
                  <tr>
                  <td width="579" colspan="4"><center>&nbsp;</center></td>
			  <tr>
			    <td width="579" colspan="4"><center>
			        Obrsevações do pacote:
			    </center></td>
			  <tr>
			    <td width="579" colspan="4"><center>
		        <textarea name="obs1" cols="50" rows="5" class="input" id="horad" style="width:550px;"><%=(atualiza.Fields.Item("obs1").Value)%></textarea></center></td>
                 <tr>
                  <td width="579" colspan="4"><center>&nbsp;</center></td>
                  <tr>
                  <td width="579" colspan="4"><center>&nbsp;</center></td>
                 
                  <tr>
                  <td width="579" colspan="4"><center>
			      <input type="submit" value="Gravar e Visualizar" name="B1"></center></td>
		      </tr>
			  </table>

    <input type="hidden" name="MM_update" value="form1">
<input type="hidden" name="MM_recordId" value="<%= atualiza.Fields.Item("id").Value %>"> </form>


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