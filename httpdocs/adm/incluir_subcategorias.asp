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
  MM_editTable = "subcategorias"
  MM_editRedirectUrl = "listar_subcategorias.asp"
  MM_fieldsStr  = "categoria|value|subcategoria|value|texto|value"
  MM_columnsStr = "categoria|',none,''|subcategoria|',none,''|texto|',none,''"

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
%><head>
<!-- 
'Desenvolvedor: AOL | Agência Online
'Telefone: 11 984-000-894 (TIM)
'Desenvolvimento:  20/06/2015
'Email: contato@aolsite.com.br
-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
 <title>AOL | Agência online</title>
 <link href="css/aol.css" rel="stylesheet" type="text/css">
 <link href="css/icones.css" rel="stylesheet" type="text/css">

</head>
 
 





<link rel="icon" type="image/ico" href="images/fav.ico">
<body>


<div class="header">
<div class="menu"> <!--#include file="menunovo.asp" -->
</aol>

</div></div>

<div class="conteudo"> <div class="pading">
  
  <r><a href="incluir_subcategorias.asp"><span class=" icon-mais" ></span> <br><br> Cadastrar</a></r><br>
<br>
<div class="titulo"><strong> &nbsp;&nbsp;&nbsp;
GERÊNCIAR SUB-CATEGORIAS</strong></div>
  <div class="clr"></div> 
<aol class="box">


<!--- conteudo programado ------->


 <form method="POST" action="<%=MM_editAction%>" name="form1">

	<input type="hidden" name="img5" size="19"><input type="hidden" name="img4" size="19"><input type="hidden" name="img3" size="19"><input type="hidden" name="img2" size="19"><input type="hidden" name="img1" size="19">
												<input type="hidden" name="data" size="19" value="<%=date%>">
												<table border="0" style="border-collapse: collapse" width="90%" cellpadding="0">
												  <tr>
												    <td width="579" colspan="2"><img border="0" src="seta.gif" width="6" height="6"> <span style="text-transform: uppercase"> <b><font face="Verdana">INSERIR 
												      SUB-CATEGORIA</font></b></span></td>
											      </tr>
												  <tr>
												    <td width="579" colspan="2">&nbsp;</td>
											      </tr>
												  <tr>
												    <td background="bg_pontos.gif" width="579" colspan="2"><img border="0" src="inclui1.gif"></td>
											      </tr>
												  <tr>
												    <td width="579" colspan="2">&nbsp;</td>
											      </tr>
												  <tr>
												    <td width="124" height="22"><p align="right">Categoria:&nbsp;&nbsp; </td>
												    <td width="455" height="22"><select size="1" name="categoria" class="input">
												      <%
While (NOT cat.EOF)
%>
												      <option value="<%=(cat.Fields.Item("id").Value)%>" <%If (Not isNull((cat.Fields.Item("nome").Value))) Then If (CStr(cat.Fields.Item("nome").Value) = CStr((cat.Fields.Item("nome").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(cat.Fields.Item("nome").Value)%></option>
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
												    <td width="124"><p align="right">Sub-Categoria:&nbsp;&nbsp; </td>
												    <td width="455"><input type="text" name="subcategoria" size="38"></td>
											      </tr>
												  <tr>
												    <td width="579" colspan="2">&nbsp;</td>
											      </tr>
												  <tr>
												    <td width="579" colspan="2"><table border="0" style="border-collapse: collapse" width="100%">
												      <tr>
												        <td bgcolor="#FFFFFF"><p align="center">                                                        
												          <p align="center">                                                          
												          <p align="center">                                                          
												          <p align="center">                                                          
												          <p align="center">                                                          
												          <p align="center">                                                          
												          <div align="left">Texto<br>
											              </div>
												          <div align="left">
												            <textarea name="texto" class="input" id="texto_materia" rows="1" style="width:550px;"></textarea>
											              </div></td>
											          </tr>
												      <tr>
												        <td>&nbsp;</td>
											          </tr>
												      </table></td>
											      </tr>
												  <tr>
												    <td width="579" colspan="2">&nbsp;</td>
											      </tr>
												  <tr>
												    <td width="579" colspan="2">&nbsp;</td>
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


 </aol></div></div> </body> </html>
<%
cat.Close()
Set cat = Nothing
%>


 <style>
  input, .input{ background: #ccc; border-radius:5px; margin:5px; padding:10px;}
  
 .botaoverde{ background:#093; color:#FFF;}
 .botaoverde:hover{ background: #84CA02;color:#FFF;}
 
 
 </style>
 
 