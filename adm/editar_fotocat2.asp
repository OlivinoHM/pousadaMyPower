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
  MM_editTable = "categoria"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "listar_categoria.asp"
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
Recordset1.Source = "SELECT *  FROM categoria  WHERE id = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>

<html>

<head>
<LINK 
href="css/estilos_adm.css" type=text/css 
rel=STYLESHEET>
<title></title>
</head>

<body bgcolor="#E6E6E6" topmargin="0">


	<form method="POST" action="<%=MM_editAction%>" name="form1">

	<table height="27%" cellSpacing="0" cellPadding="0" width="100%" border="0">
		<tr>
			<td vAlign="top" align="middle">
			<table border="0" style="border-collapse: collapse" width="741" cellpadding="0">
				<tr>
					<td width="739" align="right"></td>
				</tr>
			</table>
			<table cellSpacing="0" cellPadding="0" width="775" align="center" bgColor="#ffffff" background="bg_pag.gif" border="0">
				<tr>
					<td width="775">
					<img height="30" alt="" src="borda_up.gif" width="775"></td>
				</tr>
				<tr>
					<td width="775">
					<div align="center">
						<table cellSpacing="0" cellPadding="0" width="728" align="center" border="0">
							<tr>
								<td width="15">&nbsp;</td>
								<td width="116" valign="top">
								<table width="100%" border="0">
									<tr>
										<td>
										<p align="center">
										</td>
									</tr>
									<tr>
										<td>&nbsp;
										</td>
									</tr>
								</table>
								</td>
								<td width="12">&nbsp;</td>
								<td vAlign="top" bgColor="#ffffff" width="579">
								<table border="0" style="border-collapse: collapse" width="579" cellpadding="0">
									<tr>
										<td width="579" colspan="4">
										<img border="0" src="seta.gif" width="6" height="6">
										<span style="text-transform: uppercase">
										<b><font face="Verdana">ATUALIZAR 
										FOTO</font></b></span></td>
									</tr>
									<tr>
										<td width="579" colspan="4">&nbsp;</td>
									</tr>
									<tr>
										<td background="bg_pontos.gif" width="579" colspan="4">
										<img border="0" src="inclui1.gif"></td>
									</tr>
									<tr>
										<td width="579" colspan="4">&nbsp;</td>
									</tr>
									<tr>
										<td width="129">
										<p align="right">IMG:&nbsp;&nbsp; </td>
										<td width="34">
												<input type="hidden" name="img" size="4" value="<%=(Recordset1.Fields.Item("img").Value)%>"><input type="hidden" name="categoria" size="10" value="<%=(Recordset1.Fields.Item("categoria").Value)%>">
                          </td>
										<td width="205">
												<img src="<%=(Recordset1.Fields.Item("img").Value)%>" name=img width="230" height="153"></td>
										<td width="211">
												<table border="0" width="86%" bgcolor="#FFFFFF">
													<tr>
												<td width="170" bgColor="#FFFFFF">
												<table border="0" style="border-collapse: collapse" width="100%" cellspacing="1" cellpadding="0" bgcolor="#000000">
													<tr>
														<td width="12">
														<a href="javascript:MM_openBrWindow('listar_categoria.asp?id=33&amp;id1=9','','width=815,height=400,scrollbars=yes')">
														<img alt="Inserir/Listar" src="listar1.gif" border="0" width="16" height="16"></a></td>
														<td bgcolor="#000000">
														<table border="0" style="border-collapse: collapse" width="100%">
															<tr>
																<td>
												<b>
												<font face="Verdana" color="#006699" size="1">
												<a onClick="window.open('imgcat.asp?img=des01&id=1','Janela','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=190'); return false;" href="#">
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
												<p style="MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px">
												(Tamanho: 230x153 pixel)</td>
													</tr>
												</table>
										</td>
									</tr>
									<tr>
										<td width="129" height="25">
										<p align="right">&nbsp;&nbsp; </td>
										<td width="450" height="25" colspan="3">&nbsp;</td>
									</tr>
									<tr>
										<td width="579" colspan="4">&nbsp;</td>
									</tr>
									<tr>
										<td width="579" colspan="4">&nbsp;</td>
									</tr>
									<tr>
										<td width="579" colspan="4">&nbsp;
										</td>
									</tr>
									<tr>
										<td width="579" colspan="4">&nbsp;</td>
									</tr>
									<tr>
										<td width="579" colspan="4">
										&nbsp;&nbsp;
										<input type="submit" value="Gravar e Visualizar" name="B1"></td>
									</tr>
									<tr>
										<td width="579" colspan="4">&nbsp;</td>
									</tr>
									</table>
								<p></td>
							</tr>
						</table>
					</div>
					</td>
				</tr>
				<tr>
					<td width="775">
					<img height="32" alt="" src="borda_down.gif" width="775"></td>
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


</body>

</html><%
Recordset1.Close()
Set Recordset1 = Nothing
%>