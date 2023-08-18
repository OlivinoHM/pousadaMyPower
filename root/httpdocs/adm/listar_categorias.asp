<%@LANGUAGE="VBSCRIPT"%>
<% session.lcid=1046 %>

<!--#include file="Connections/dlfelix.asp" -->

		 <%	
	Set Recordset1 = Server.CreateObject("ADODB.Recordset")
	
    Recordset1.ActiveConnection = MM_dlfelix_STRING
Recordset1.Source = "SELECT *  FROM categorias ORDER BY id desc"
	Recordset1.CursorLocation = 3
	Recordset1.CursorType = 1
	Recordset1.PageSize = 10
	Recordset1.CacheSize = 10
	Recordset1.Open()

	Dim contadorvideos

	If request("pagina") <> "" Then
		Recordset1.AbsolutePage = request("pagina")
		pagina2 = request("pagina")
	Else
		If NOT Recordset1.EOF Then Recordset1.AbsolutePage = 1
		pagina = 1
	End If %>
	
	

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 20
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>

<html>

<head>
<LINK 
href="css/estilos_adm.css" type=text/css 
rel=STYLESHEET>
<title>CATA - Intranet</title>
</head>

<body bgcolor="#E6E6E6" topmargin="0">


	<form method="POST" action="<%=MM_editAction%>" name="form1">

	<input type="hidden" name="img5" size="19"><input type="hidden" name="img4" size="19"><input type="hidden" name="img3" size="19"><input type="hidden" name="img2" size="19"><input type="hidden" name="img1" size="19">
												<input type="hidden" name="data" size="19" value="<%=date%>">

	<table height="27%" cellSpacing="0" cellPadding="0" width="100%" border="0">
		<tr>
			<td vAlign="top" align="middle">
			<table border="0" style="border-collapse: collapse" width="741" cellpadding="0">
				<tr>
					<td width="739" align="right"><!--#include file="menu.asp" --></td>
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
										&nbsp;</td>
									</tr>
									<tr>
										<td>
										&nbsp;</td>
									</tr>
								</table>
								</td>
								<td width="12">&nbsp;</td>
								<td vAlign="top" bgColor="#ffffff" width="579">
								<table border="0" style="border-collapse: collapse" width="579" cellpadding="0">
									<tr>
										<td width="579">
										<img border="0" src="seta.gif" width="6" height="6">
										<span style="text-transform: uppercase">
										<b><font face="Verdana">CATEGORIAS</font></b></span></td>
									</tr>
									<tr>
										<td width="579">
										&nbsp;</td>
									</tr>
									<tr>
										<td width="579">
										<table border="0" style="border-collapse: collapse" width="100%">
											<tr>
												<td><font color="#868686"><%

			
							Response.Write "<b>" & Recordset1.RecordCount & "</b> Fotos de laudo - Mostrando p&aacute;gina <b>" &_
							pagina & "</b> de <b>" & Recordset1.PageCount & "</b>"
%>
</font>
</td>
												<td width="120">		<% strProcura =	"acao=procurar"

			
							


			If pagina > 1 Then
				Response.Write "<a href=""listar_fotoslaudo.asp?" & strProcura & "&pagina=1""><img src=""imagens/First.gif"" border=""0"" alt=""Primeira""></a>&nbsp;" & Chr(13)
				Response.Write "<a href=""listar_fotoslaudo.asp?" & strProcura & "&pagina=" & (pagina - 1) & """><img src=""imagens/Previous.gif"" border=""0"" alt=""Anterior""></a>&nbsp;" & Chr(13)
			Else
				Response.Write "<img src=""imagens/First_off.gif"" border=""0"" alt=""Primeira"">&nbsp;" & Chr(13)
				Response.Write "<img src=""imagens/Previous_off.gif"" border=""0"" alt=""Anterior"">&nbsp;" & Chr(13)
			End If

			If CInt(pagina) <> CInt(Recordset1.PageCount) Then 
				Response.Write "<a href=""listar_fotoslaudo.asp?" & strProcura & "&pagina=" & (pagina + 1) & """><img src=""imagens/Next.gif"" border=""0"" alt=""Pr&oacute;xima""></a>&nbsp;" & Chr(13)
				Response.Write "<a href=""listar_fotoslaudo.asp?" & strProcura & "&pagina=" & Recordset1.PageCount & """><img src=""imagens/Last.gif"" border=""0"" alt=""&Uacute;ltima""></a>&nbsp;" & Chr(13)
			Else
				Response.Write "<img src=""imagens/Next_off.gif"" border=""0"" alt=""Pr&oacute;xima"">&nbsp;" & Chr(13)
				Response.Write "<img src=""imagens/Last_off.gif"" border=""0"" alt=""&Uacute;ltima"">&nbsp;" & Chr(13)
			End If %></td>
											</tr>
										</table>
										</td>
									</tr>
									<tr>
										<td width="579">&nbsp;</td>
									</tr>
									<tr>
										<td background="bg_pontos.gif" width="579">
										<img border="0" src="inclui1.gif"></td>
									</tr>
									<tr>
										<td width="579">
										<table border="0" style="border-collapse: collapse" width="100%">
											<tr>
												<td>&nbsp;</td>
												<td width="65" align="center">
												<font color="#666666">Editar</font></td>
												<td width="82" align="center">
												<font color="#666666">Excluir</font></td>
											</tr>
										</table>
										</td>
									</tr>


																				<% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%><tr>
										<td width="579">
										<table border="0" style="border-collapse: collapse" width="579" cellpadding="6">
<tr>
<td width="367"><b><font color="#868686"></font><%=(Recordset1.Fields.Item("nome").Value)%></b></td>
												<td width="47" align="center">
												<a href="editar_fotoslaudo.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>">
												<img border="0" src="editar.gif"></a></td>
												<td width="61" align="center">
												<a onclick="javascript:if (!confirm('Você tem certeza que deseja excluir?')){return false;}" href="del_fotoslaudo.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>">
												<img height="16" alt="del" src="excluir.gif" width="16" border="0"></a></td>
											</tr>
										</table>
										</td>
									</tr>
									<tr>
										<td background="bg_pontos.gif" width="579">
										<img border="0" src="inclui1.gif"></td>
									</tr>                                      
									<tr>
										<td width="579">
										&nbsp;</td>
									</tr>
									<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
									<tr>
										<td width="579">&nbsp;</td>
									</tr>
									<tr>
										<td width="579">
										<table border="0" style="border-collapse: collapse" width="100%">
											<tr>
												<td><font color="#868686"><%

			
							Response.Write "<b>" & Recordset1.RecordCount & "</b> Fotos de laudo - Mostrando p&aacute;gina <b>" &_
							pagina & "</b> de <b>" & Recordset1.PageCount & "</b>"
%></font>
</td>
												<td width="120">		<% strProcura =	"acao=procurar"

			
							


			If pagina > 1 Then
				Response.Write "<a href=""listar_fotoslaudo.asp?" & strProcura & "&pagina=1""><img src=""imagens/First.gif"" border=""0"" alt=""Primeira""></a>&nbsp;" & Chr(13)
				Response.Write "<a href=""listar_fotoslaudo.asp?" & strProcura & "&pagina=" & (pagina - 1) & """><img src=""imagens/Previous.gif"" border=""0"" alt=""Anterior""></a>&nbsp;" & Chr(13)
			Else
				Response.Write "<img src=""imagens/First_off.gif"" border=""0"" alt=""Primeira"">&nbsp;" & Chr(13)
				Response.Write "<img src=""imagens/Previous_off.gif"" border=""0"" alt=""Anterior"">&nbsp;" & Chr(13)
			End If

			If CInt(pagina) <> CInt(Recordset1.PageCount) Then 
				Response.Write "<a href=""listar_fotoslaudo.asp?" & strProcura & "&pagina=" & (pagina + 1) & """><img src=""imagens/Next.gif"" border=""0"" alt=""Pr&oacute;xima""></a>&nbsp;" & Chr(13)
				Response.Write "<a href=""listar_fotoslaudo.asp?" & strProcura & "&pagina=" & Recordset1.PageCount & """><img src=""imagens/Last.gif"" border=""0"" alt=""&Uacute;ltima""></a>&nbsp;" & Chr(13)
			Else
				Response.Write "<img src=""imagens/Next_off.gif"" border=""0"" alt=""Pr&oacute;xima"">&nbsp;" & Chr(13)
				Response.Write "<img src=""imagens/Last_off.gif"" border=""0"" alt=""&Uacute;ltima"">&nbsp;" & Chr(13)
			End If %></td>
											</tr>
										</table>
										</td>
									</tr>
									</table>
								<p>&nbsp;</td>
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
					<td><!--#include file="rodape.asp" --></td>
				</tr>
			</table>
			</td>
		</tr>
	</table>

    <input type="hidden" name="MM_insert" value="form1">
	</form>


</body>

</html>