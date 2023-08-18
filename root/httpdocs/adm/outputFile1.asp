<%@LANGUAGE="VBSCRIPT"%>

<%	Server.ScriptTimeout = 900

	path = "C:\Inetpub\vhosts\pousadamypower.com.br\httpdocs\temp\"
	path_destino = "C:\Inetpub\vhosts\pousadamypower.com.br\httpdocs\images\"
	
	
		Set Upload = Server.CreateObject("Persits.Upload")
		
	Upload.OverwriteFiles = False 
	Count = Upload.Save(path & Upload.form("blob"))
	
	Set File = Upload.Files(1)
		If File.ImageType = "UNKNOWN" Then
			File.Delete
end if

Function rndit()
	Dim RndTemp 
	Randomize 
	RndTemp = Rnd * 10000 
	RndTemp= Int(RndTemp) 
	rndit = RndTemp 
End Function 
PasswordTemp = rndit()
	

	Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.GetFolder(path)
		Set ff = fso.GetFolder(path_destino)
	Set fc = f.Files
	For Each Arquivo in fc
		Set ImagemPQ = Server.CreateObject("Persits.Jpeg")
			ImagemPQ.Open path & Arquivo.Name
			largura = 800
			altura = 600
		if ImagemPQ.OriginalHeight < altura AND ImagemPQ.OriginalWidth < largura then
			limite1 = (cInt(ImagemPQ.OriginalWidth - largura)/3)
			limite2 = (cInt(ImagemPQ.OriginalHeight - altura)/3)
			ImagemPQ.Crop limite1, limite2, (largura + limite1), (altura + limite2)
		else
			temp1 = largura * ImagemPQ.OriginalHeight / ImagemPQ.OriginalWidth
			if temp1 > altura then
				 ImagemPQ.Width = largura
				 ImagemPQ.Height = ImagemPQ.Width * ImagemPQ.OriginalHeight / ImagemPQ.OriginalWidth
				 limite = (cInt(ImagemPQ.Height - altura)/2) - 1
				 ImagemPQ.Crop 0, limite, largura, (limite + altura) 
			else
				ImagemPQ.Height = altura
				ImagemPQ.Width = ImagemPQ.Height * ImagemPQ.OriginalWidth / ImagemPQ.OriginalHeight 
				limite = (cInt(ImagemPQ.Width - largura)/2) - 1
				ImagemPQ.Crop limite, 0, (limite + largura), altura
			end if
		end if 
	
		ImagemPQ.Save path_destino & "mini_" & Arquivo.Name
		Set ImagemPQ = Nothing

		Set Imagem = Server.CreateObject("Persits.Jpeg")
		Imagem.Open path & Arquivo.Name
		Imagem.Save path_destino &  Arquivo.Name
		Set Imagem = Nothing
		
		Set fso = CreateObject("Scripting.FileSystemObject")
	    Set FileObject = fso.GetFile(Server.MapPath("/temp/")&"\" & Arquivo.Name & "\")
		FileObject.Delete
	 
	Set fileObject = Nothing 
	
	Next
	
	Set fso = nothing
	Set ff = nothing
	Set f = nothing
	Set fc = nothing

	%>



<title>Upload</title>
<style type="text/css">
<!--
.style1 {color: #FF0000}
.style2 {color: #FFCC00}
-->
</style>
<body bgcolor="#DBDBDB" link="#0000FF" vlink="#0000FF" alink="#0000FF" topmargin="0" leftmargin="0">
<form><span class="gen"><a name="top"></a></span>








  <% if Request.QueryString("img") = "des01" then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img.value = selected_username;
	opener.document.img.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% end if %>

  <% if Request.QueryString("img") = "img1" then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img1.value = selected_username;
	opener.document.img1.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% end if %>


  <% if Request.QueryString("img") = "img2" then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img2.value = selected_username;
	opener.document.img2.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% end if %>
  <% if request.querystring("pag") = "des02" then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].nove.value = selected_username;
	opener.document.nove.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% end if %>


 <% if request.querystring("pag") = "des03" then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].treze.value = selected_username;
	opener.document.treze.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% end if %>

 <% if request.querystring("pag") = "des04" then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].dezessete.value = selected_username;
	opener.document.dezessete.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% end if %>


<div align="center">
<table width="366" border="0" bgcolor="#CCCCCC" height="141" cellpadding="3" cellspacing="0">
  <tr>
    <td height="21" bgcolor="#000000" width="360" colspan="2"> 
      <div align="center" class="style1"><font face="Verdana" size="1">&nbsp;</font><b><font face="Arial" size="2">Upload 
        concluído</font></b></div></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="21" bgcolor="#666666" width="360" colspan="2"> 
      <p align="center"><font face="Arial" size="2"><b><span class="style2">Atenção: 
          Abaixo clique em </span><font color="#33CCCC">Confirmar</font></b></font></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="97" bgcolor="#808080" width="122"> 
      <p align="center">
		<img src="/images/mini_<% = File.ExtractFileName %>" width="150" height="100"></td>
    <td height="97" bgcolor="#808080" width="232"> 
      <div align="center">
		<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="5">
			<tr>
				<td bgcolor="#000000">
				<table border="0" width="100%" id="table4">
					<tr>
						<td>
						<p align="center"><span class="genmed">
						<input type="submit" class="liteoption" onClick="refresh_username(this.form.username_list.options[this.form.username_list.selectedIndex].value);return false;" name="use" value="Confirmar" style="color: #FFFFFF; font-family: Arial; font-size: 8pt; border: 1px solid #006699; background-color: #003366" />
</span></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		<p></div>
   
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><span class="genmed">
				<select disabled name="username_list" size="1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #000000">
				<option value="/images/mini_<% = File.ExtractFileName %>">
				mini_<% = File.ExtractFileName %></option></select></span></form>
</div>
