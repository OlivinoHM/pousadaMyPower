<%@LANGUAGE="VBSCRIPT"%>
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

Response.Expires=0
Response.Buffer = TRUE
Response.Clear
byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)
Dim UploadRequest
Set UploadRequest = CreateObject("Scripting.Dictionary")

BuildUploadRequest  RequestBin

contentType = UploadRequest.Item("blob").Item("ContentType") 
filepathname = UploadRequest.Item("blob").Item("FileName") 
filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\")) 
filename = Replace(filename,".","_"& rndit &".") 
value = UploadRequest.Item("blob").Item("Value")

 Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")

 pathEnd = Len(Server.mappath(Request.ServerVariables("PATH_INFO")))-14
  	    Set MyFile = ScriptObject.CreateTextFile("C:\Inetpub\vhosts\pousadamypower.com.br\httpdocs\images\"&filename)
   'Set MyFile = ScriptObject.CreateTextFile("d:\caminho_dapasta\"&filename)
   'Set MyFile = ScriptObject.CreateTextFile(Left(Server.mappath(Request.ServerVariables("PATH_INFO")),pathEnd)&"\"&filename)
 
 For i = 1 to LenB(value) 
	 MyFile.Write chr(AscB(MidB(value,i,1)))
 Next
 MyFile.Close
'renomeia ao arquivo
Function rndit()
	Dim RndTemp 
	Randomize 
	RndTemp = Rnd * 10000 
	RndTemp= Int(RndTemp) 
	rndit = RndTemp 
End Function 
PasswordTemp = rndit() 
%>

<title>Upload</title>
<body bgcolor="#DBDBDB" link="#0000FF" vlink="#0000FF" alink="#0000FF" topmargin="0" leftmargin="0">
<form>

<span class="gen"><a name="top"></a></span>







<% if (Request.QueryString("img")) = "1" Then %>
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
<% End If %>


<% if (Request.QueryString("img")) = "2" Then %>
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
<% End If %>



<% if (Request.QueryString("img")) = "3" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img3.value = selected_username;
	opener.document.img3.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>


<% if (Request.QueryString("img")) = "4" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img4.value = selected_username;
	opener.document.img4.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>


<% if (Request.QueryString("img")) = "5" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img5.value = selected_username;
	opener.document.img5.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>



<% if (Request.QueryString("img")) = "6" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img6.value = selected_username;
	opener.document.img6.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>



<% if (Request.QueryString("img")) = "7" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img7.value = selected_username;
	opener.document.img7.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>

<% if (Request.QueryString("img")) = "8" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img8.value = selected_username;
	opener.document.img8.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>

<% if (Request.QueryString("img")) = "9" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img9.value = selected_username;
	opener.document.img9.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>

<% if (Request.QueryString("img")) = "10" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].img10.value = selected_username;
	opener.document.img10.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>


<% if (Request.QueryString("imgpeq")) = "3" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['post'].imgpeq3.value = selected_username;
	opener.document.imgpeq3.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>


<% if (Request.QueryString("imgpeq")) = "4" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['post'].imgpeq4.value = selected_username;
	opener.document.imgpeq4.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>


<% if (Request.QueryString("imgpeq")) = "5" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['post'].imgpeq5.value = selected_username;
	opener.document.imgpeq5.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>



<% if (Request.QueryString("img")) = "11" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['post'].img11.value = selected_username;
	opener.document.img11.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>



<% if (Request.QueryString("img")) = "12" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['post'].img12.value = selected_username;
	opener.document.img12.src = selected_username; 
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>



<% if (Request.QueryString("img")) = "arquivo" Then %>
  <script language="javascript" type="text/javascript">
<!--
function refresh_username(selected_username)
{
	opener.document.forms['form1'].esboco.value = selected_username;
	opener.focus();
	window.close();
}
//-->
</script>
<% End If %>



<div align="center">
<table width="366" border="0" bgcolor="#CCCCCC" height="142" cellpadding="3" cellspacing="0">
  <tr>
    <td height="21" bgcolor="#000000" width="360" colspan="2"> 
      <font color="#CCCCCC"><font face="Verdana" size="1">&nbsp;</font><b><font face="Arial" size="2">Upload 
		concluído</font></b></font></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="21" bgcolor="#666666" width="360" colspan="2"> 
      <p align="center"><font face="Arial" size="2"><b><font color="#FFFFFF">Atenção: 
		Abaixo clique em </font><font color="#33CCCC">selecionar</font></b></font></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="98" bgcolor="#808080" width="119"> 
      <p align="center"> 
		<img src="/images/<%= filename%>" width="99" height="74"></td>
    <td height="98" bgcolor="#808080" width="235"> 
      <div align="center">
		<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="5">
			<tr>
				<td bgcolor="#000000">
				<table border="0" width="100%" id="table4">
					<tr>
						<td>
						<p align="center"><span class="genmed">
						<input type="submit" class="liteoption" onClick="refresh_username(this.form.username_list.options[this.form.username_list.selectedIndex].value);return false;" name="use" value="Selecionar" style="color: #FFFFFF; font-family: Arial; font-size: 8pt; border: 1px solid #006699; background-color: #003366" /></span></td>
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
				<option value="/images/<%= filename%>">
				<%= filename%></option></select></span></form>&nbsp;</div>
<!--#include file="uploadfla.asp"-->