<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<script Language="JavaScript">
function ValidaPagina(form)
{
	if(CampoBranco(form.blob)==true)
	{
		alert("Escolha um Arquivo a ser enviado.");
		form.blob.focus();
		return false;
	}
}
function CampoBranco(campo)
{
	if (campo.value == "")
		return true;
	else
		return false;
}
-->
</script>
      
<script language="VBScript">
SUB enviar_ONCLICK()
  billform.enviar.Value = "Aguarde..."
END SUB
</script>
<HTML>
<HEAD>
	<title>Upload</title>
</HEAD>
<body bgcolor="#DBDBDB" vlink="#0000FF" alink="#0000FF" text="#0000FF" link="#0000FF" topmargin="0" leftmargin="0">
<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="outputFile1.asp?img=<%=request.querystring("img")%>" onSubmit="return ValidaPagina(this);" name="billform">
  <div align="center">
  <table width="417" border="0" height="127" cellspacing="0" cellpadding="4">
    <tr bgcolor="#E5E5E5"> 
      <td colspan="2" height="31" bgcolor="#333333"> 
        <div align="center"><b><font face="Arial" color="#FFFFFF">Inserir Foto </font><font face="Arial" color="#FFFF00"></font></b></div>
      </td>
    </tr>
    <tr bgcolor="#A6D2FF"> 
      <td width="77" height="41" bgcolor="#999999"> 
        <div align="right">
			<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#000000">
			<b>Imagem: 
          </b> </font></div>
      </td>
      <td width="364" height="41" bgcolor="#999999"> 
        <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
        <input type="file" name="blob" size="30" style="border: 1px solid #006666; ">
 
      	</font>
 
      </td>
    </tr>
    <tr bgcolor="#A6D2FF"> 
      <td colspan="2" height="55" bgcolor="#999999"> 
        <p align="center" style="margin-top: 0; margin-bottom: 0"> 
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          <input type="submit" value="Continuar &gt;&gt;" name="enviar" style="color: #FFFFFF; border: 1px solid #808080; background-color: #666666" >
      	</font>
		<p align="center" style="margin-top: 0; margin-bottom: 0">&nbsp; 
          </td>
    </tr>
    </table>
	</div>
</FORM>

<br>
  </BODY>
</HTML>