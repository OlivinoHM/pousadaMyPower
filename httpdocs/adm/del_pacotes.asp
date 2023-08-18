<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/dlfelix.asp" -->

<%

if(Request.QueryString("id") <> "") then Command1__id1 = Request.QueryString("id")

%>
<%
Dim del__MMColParam
del__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  del__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim del
Dim del_numRows

Set del = Server.CreateObject("ADODB.Recordset")
del.ActiveConnection = MM_dlfelix_STRING
del.Source = "SELECT *  FROM pgpacotes  WHERE id = " + Replace(del__MMColParam, "'", "''") + ""
del.CursorType = 0
del.CursorLocation = 2
del.LockType = 1
del.Open()

del_numRows = 0
%>
<%

set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = MM_dlfelix_STRING
Command1.CommandText = "DELETE FROM pgpacotes WHERE id = '" + Replace(Command1__id1, "'", "''") + "'"
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()
Response.redirect "listar_pacotes.asp"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>.</title>
</head>

<body>
</body>
</html>
<%
del.Close()
Set del = Nothing
%>