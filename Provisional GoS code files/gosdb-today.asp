<%@ Language=VBScript %>
<% Option Explicit %>
<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">

<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
</head>
<body>

<%

' This page is run every date at 02:00 by a Plesk trigger, to provide a simple indication of the date of the database backup on Steve's PC 


Dim conn, sql
Set conn = Server.CreateObject("ADODB.Connection")

%><!--#include file="conn_update.inc"--><%

	sql = "update today set " 
	sql = sql & "date = GETDATE() "
	conn.Execute sql

%>

</body>
</html>
