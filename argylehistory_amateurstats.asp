
<%@ Language=VBScript %>
<% Option Explicit %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
   <meta name="Author" content="Greens on Screen">
   <meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<title>Greens on Screen: Complete History of Plymouth Argyle</title>

<link rel="stylesheet" type="text/css" href="gos2.css">
<style>
<!--
.head1{margin: 6 4 3 4; font-family: verdana, arial; line-height: 1.2; font-size: 18px; font-weight:bold; text-align: center; color: #457B44; }
-->
   </style>
</head>
<body><!--#include file="top_code.htm"-->
<center>

<% 
Dim stat
stat = Request.QueryString("stat")
%>
    <p class="head1">The History of Argyle</p>
    <p style="margin-right:0; margin-top:3; margin-bottom:6; align:center; font-size: 11px;" >
    <span xmlns:dc="http://purl.org/dc/elements/1.1/" href="http://purl.org/dc/dcmitype/Text" property="dc:title" rel="dc:type">
    This <span xmlns:dc="http://purl.org/dc/elements/1.1/" href="http://purl.org/dc/dcmitype/Text" rel="dc:type">work</span> is 
    licensed under a
<b> 
<u> <a rel="license" href="http://creativecommons.org/licenses/by-nc-nd/3.0/">
Creative Commons Licence</a></u></b>. </span></p>
    <p style="margin-right:0; margin-top:3; margin-bottom:12; text-align:center; font-size: 11px;" >
    Statistics researched, produced and donated to Greens on 
    Screen by Roger Walters.</p> 
	
    <img border="0" src="images/history/<%=stat%>.png">

<br>
<!--#include file="base_code.htm"-->
</body>
</html>