<%@ Language=VBScript %>
<% Option Explicit %>
<%
Dim linecount, i, message, yesno, current_status, timestamp, material_seq

linecount = Request.Form("linecount")

message = ""

Dim conn, sql, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_update.inc"--><%
	   
	
'Loop through the material in the form

For i = 1 to linecount

	yesno = Request.Form("pub" & i)
	current_status = Request.Form("current_status" & i)
	timestamp = Request.Form("ts" & i)
	material_seq = Request.Form("ms" & i)
	
	sql = "select event_published "
	sql = sql & "from event_control " 
	sql = sql & "where publish_timestamp = '" & timestamp & "' "
	sql = sql & "  and material_seq = " & material_seq 

	rs.open sql,conn,1,2
	
	current_status = rs.Fields("event_published")
	
	rs.close

	message = message & " " & i & " - from " & current_status & " to " & yesno & " - " & timestamp & " - " & material_seq & "<br>"

	if yesno <> current_status then		
		
		sql = "update event_control "
		sql = sql & "set event_published = '" & yesno & "' "
		sql = sql & "where publish_timestamp = '" & timestamp & "' "
		sql = sql & "  and material_seq = " & material_seq 
	
		on error resume next
		conn.Execute sql
		if err = 0 then 
		 	message = message & "Published status changed"
		  else 
			Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
		end if
		On error GoTo 0
		
	end if
		 
Next

Dim strTo,strFrom,strCc,strBcc,subject
	   								
strTo = "material_published@greensonscreen.co.uk"
strFrom = "material_published@greensonscreen.co.uk"
strCc = ""	   				
subject = "GoS event status"

%><!--#include file="emailcode.asp"--><%
%>

<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">

<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
</head>
<body><!--#include file="top_code.htm"-->

<p class="style1bold" style="margin:72 auto 24">Thanks, your wishes have been granted!</p>

<p class="style1bold" style="margin:24 auto 72"><a href=index.asp>Go to GoS Home</a></p>

</body>
</html>