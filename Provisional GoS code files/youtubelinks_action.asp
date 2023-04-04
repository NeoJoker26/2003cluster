<%@ Language=VBScript %>
<% Option Explicit %>
<%
Dim eventdate, content, datetime, youtubecode, delete


eventdate = Request.Form("eventdate")
content = Request.Form("content")
datetime = Request.Form("datetime")
youtubecode = Request.Form("youtubecode")
delete = Request.Form("delete")

response.write("A" & eventdate & "B" & datetime)

Dim conn, sql, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_update.inc"--><%

if delete = "Y" then

	'existing rows exist for this date and content type, so delete them first (the default action)

	sql = "delete from event_control "
	sql = sql & "where event_date = '" & eventdate & "' "
	sql = sql & "  and material_details2 = '" & content & "' "			
	on error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0	

end if
	  
'Now insert new row 
	
sql = "set dateformat ymd; "
sql = sql & "insert into event_control (event_date, event_published, event_type, material_type, material_seq, publish_timestamp, updateno, material_details1, material_details2)"
sql = sql & "values ("
sql = sql & "'" & eventdate & "',"
sql = sql & "'Y',"
sql = sql & "'M',"
sql = sql & "'Y',"
sql = sql & "1,"
sql = sql & "'" & datetime & "',"
sql = sql & "99,"
sql = sql & "'" & youtubecode & "',"
sql = sql & "'" & content & "'"
sql = sql & ")"	
			
on error resume next
conn.Execute sql
if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
On Error GoTo 0

Dim strTo,strFrom,strCc,strBcc,subject,message
	   								
strTo = "youtube_added@greensonscreen.co.uk"
strFrom = "youtube_added@greensonscreen.co.uk"
strCc = ""
subject = "GoS YouTube link added"
message=""
	   			
%>
<!--#include virtual="/emailcode.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">

<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
</head>
<body><!--#include file="top_code.htm"-->

<p class="style1" style="margin:36 auto">The YouTube link has been added.</p>

</body>
</html>