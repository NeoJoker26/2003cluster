<%@ Language=VBScript %>
<% Option Explicit %>
<%
Dim code, date, initials, filecount, filename, caption, currentdatetime, millisecs
Dim message, updateno, rowcount, i

code = Request.Form("code")
date = left(code,10)
initials = Ucase(rtrim(mid(code,12)))
filecount = Request.Form("filecount")

message = ""

Dim conn, sql, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_update.inc"--><%

'Get highest update number for this date

sql = "select isnull(max(updateno),0) as updateno "
sql = sql & "from event_control "
sql = sql & "where event_date = '" & date & "' "			
rs.open sql,conn,1,2
updateno = rs.Fields("updateno") + 1
rs.close

'prepare a common date/time for consistency if inserts are being performed

sql = "select convert(varchar,getdate(),120) as currentdatetime "
rs.open sql,conn,1,2
currentdatetime = rs.Fields("currentdatetime")
rs.close	   
	
'Loop through the titles in form

For i = 1 to filecount

	filename = "Request.Form(""filename"" & i)"
	filename = eval(filename)

	caption = "Request.Form(""caption"" & i)"
	caption = eval(caption)
	caption = trim(caption)
	caption = replace(caption,"'","''")	'convert to double apostrophe for SQL string

	message = message & " " & i & " - " & filename & " - " & caption & "<br>"
		 
	'Check if a event_control row already exists for this set: if so, update; if not, insert

	sql = "select count(*) as count "
	sql = sql & "from event_control " 
	sql = sql & "where event_date = '" & date & "' "
	sql = sql & "  and material_type = 'A' "
	sql = sql & "  and material_details1 = '" & filename & "' " 

	rs.open sql,conn,1,2
	rowcount = rs.Fields("count")
	rs.close

	if rowcount > 0 then

		sql = "update event_control set "
		sql = sql & "material_details2 = '" & caption  & "' "	
		sql = sql & "where event_date = '" & date & "' "
		sql = sql & "  and material_type = 'A' "
		sql = sql & "  and material_details1 = '" & filename & "' " 
	
		on error resume next
		conn.Execute sql
		if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
		On Error GoTo 0
	
  	else

		sql = "set dateformat ymd; "
		sql = sql & "insert into event_control (event_date, event_published, event_type, material_type, material_seq, publish_timestamp, publish_by, updateno, material_details1, material_details2)"
		sql = sql & "values ("
		sql = sql & "'" & date & "',"
		sql = sql & "'N',"
		sql = sql & "'M',"
		sql = sql & "'A',"
		sql = sql & i & ","
		sql = sql & "'" & currentdatetime & "',"
		sql = sql & "'" & initials & "',"
		sql = sql & updateno & ","
		sql = sql & "'" & filename  & "',"
		sql = sql & "'" & caption  & "'"
		sql = sql & ")"	
			
		on error resume next
		conn.Execute sql
		if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
		On Error GoTo 0

	end if
	
Next

sql = "select datepart(ms,publish_timestamp) as millisecs "
sql = sql & "from event_control " 
sql = sql & "where event_date = '" & date & "' "
sql = sql & "  and material_type = 'A' "
sql = sql & "  and material_details1 = '" & filename & "' "

rs.open sql,conn,1,2
millisecs = rs.Fields("millisecs") 		'The millisec value for tha last clip (only one valid one required later)
rs.close

Dim strTo,strFrom,strCc,strBcc,subject
	   								
strTo = "audio_added@greensonscreen.co.uk"
strFrom = "audio_added@greensonscreen.co.uk"
strCc = ""
subject = "GoS audio update - " & initials & " - " & millisecs
	   			
%><!--#include virtual="/emailcode.asp"--><%

%>

<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">

<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
</head>
<body><!--#include file="top_code.htm"-->

<p class="style1bold" style="margin:36 auto">Thanks, the titles have been added.</p>
<form action="gosdb-match.asp?date=<%response.write(date)%>&phase=review" method="post" name="form2">
<input type="submit" name="b2" value="Preview the Match Page" style="font-size: 12px;">
</form>
<form action="matchcontent.asp" method="post" name="form3">
<%
response.write("<input type=""hidden"" name=""code"" value=""" & code & ":" & millisecs & """>")
%>
<input type="submit" name="b3" value="Update the Match Page and What's New" style="font-size: 12px;">
</form>

</body>
</html>