<%@ Language=VBScript %>
<% Option Explicit %>
<%
Dim code, date, initials, author, title_pre1, title_pre2, title, title_post, process, photoset, settype, count, photoname, caption, sequence
Dim message, workcaption, workpos, rowcount, current_timestamp, millisecs, updateno, livecount, photopage

code = Request.Form("code")
date = left(code,10)
initials = Ucase(rtrim(mid(code,12)))
settype = Request.Form("type")
process = Request.Form("process")
photoset = Request.Form("photoset")
author = Request.Form("author")
title_pre1 = replace(Request.Form("title_pre1"),"'","''")	'convert to double apostrophe for SQL string
title_pre2 = replace(Request.Form("title_pre2"),"'","''")	'convert to double apostrophe for SQL string
title = replace(Request.Form("title"),"'","''")				'convert to double apostrophe for SQL string
title_post = replace(Request.Form("title_post"),"'","''")	'convert to double apostrophe for SQL string


message = ""

Dim conn, sql, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_update.inc"--><%
	   
'First delete all captions for this set, ready for rebuild

	sql = "delete from photo_event " 
	sql = sql & "where date = '" & date & "' "
	sql = sql & "  and photo_set = '" & photoset & "' "
	sql = sql & "  and type = '" & settype & "' "
	sql = sql & "  and comment_seq = 0 "	'0=caption	

	on error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0	
		
	Dim RegEx
	Set RegEx = New RegExp
	RegEx.Pattern = "\s"
	RegEx.Global = True
	
'Loop through the captions in form

livecount = 0

For count = 1 to Request.Form("caption").Count

	photoname = "Request.Form(""filename"")(" & count & ")"
	photoname = eval(photoname)

	caption = "Request.Form(""caption"")(" & count & ")"
	caption = eval(caption)
	caption = trim(caption)
	caption = replace(caption,"'","''")	'convert to double apostrophe for SQL string

	
	caption = RegEx.Replace(caption," ")	'blank out any unprintable characters
	caption = trim(caption)					'remove leading or trailing blanks	
	
	'if right(caption,1) = "." then 	'eliminate final full-stop, but only if it's the only one in the caption
		'workcaption = left(caption,len(caption)-1)
		'workpos = instr(workcaption,".")
		'if workpos = 0 then caption = left(caption,len(caption)-1)	'eliminate final full stop
	'end if

	sequence = "Request.Form(""sequence"")(" & count & ")"
	sequence = eval(sequence)
	
	response.write("|" & sequence & "|")
	
	if sequence > 0 then livecount = livecount + 1

	message = message & " " & count & " - " & caption & " - " & sequence & " - " & photoname & "<br>"
		
	sql = "insert into photo_event (date, initials, author, title_pre1, title_pre2, title, title_post, photo_set, type, photo_name, photo_seq, comment_seq, text) "
	sql = sql & "values ("
	sql = sql & "'" & date & "',"
	sql = sql & "'" & initials & "',"
	sql = sql & "'" & author & "',"
	sql = sql & "'" & title_pre1 & "',"
	sql = sql & "'" & title_pre2 & "',"
	sql = sql & "'" & title & "',"
	sql = sql & "'" & title_post & "',"
	sql = sql & photoset & ","
	sql = sql & "'" & settype & "',"
	sql = sql & "'" & photoname & "',"
	sql = sql & sequence & ","
	sql = sql & "0,"
	sql = sql & "'" & caption & "'"
	sql = sql & ")"	
			
	on error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0	
		 
Next

'Check if a event_control row already exists for this set: if so, update; if not, insert

sql = "select count(*) as count "
sql = sql & "from event_control " 
sql = sql & "where event_date = '" & date & "' "
sql = sql & "  and event_type = '" & settype & "' "
sql = sql & "  and material_type = 'I' "
sql = sql & "  and material_seq = " & photoset 

rs.open sql,conn,1,2
rowcount = rs.Fields("count")
rs.close

if rowcount > 0 then

	sql = "update event_control set "
	sql = sql & "whatsnew_heading = '" & title & "', " 
	if process > 1 then sql = sql & "publish_timestamp = DEFAULT, " 
	if author <> "" then 
		if left(author,21) = "Plymouth Argyle Media" then
			sql = sql & "material_details1 = '" & livecount & " photos thanks to Plymouth Argyle Media' "
		 else
		 	sql = sql & "material_details1 = '" & livecount & " photos from " & left(author,instr(author," ")-1)  & "' "
		end if
	  else
		sql = sql & "material_details1 = NULL "
	end if
		
	sql = sql & "where event_date = '" & date & "' "
	sql = sql & "  and event_type = '" & settype & "' "
	sql = sql & "  and material_type = 'I' "
	sql = sql & "  and material_seq = " & photoset 
	
	on error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
  else

	'Get highest update number for this date

	sql = "select isnull(max(updateno),0) as updateno "
	sql = sql & "from event_control "
	sql = sql & "where event_date = '" & date & "' "
	sql = sql & "  and event_type = '" & settype & "' "			
	rs.open sql,conn,1,2
	updateno = rs.Fields("updateno") + 1
	rs.close


	sql = "insert into event_control (event_date, event_published, event_type, material_type, material_seq, publish_timestamp, publish_by, updateno, whatsnew_heading, material_details1)"
	sql = sql & "values ("
	sql = sql & "'" & date & "',"
	sql = sql & "'N',"
	sql = sql & "'" & settype & "',"
	sql = sql & "'I',"
	sql = sql & photoset & ","
	sql = sql & "DEFAULT,"
	sql = sql & "'" & initials & "',"
	sql = sql & updateno & ","
	sql = sql & "'" & title & "',"
	if author <> "" then 
		if author = "Plymouth Argyle Media" then
			sql = sql & "'" & livecount & " photos thanks to Plymouth Argyle Media'"
		 else
		 	sql = sql & "'" & livecount & " photos from " & left(author,instr(author," ")-1)  & "'"
		end if
	  else
		sql = sql & "NULL"
	end if
	sql = sql & ")"	
			
	on error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0

end if

sql = "select datepart(ms,publish_timestamp) as millisecs "
sql = sql & "from event_control " 
sql = sql & "where event_date = '" & date & "' "
sql = sql & "  and event_type = '" & settype & "' "
sql = sql & "  and material_type = 'I' "
sql = sql & "  and material_seq = " & photoset 

rs.open sql,conn,1,2
millisecs = rs.Fields("millisecs")
rs.close

Dim strTo,strFrom,strCc,strBcc,subject
	   								
strTo = "captions_updated@greensonscreen.co.uk"
strFrom = "captions_updated@greensonscreen.co.uk"
strCc = ""
subject = "GoS photo captions update - " & process & " - " & settype & " - " & photoset & " - " & author& " - " & millisecs	   				
	   				
%><!--#include file="emailcode.asp"--><%
%>

<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">

<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
</head>
<body><!--#include file="top_code.htm"-->

<p class="style1bold" style="margin:36 auto">Thanks, the captions have been added/amended.</p>
<form action="photocaptions1.asp" method="post" name="form1">
<%
response.write("<input type=""hidden"" name=""code"" value=""" & code & """>")
response.write("<input type=""hidden"" name=""type"" value=""" & settype & """>")
response.write("<input type=""hidden"" name=""process"" value=""1"">")		'Whatever the original process, it needs to be a '1' to continue adding captions
%>
<input type="submit" name="b1" value="Add more captions for this set" style="font-size: 12px;">
</form>
<%
photopage = "photodisplay"
if settype = "H" then photopage = "photos"
if settype = "W" then photopage = "photoslideshow"
%>
<form action="<%response.write("photos.asp?parm=" & date & settype & photoset)%>&phase=review" method="post" name="form2">
<%
response.write("<input type=""hidden"" name=""code"" value=""" & code & """>")
response.write("<input type=""hidden"" name=""type"" value=""" & settype & """>")
response.write("<input type=""hidden"" name=""process"" value=""1"">")	
%>
<input type="submit" name="b2" value="Preview your photos" style="font-size: 12px;">
</form>
<%
if settype = "M" then
	response.write("<form action=""matchcontent.asp"" method=""post"" name=""form3"">")
	response.write("<input type=""hidden"" name=""code"" value=""" & code & ":" & millisecs & """>")
	response.write("<input type=""submit"" name=""b3"" value=""Update the Match Page and What's New"" style=""font-size: 12px;"">")
	response.write("</form>")
end if

if settype = "O" or settype = "F" or settype = "E" or settype = "H" then
	response.write("<form action=""photocaptions2.asp"" method=""post"" name=""form4"">")
	response.write("<input type=""hidden"" name=""code"" value=""" & code & ":" & millisecs & """>")
	response.write("<input type=""submit"" name=""b4"" value=""Update What's New"" style=""font-size: 12px;"">")
	response.write("</form>")
end if
%>



</body>
</html>