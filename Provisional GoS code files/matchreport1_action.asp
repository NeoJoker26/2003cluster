<%@ Language=VBScript %>
<% Option Explicit %>

<%
Dim code1, code2, date, initials, output, headline, report, reportpart, reportparts, rowcount, millisecs, updateno, oldreportind, acknowledge 

code1 = Request.Form("code1")
code2 = Request.Form("code2")
date = left(code2,10)
initials = Ucase(rtrim(mid(code2,12)))
oldreportind = Request.Form("oldreportind")
acknowledge = Request.Form("acknowledge")

headline = replace(Request.Form("headline"),"'","''")
report = replace(Request.Form("report"),"'","''") 
report = replace(report,"‘","''")
report = replace(report,"’","''")
report = replace(report,"“","""")
report = replace(report,"”","""")
report = replace(report,"—","-")
report = replace(report,"–","-")
report = replace(report,"£","&pound;")

reportparts = split(report,Chr(13)&Chr(10))
report = ""
	
for each reportpart in reportparts
	if trim(reportpart) > "" then report = report & trim(reportpart) & "|p|"
next
		
if right(report,3) = "|p|" then report = left(report,len(report)-3)		'remove final paragraph marker

output = ""

Dim conn, sql, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_update.inc"--><%

'Check if the match_extra row yet exists for this set: if so, update; if not, insert a near-null row

sql = "select count(*) as count "
sql = sql & "from match_extra " 
sql = sql & "where date = '" & date & "' "

rs.open sql,conn,1,2
rowcount = rs.Fields("count")
rs.close

if rowcount > 0 then

	sql = "update match_extra set "
	sql = sql & "headline = '" & headline & "' "
	sql = sql & ", report = '" & report & "' "
	sql = sql & ", report_published = 'N' "
	if acknowledge > "" then sql = sql & ", report_acknowledge = '" & acknowledge & "' "
	sql = sql & "where date = '" & date & "' "
	
	on error resume next
	conn.Execute sql
	if err <> 0 then output = output & "<p>SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
	On Error GoTo 0	
 
 else
 
	sql = "insert into match_extra (date, headline, report, report_published, report_acknowledge) "
	sql = sql & "values ("
	sql = sql & "'" & date & "',"
	sql = sql & "'" & headline & "',"
	sql = sql & "'" & report & "',"
	sql = sql & "'N',"
	if acknowledge > "" then 
		sql = sql & "'" & acknowledge & "' "
	  else
	  	sql = sql & "NULL"
	end if
	sql = sql & ")"	
			
	on error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0

end if


'Get highest update number for this date

sql = "select isnull(max(updateno),0) as updateno "
sql = sql & "from event_control "
sql = sql & "where event_date = '" & date & "' "			
rs.open sql,conn,1,2
updateno = rs.Fields("updateno") + 1
rs.close

		
'Check if a event_control row already exists for this report: if not, insert

sql = "select count(*) as count "
sql = sql & "from event_control " 
sql = sql & "where event_date = '" & date & "' "
sql = sql & "  and material_type = 'S' "

rs.open sql,conn,1,2
rowcount = rs.Fields("count")
rs.close


if rowcount = 0 then

	sql = "insert into event_control (event_date, event_published, event_type, material_type, material_seq, publish_timestamp, publish_by, updateno, material_details1, material_details2)"
	sql = sql & "values ("
	sql = sql & "'" & date & "',"
	sql = sql & "'N',"
	sql = sql & "'M',"
	sql = sql & "'S',"
	sql = sql & "0,"
	sql = sql & "DEFAULT,"
	sql = sql & "'" & initials & "',"
	sql = sql & updateno & ","
	sql = sql & "'Match Report',"
	sql = sql & "NULL"
	sql = sql & ")"	
		
	on error resume next
	conn.Execute sql
	if err <> 0 then output = output & "<p>SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
	On Error GoTo 0
	
  else

	'set the published flag off in event_control to match that in match_extra (for the match summary/report)
	sql = "update event_control set "
	sql = sql & "event_published = 'N', "
	sql = sql & "publish_by = '" & initials & "' "
	sql = sql & "where event_date = '" & date & "' "
	sql = sql & "  and material_type = 'S' "
	
	on error resume next
	conn.Execute sql
	if err <> 0 then output = output & "<p>SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
	On Error GoTo 0	

end if

sql = "select datepart(ms,publish_timestamp) as millisecs "
sql = sql & "from event_control " 
sql = sql & "where event_date = '" & date & "' "
sql = sql & "  and material_type = 'S' "

rs.open sql,conn,1,2
millisecs = rs.Fields("millisecs") 		
rs.close

Dim strTo,strFrom,strCc,strBcc,message,subject
	   								
strTo = "steve@greensonscreen.co.uk"
strFrom = "match_report_text@greensonscreen.co.uk" 
strCc = ""					
subject = "GoS match report update"
report = replace(report,"|p|","<p>") 	'ensure new paragraphs for the email
report = replace(report,"''","'") 		'revert report to single apostrophes for the email
report = replace(report,"&pound;","£")	'revert report to £ sign for the email
message = "<p>Acknowledge: " & acknowledge & " Match: " & date & " By: " & initials & "</p>"
message = message & "<p>" & headline & "</p><p>" & report & "</p><p>" & output & "</p>" 
message = message & "<p>Report length: " & len(report) & "</p>" 
message = message & "<p><a href=""http://www.greensonscreen.co.uk/matchreport1.asp?date=" & date & "&oldreportind=y&acknowledge=" & acknowledge & "&code=" & code1 & """>Amend Report</a></p>" 
message = message & "<p><a href=""http://www.greensonscreen.co.uk/gosdb-match.asp?date=" & date & """>Match Page</a></p>"  

%><!--#include file="emailcode.asp"--><%
%>

<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">

<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
</head>
<body><!--#include file="top_code.htm"-->

<%	
if output > "" then response.write(output)	
%>
<p class="style1bold" style="margin:36 auto">Thanks, the report has been added.</p>
<form action="gosdb-match.asp?date=<%response.write(date)%>&phase=review" method="post" name="form1">
<input type="submit" name="b1" value="Preview the Match Page" style="font-size: 12px;">
</form>
<form action="matchreport1.asp" method="post" name="form2">
<%
response.write("<input type=""hidden"" name=""code"" value=""" & code1 & """>")
response.write("<input type=""hidden"" name=""matchdate"" value=""" & date & """>")
response.write("<input type=""hidden"" name=""oldreportind"" value=""" & oldreportind & """>")
response.write("<input type=""hidden"" name=""acknowledge"" value=""" & acknowledge & """>")
%>
<input type="submit" name="b2" value="Amend the report" style="font-size: 12px; padding:2px 5px;">
</form>
<form action="matchcontent.asp" method="post" name="form3">
<%
response.write("<input type=""hidden"" name=""code"" value=""" & code2 & ":" & millisecs & """>")
response.write("<input type=""hidden"" name=""oldreportind"" value=""" & oldreportind & """>")
response.write("<input type=""submit"" name=""b3"" style=""font-size: 12px; padding:2px 5px;"" value=""Update the Match Page")
if oldreportind = "" then response.write(" and What's New")
response.write(""">")
%>
</form>

</body>
</html>