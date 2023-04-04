<%@ Language=VBScript %>
<% Option Explicit %>

<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">

<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
<style>
<!--
p {font-size: 11px; text-align:left;}
-->
</style>
</head>
<body><!--#include file="top_code.htm"-->

<%
Dim i, initials, matchcount, output, dates

initials = Request.Form("initials")
matchcount = Request.Form("matchcount")

output = ""

Dim conn, sql, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_update.inc"--><%

for i = 1 to matchcount

	if Request.Form("was" & i) <> Request.Form("willbe" & i) then 

		sql = "update season_this set "
		sql = sql & "reporter = '" & Request.Form("willbe" & i)  & "' "	
		sql = sql & "where date = '" & Request.Form("date" & i) & "' "
	
		on error resume next
		conn.Execute sql
		if err <> 0 then 
			output = output & "<p>SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
		  else
		  	dates = dates & "'" & Request.Form("date" & i) & "',"
		 end if
		On Error GoTo 0
		
	end if

next

if dates > "" then
	
		dates = left(dates,len(dates)-1)	'drop last comma

		sql = "select date, opposition, homeaway, reporter "
		sql = sql & "from season_this "
		sql = sql & "where date in (" & dates & ") " 
		sql = sql & "order by date "

		rs.open sql,conn,1,2
					
			Do While Not rs.EOF
			
				output = output & "<p>" & rs.Fields("date") & " " & rs.Fields("opposition") & " (" & rs.Fields("homeaway") & ") "
				if rs.Fields("reporter") = "? " then
					output = output & "is now unallocated<br>"
				  else
				  	output = output & "now with " & rs.Fields("reporter") & "</p>"
				end if
				
				rs.MoveNext
			Loop
	
		rs.close
		
	else
	
		output = output & "<p>No changes detected</p>"

end if
		
response.write("<div style=""width:400; margin:72px auto"">" & output & "</div>")	

if dates > "" then

		Dim strTo,strFrom,strCc,strBcc,message,subject
	   								
		strTo = "steve@greensonscreen.co.uk; malcolmtownrow@blueyonder.co.uk; mathewlawrie@yahoo.com;"
		strFrom = "match_report_schedule@greensonscreen.co.uk"
		strCc = "" 
		subject = "GoS match report schedule update"
		message = "The following changes have been made to the schedule:<br>" & output	
		
		%><!--#include file="emailcode.asp"--><%
	
end if
%>

<!--#include file="base_code.htm"-->
</body>

</html>