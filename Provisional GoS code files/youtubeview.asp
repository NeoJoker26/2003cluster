<%@ Language=VBScript %> 
<% Option Explicit %> 

<!doctype html>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=ISO-8859-19" />
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GoS-DB Match Page</title>
<link rel="stylesheet" type="text/css" href="gos2.css">
<link rel="stylesheet" type="text/css" href="highslide/highslide.css" />

<style>
<!--

.video {border-width: 0; vertical-align:bottom; margin-right: 6px; width: 12px;} 
table {border-collapse: collapse;}
th, td {border: 1px solid #c0c0c0; padding:2px 10px;}
.right {text-align: right}

-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->

<div id="container">

<%

Dim conn,sql,rs
Dim parm,output,subhead,sqlqual1,sqlqual2,sqlorder

parm = Request.QueryString("parm")

if parm = 1 then
	subhead = "<b>Last Ten Days</b> | <a href=""youtubeview.asp?parm=0"">All Identified Videos</a>"
	sqlqual1 = 	"where event_published = 'Y' and event_type = 'V' "
	sqlqual2 = "and datediff(""dd"", publish_timestamp, getdate()) < 10 "
	sqlorder = "order by cast(publish_timestamp as date) desc, date "
  else
	subhead = "<b>All Identified Videos</b> | <a href=""youtubeview.asp?parm=1"">Last Ten Days</a>"
	sqlqual1 = 	"where event_published = 'Y' and ((event_type = 'M' and material_type = 'Y') or event_type = 'V') "
	sqlqual2 = ""
  	sqlorder = "order by date "
end if	
 

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_read.inc"-->

<p style="margin: 18px auto 6px; font-size: 18px; color:#457b44; font-family:Arial,Helvetica,sans-serif;">VIDEOS FOUND ONLINE</p>

<%
	output = "<p class=""style4"" style=""margin: 6px auto;"">" & subhead & "</p>"

	sql = "select date, publish_timestamp, opposition, goalsfor, goalsagainst, homeaway, material_details1, straight_to_youtube "
	sql = sql & "from match join event_control on date = event_date "
	sql = sql & sqlqual1 & sqlqual2 & sqlorder
	
	rs.open sql,conn,1,2
	
	if rs.RecordCount > 0 then 

       	output = output & "<table style=""border-collapse: collapse; border: 1px solid #d0d0d0; margin-top: 12px; padding: 2px 1px;"">"
					
		Do While Not rs.EOF
		
			output = output & "<tr>"
			if parm = 1 then output = output & "<td class=""right"">" & DateDiff("d", Now(), rs.Fields("publish_timestamp")) & "</td>"
			output = output & "<td class=""right"">" & right("0" & day(rs.Fields("date")),2) & " " & monthname(month(rs.Fields("date")),True) & " " & year(rs.Fields("date")) & "</td>"
			
			output = output & "<td>"
			if rs.Fields("homeaway") = "H" then
				output = output & "Argyle " & rs.Fields("goalsfor") & "-" & rs.Fields("goalsagainst") & " " & rs.Fields("opposition") & "</td>"
		 	 else 
				output = output & rs.Fields("opposition") & " " & rs.Fields("goalsagainst") & "-" & rs.Fields("goalsfor") & " Argyle" & "</td>"   		  
			end if	  
			
			output = output & "<td><a href=""gosdb-match.asp?date=" & rs.Fields("date") & """>Match Page</a></td>"
			output = output & "<td style=""white-space: nowrap;"">"
			if rs.Fields("straight_to_youtube") = "Y" then
				output = output & " "
			  else
				output = output & "<a href=""https://www.youtube.com/embed/" & rs.Fields("material_details1") 
				output = output & "?rel=0&wmode=transparent&fs=0&autoplay=1"" onclick=""return hs.htmlExpand(this, {objectType: 'iframe'})"" class=""highslide"">"
				output = output & "Play Here</a>"
			end if
			output = output & "</td>"
			output = output & "<td><a href=""https://www.youtube.com/watch?v=" & rs.Fields("material_details1") & """>Play on YouTube</a></td>"
		
		    output = output & "</tr>"
			rs.MoveNext
		
		Loop
		
       	output = output & "</table>"		
						     	
      	rs.close
      	
	  else
	
		output = "<p class=""style1"">No video clips found</p>"
	
	end if	
		
conn.close

response.write(output)

%>

</div>
<br>
<!--#include file="base_code.htm"-->
</body>
</html>