<%@ Language=VBScript %> 
<% Option Explicit %>

<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Author" content="Trevor Scallan">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>GoS-DB Miscellaneous Report</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
.tables td {border: 1px solid #c0c0c0; text-align: left; margin: 0; white-space:nowrap; padding-left:5; padding-right:5; padding-top:1; padding-bottom:1}
.tables .center {text-align: center;} 
.tables .right {text-align: right;} 
.tables .bold {background-color: #e0f0e0} 
-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline, selected_comp(3), selected_ha(2), competition, homeaway, colpref, homeaway_val, startdate, enddate, work1, summerdate, warn, datestyle, heading 


Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%
%>
  <center>
  <table border="0" cellspacing="0" style="border-collapse: collapse" 
  cellpadding="0" width="980">
    <tr>
    <td width="260" valign="top" style="text-align:center;">

	<p style="text-align: center; margin-top:0; margin-bottom:3">
	<a href="gosdb.asp"><font color="#404040"><img border="0" src="images/gosdb-small.jpg" align="left"></font></a><font color="#404040"> 
	<b><font style="font-size: 15px">Search by<br>
	</font></b><span style="font-size: 15px"><b>Player</b></span></font><p style="text-align: center; margin-top:0; margin-bottom:0">
	<b>
	<a href="gosdb.asp">Back to<br>GoS-DB Hub</a></b></p>

	</td>
    
  	<td width="460" align="center" style="text-align: center" valign="top">	
	<p style="margin-top:12; margin-bottom:0; text-align:center; font-size:18px; color:#006E32">
    MISCELLANEOUS REPORTS</p>  
    
	<p style="margin-top:6; margin-bottom:0; text-align:center; font-size:13px">
    <b>Report 2: Consecutive Results</b></p>  
    </td>
        
	<td width="260" valign="top"  align="justify">
	'<span style="font-size: 10px">Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
     
    </td>
    </tr>   
	</table>
	<center>
	
	<form style="font-size: 10px; padding: 0; margin: 0;" action="gosdb-misc2.asp" method="post" name="form1">
	
      
<%
competition = Request.Form("competition")
homeaway = Request.Form("homeaway")

select case competition
	case "FLG"
		selected_comp(1) = "selected"
		colpref = "l"
		heading = "Football League"
	case else
		selected_comp(0) = "selected"
		colpref = ""
		heading = "All Competitions"
 end select
 
 outline = "<select name=""competition"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ALL"" " & selected_comp(0) & ">All Competitions</option>"
 outline = outline & "<option value=""FLG"" " & selected_comp(1) & ">Football League</option>"  
 outline = outline & "</select>"
 

 select case homeaway
	case "ho" 
		selected_ha(1) = "selected"
		homeaway_val = "H" 
		heading = heading & " (home only)"
	case "ao" 
		selected_ha(2) = "selected" 
		homeaway_val = "A"
		heading = heading & " (away only)"
 	case else
		selected_ha(0) = "selected"
		homeaway_val = " "
 end select	

 outline = outline & " <select name=""homeaway"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ha"" " & selected_ha(0) & ">Home & Away</option>" 
 outline = outline & "<option value=""ho"" " & selected_ha(1) & ">Home only</option>"
 outline = outline & "<option value=""ao"" " & selected_ha(2) & ">Away only</option>"
 outline = outline & "</select>"
 
 outline = outline & "<br><input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 9 0 0 0"" value=""Select options and click to redisplay"" name=""B1""></p>" 
 
 outline = outline & "</form>"
 
 outline = outline & "<p style=""margin-top:9; margin-bottom:9; text-align:center; font-size:14px; color:#006E32""><b>" & heading & "</b></p>"
 
 response.write(outline)
 
 %>
 
 <table border="0" cellpadding="10" cellspacing="0" style="border-collapse: collapse; margin:0 0 12 0" bordercolor="#111111" width="960">
  <tr>
    <td style="text-align: right">
 <%
	sql = "select top 10 start_" & colpref & "wins, date, "
	sql = sql & "datediff(day, start_" & colpref & "wins, date) as interval, "
	sql = sql & colpref & "wins "
	sql = sql & "from consecutive_results a "
	sql = sql & "where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and not exists ( "
	sql = sql & " select * from consecutive_results b "
	sql = sql & " where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and b.start_" & colpref & "wins = a.start_" & colpref & "wins "
	sql = sql & " and b.date > a.date "
	sql = sql & ") "
	sql = sql & "order by " & colpref & "wins desc, date desc "	

rs.open sql,conn,1,2
	
	outline = "<table class=""tables"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""margin-top: 9; border-collapse: collapse"">"
 	outline  = outline & "<tr>"
    outline  = outline & "<td colspan=""4""><b>Consecutive Wins</b></td>"
	outline  = outline & "</tr>"
	outline  = outline & "<tr>"
	outline  = outline & "<td><b>Games</b></td>"
	outline  = outline & "<td><b>First Match</b></td>"
	outline  = outline & "<td><b>Last Match</b></td>"
	outline  = outline & "<td><b>Days</b></td>"
	outline  = outline & "</tr>"

Do While Not rs.EOF

	work1 = split(FormatDateTime(rs.Fields("start_" & colpref & "wins"),1)," ")
	startdate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	summerdate = "1 Jul"  & " " & work1(2) 
	work1 = split(FormatDateTime(rs.Fields("date"),1)," ")
	enddate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	warn = ""
	if DateDiff("d",startdate,summerdate) > 0 and DateDiff("d",summerdate,enddate) > 0 then warn = "*" 
	datestyle = ""
	if DateDiff("d",startdate,Now) < 1727 then datestyle = " & class=""bold"""
	outline  = outline & "<tr>"
	outline  = outline & "<td class=""center"">" & rs.Fields(colpref & "wins") & "</td>"
	outline  = outline & "<td" & datestyle & ">" & startdate & "</td>"
	outline  = outline & "<td" & datestyle & ">" & enddate & "</td>"
	outline  = outline & "<td class=""right"">" & warn & rs.Fields("interval") & "</td>"
	outline  = outline & "</tr>"
  	rs.MoveNext
Loop
	
rs.close

outline = outline & "</table>"
	
response.write(outline)

%>    
    </td>
    <td style="text-align: center">
<%
	sql = "select top 10 start_" & colpref & "draws, date, "
	sql = sql & "datediff(day, start_" & colpref & "draws, date) as interval, "
	sql = sql & colpref & "draws "
	sql = sql & "from consecutive_results a "
	sql = sql & "where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and not exists ( "
	sql = sql & " select * from consecutive_results b "
	sql = sql & " where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and b.start_" & colpref & "draws = a.start_" & colpref & "draws "
	sql = sql & " and b.date > a.date "
	sql = sql & ") "
	sql = sql & "order by " & colpref & "draws desc, date desc "	

rs.open sql,conn,1,2
	
	outline  = "<table class=""tables"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""margin-top: 9; border-collapse: collapse"">"
 	outline  = outline & "<tr>"
    outline  = outline & "<td colspan=""4""><b>Consecutive Draws</b></td>"
	outline  = outline & "</tr>"
	outline  = outline & "<tr>"
	outline  = outline & "<td><b>Games</b></td>"
	outline  = outline & "<td><b>First Match</b></td>"
	outline  = outline & "<td><b>Last Match</b></td>"
	outline  = outline & "<td><b>Days</b></td>"
	outline  = outline & "</tr>"

Do While Not rs.EOF

	work1 = split(FormatDateTime(rs.Fields("start_" & colpref & "draws"),1)," ")
	startdate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	summerdate = "1 Jul"  & " " & work1(2) 
	work1 = split(FormatDateTime(rs.Fields("date"),1)," ")
	enddate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	warn = ""
	if DateDiff("d",startdate,summerdate) > 0 and DateDiff("d",summerdate,enddate) > 0 then warn = "*" 
	datestyle = ""
	if DateDiff("d",startdate,Now) < 1727 then datestyle = " & class=""bold""" 
	outline  = outline & "<tr>"
	outline  = outline & "<td class=""center"">" & rs.Fields(colpref & "draws") & "</td>"
	outline  = outline & "<td" & datestyle & ">" & startdate & "</td>"
	outline  = outline & "<td" & datestyle & ">" & enddate & "</td>"
	outline  = outline & "<td class=""right"">" & warn & rs.Fields("interval") & "</td>"
	outline  = outline & "</tr>"
  	rs.MoveNext
Loop
	
rs.close

outline = outline & "</table>"
	
response.write(outline)

%>    
    </td>
    <td>
<%
	sql = "select top 10 start_" & colpref & "defeats, date, "
	sql = sql & "datediff(day, start_" & colpref & "defeats, date) as interval, "
	sql = sql & colpref & "defeats "
	sql = sql & "from consecutive_results a "
	sql = sql & "where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and not exists ( "
	sql = sql & " select * from consecutive_results b "
	sql = sql & " where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and b.start_" & colpref & "defeats = a.start_" & colpref & "defeats "
	sql = sql & " and b.date > a.date "
	sql = sql & ") "
	sql = sql & "order by " & colpref & "defeats desc, date desc "	

rs.open sql,conn,1,2
	
	outline  = "<table class=""tables"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""margin-top: 9; border-collapse: collapse"">"
 	outline  = outline & "<tr>"
    outline  = outline & "<td colspan=""4""><b>Consecutive Defeats</b></td>"
	outline  = outline & "</tr>"
	outline  = outline & "<tr>"
	outline  = outline & "<td><b>Games</b></td>"
	outline  = outline & "<td><b>First Match</b></td>"
	outline  = outline & "<td><b>Last Match</b></td>"
	outline  = outline & "<td><b>Days</b></td>"
	outline  = outline & "</tr>"

Do While Not rs.EOF

	work1 = split(FormatDateTime(rs.Fields("start_" & colpref & "defeats"),1)," ")
	startdate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	summerdate = "1 Jul"  & " " & work1(2) 
	work1 = split(FormatDateTime(rs.Fields("date"),1)," ")
	enddate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	warn = ""
	if DateDiff("d",startdate,summerdate) > 0 and DateDiff("d",summerdate,enddate) > 0 then warn = "*" 
	datestyle = ""
	if DateDiff("d",startdate,Now) < 1727 then datestyle = " & class=""bold""" 
	outline  = outline & "<tr>"
	outline  = outline & "<td class=""center"">" & rs.Fields(colpref & "defeats") & "</td>"
	outline  = outline & "<td" & datestyle & ">" & startdate & "</td>"
	outline  = outline & "<td" & datestyle & ">" & enddate & "</td>"
	outline  = outline & "<td class=""right"">" & warn & rs.Fields("interval") & "</td>"
	outline  = outline & "</tr>"
  	rs.MoveNext
Loop
	
rs.close

outline = outline & "</table>"
	
response.write(outline)

%>    
    </td>
  </tr>
  <tr>
    <td style="text-align: right">
<%
	sql = "select top 10 start_" & colpref & "nodefeats, date, "
	sql = sql & "datediff(day, start_" & colpref & "nodefeats, date) as interval, "
	sql = sql & colpref & "nodefeats "
	sql = sql & "from consecutive_results a "
	sql = sql & "where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and not exists ( "
	sql = sql & " select * from consecutive_results b "
	sql = sql & " where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and b.start_" & colpref & "nodefeats = a.start_" & colpref & "nodefeats "
	sql = sql & " and b.date > a.date "
	sql = sql & ") "
	sql = sql & "order by " & colpref & "nodefeats desc, date desc"	

rs.open sql,conn,1,2
	
	outline  = "<table class=""tables"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""margin-top: 9; border-collapse: collapse"">"
 	outline  = outline & "<tr>"
    outline  = outline & "<td colspan=""4""><b>Without Losing</b></td>"
	outline  = outline & "</tr>"
	outline  = outline & "<tr>"
	outline  = outline & "<td><b>Games</b></td>"
	outline  = outline & "<td><b>First Match</b></td>"
	outline  = outline & "<td><b>Last Match</b></td>"
	outline  = outline & "<td><b>Days</b></td>"
	outline  = outline & "</tr>"

Do While Not rs.EOF

	work1 = split(FormatDateTime(rs.Fields("start_" & colpref & "nodefeats"),1)," ")
	startdate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	summerdate = "1 Jul"  & " " & work1(2) 
	work1 = split(FormatDateTime(rs.Fields("date"),1)," ")
	enddate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	warn = ""
	if DateDiff("d",startdate,summerdate) > 0 and DateDiff("d",summerdate,enddate) > 0 then warn = "*" 
	datestyle = ""
	if DateDiff("d",startdate,Now) < 1727 then datestyle = " & class=""bold""" 
	outline  = outline & "<tr>"
	outline  = outline & "<td class=""center"">" & rs.Fields(colpref & "nodefeats") & "</td>"
	outline  = outline & "<td" & datestyle & ">" & startdate & "</td>"
	outline  = outline & "<td" & datestyle & ">" & enddate & "</td>"
	outline  = outline & "<td class=""right"">" & warn & rs.Fields("interval") & "</td>"
	outline  = outline & "</tr>"
  	rs.MoveNext
Loop
	
rs.close

outline = outline & "</table>"
	
response.write(outline)

%>    
    </td>
    <td style="text-align: center">
<%
	sql = "select top 10 start_" & colpref & "nowins, date, "
	sql = sql & "datediff(day, start_" & colpref & "nowins, date) as interval, "
	sql = sql & colpref & "nowins "
	sql = sql & "from consecutive_results a "
	sql = sql & "where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and not exists ( "
	sql = sql & " select * from consecutive_results b "
	sql = sql & " where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and b.start_" & colpref & "nowins = a.start_" & colpref & "nowins "
	sql = sql & " and b.date > a.date "
	sql = sql & ") "
	sql = sql & "order by " & colpref & "nowins desc, date desc"	

rs.open sql,conn,1,2
	
	outline  = "<table class=""tables"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""margin-top: 9; border-collapse: collapse"">"
 	outline  = outline & "<tr>"
    outline  = outline & "<td style=""border-bottom: 0px none;"" colspan=""4""><b>Without Winning</b></td>"
	outline  = outline & "</tr>"
	outline  = outline & "<tr>"
	outline  = outline & "<td><b>Games</b></td>"
	outline  = outline & "<td><b>First Match</b></td>"
	outline  = outline & "<td><b>Last Match</b></td>"
	outline  = outline & "<td><b>Days</b></td>"
	outline  = outline & "</tr>"

Do While Not rs.EOF

	work1 = split(FormatDateTime(rs.Fields("start_" & colpref & "nowins"),1)," ")
	startdate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	summerdate = "1 Jul"  & " " & work1(2) 
	work1 = split(FormatDateTime(rs.Fields("date"),1)," ")
	enddate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	warn = ""
	if DateDiff("d",startdate,summerdate) > 0 and DateDiff("d",summerdate,enddate) > 0 then warn = "*" 
	datestyle = ""
	if DateDiff("d",startdate,Now) < 1727 then datestyle = " & class=""bold"""
	outline  = outline & "<tr>"
	outline  = outline & "<td class=""center"">" & rs.Fields(colpref & "nowins") & "</td>"
	outline  = outline & "<td" & datestyle & ">" & startdate & "</td>"
	outline  = outline & "<td" & datestyle & ">" & enddate & "</td>"
	outline  = outline & "<td class=""right"">" & warn & rs.Fields("interval") & "</td>"
	outline  = outline & "</tr>"
  	rs.MoveNext
Loop
	
rs.close


outline = outline & "</table>"
	
response.write(outline)

%>    
    </td>
    <td style="text-align: center">
<%
	sql = "select top 10 start_" & colpref & "cleansheets, date, "
	sql = sql & "datediff(day, start_" & colpref & "cleansheets, date) as interval, "
	sql = sql & colpref & "cleansheets "
	sql = sql & "from consecutive_results a "
	sql = sql & "where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and not exists ( "
	sql = sql & " select * from consecutive_results b "
	sql = sql & " where homeawayall = '" & homeaway_val & "' "
	sql = sql & " and b.start_" & colpref & "cleansheets = a.start_" & colpref & "cleansheets "
	sql = sql & " and b.date > a.date "
	sql = sql & ") "
	sql = sql & "order by " & colpref & "cleansheets desc, date desc "	

rs.open sql,conn,1,2
	
	outline  = "<table class=""tables"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""margin-top: 9; border-collapse: collapse"">"
 	outline  = outline & "<tr>"
    outline  = outline & "<td style=""border-bottom: 0px none;"" colspan=""4""><b>Clean Sheets</b></td>"
	outline  = outline & "</tr>"
	outline  = outline & "<tr>"
	outline  = outline & "<td><b>Games</b></td>"
	outline  = outline & "<td><b>First Match</b></td>"
	outline  = outline & "<td><b>Last Match</b></td>"
	outline  = outline & "<td><b>Days</b></td>"
	outline  = outline & "</tr>"

Do While Not rs.EOF

	work1 = split(FormatDateTime(rs.Fields("start_" & colpref & "cleansheets"),1)," ")
	startdate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	summerdate = "1 Jul"  & " " & work1(2) 
	work1 = split(FormatDateTime(rs.Fields("date"),1)," ")
	enddate = work1(0) & " " & left(work1(1),3) & " " & work1(2)
	warn = ""
	if DateDiff("d",startdate,summerdate) > 0 and DateDiff("d",summerdate,enddate) > 0 then warn = "*" 
	datestyle = ""
	if DateDiff("d",startdate,Now) < 1727 then datestyle = " & class=""bold"""
	outline  = outline & "<tr>"
	outline  = outline & "<td class=""center"">" & rs.Fields(colpref & "cleansheets") & "</td>"
	outline  = outline & "<td" & datestyle & ">" & startdate & "</td>"
	outline  = outline & "<td" & datestyle & ">" & enddate & "</td>"
	outline  = outline & "<td class=""right"">" & warn & rs.Fields("interval") & "</td>"
	outline  = outline & "</tr>"
  	rs.MoveNext
Loop
	
rs.close


outline = outline & "</table>"
	
response.write(outline)

%>    
    </td>

  	</tr>
  	<tr><td colspan="6">
  	<p class="header1" style="margin-top: 18; margin-bottom: 0; " align="center">
	<span style="font-weight: 400">Green background for dates in the last 5 
    years </span></p>
	<p class="style1" style="margin-top: 9; margin-bottom: 0; " align="center">
	<span style="font-weight: 400">* Indicates that the period spans seasons</span></p>
	</td></tr>
</table>

<%
conn.close
%>	
	


</center><br>

<!--#include file="base_code.htm"-->
</body>

</html>