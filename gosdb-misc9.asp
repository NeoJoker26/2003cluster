<%@ Language=VBScript %> 
<% Option Explicit %>
<% dim scope
scope = Request.Form("scope")
if scope = "" then scope = Request.Querystring("scope")	'try for a url parameter
scope = replace(scope," ","")
%>
<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Author" content="Trevor Scallan">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<title>GoS-DB Miscellaneous Report</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
#table1 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1}
#table2 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1}
.med{font-size:medium;font-weight:normal;padding:0;margin:0}#res{padding-right:1em;margin:0 16px}ol li{list-style:none}.g{margin:1em 0}li.g{font-size:small;font-family:arial,sans-serif}.s{max-width:42em}-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline, heading, season_no1, selected_s1, season_no2, selected_s2, season1opts, season2opts, selyears1, selyears2, lastyears
Dim selected_comp(3), competition, tableview, restrictions


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
    <b>Report 9: Success Ranking by Opposition</b></p>  
    </td>
        
	<td width="260" valign="top"  align="justify">
	'<span style="font-size: 10px">Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
     
    </td>
    </tr>   
	</table>
	<center>
		
	<form style="font-size: 10px; padding: 0; margin: 0;" action="gosdb-misc9.asp" method="post" name="form1">
	
      
<%
competition = Request.Form("competition")
season_no1 = Request.Form("season1")
season_no2 = Request.Form("season2")

select case competition
	case "FLG"
		selected_comp(1) = "selected"
		tableview = "v_match_FL"
		heading = "Football League"
	case "CUP"
		selected_comp(2) = "selected"
		tableview = "v_match_cups"
		heading = "Cup Competitions"
	case else
		selected_comp(0) = "selected"
		tableview = "v_match_all"
		heading = "All Competitions"
 end select
 
 outline = "<select name=""competition"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ALL"" " & selected_comp(0) & ">All Competitions</option>"
 outline = outline & "<option value=""FLG"" " & selected_comp(1) & ">Football League</option>" 
 outline = outline & "<option value=""CUP"" " & selected_comp(2) & ">All Cups</option>" 
 outline = outline & "</select>"

 sql = "select season_no, years "
 sql = sql & "from season "
 rs.open sql,conn,1,2
 
 if season_no1 = "" then season_no1 = 1
 if season_no2 = "" then season_no2 = CStr(rs.RecordCount)
 
 season1opts = ""
 season2opts = ""
   
 Do While Not rs.EOF
  if CStr(rs.Fields("season_no")) = season_no1 then 
    selected_s1 = "selected"
    if season_no1 > 1 then
    	selyears1 = rs.Fields("years") 
    end if
   else selected_s1 = ""
  end if
  season1opts = season1opts & "<option value=""" & rs.Fields("season_no") & """ " & selected_s1 & ">From " & rs.Fields("years") & "</option>"
  if CStr(rs.Fields("season_no")) = season_no2 then
    selected_s2 = "selected"
    if season_no2 < CStr(rs.RecordCount) then 
    selyears2 = rs.Fields("years") 
    end if
   else selected_s2 = ""
  end if
  season2opts = season2opts & "<option value=""" & rs.Fields("season_no") & """ " & selected_s2 & ">To " & rs.Fields("years") & "</option>"
  lastyears = rs.Fields("years")
 rs.MoveNext
 Loop
 
 if season_no1 > 1 and season_no2 < CStr(rs.RecordCount) then heading = heading & " from " & selyears1 & " to " & selyears2
 if season_no1 = 1 and season_no2 < CStr(rs.RecordCount) then heading = heading & " from 1903-1904 to " & selyears2
 if season_no1 > 1 and season_no2 = CStr(rs.RecordCount) then heading = heading & " from " & selyears1 & " to " & lastyears
 
 rs.close
 
 outline = outline & " <select name=""season1"" style=""font-size: 10px"">" & season1opts & "</select>"  
 outline = outline & " <select name=""season2"" style=""font-size: 10px"">" & season2opts & "</select>"
 outline = outline & "<br><input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 9 0 0 0"" value=""Select options and click to redisplay"" name=""B1""></p>"  
 outline = outline & "</form>"
 
 outline = outline & "<p style=""margin-top:6; margin-bottom:0; text-align:center; font-size:13px; color:#006E32""><b>" & heading & "</b></p>"
 
 response.write(outline)
 %>
</center>
 
 	<div style="width:600; text-align:justify">
 	<p style="margin-top: 12; margin-bottom: 6">These rankings are based on 
    Argyle's success against teams played more 
    than three times at home or away (to be consistent with the success rankings 
    on GoS-DB's opposition pages) and more than six times in total. 
    'Success' is determined by the average points per game, based on three 
    points for a win and one for a draw in all cases, even for cup matches.&nbsp;If a team 
    has changed its name at any time in its history, all results are combined under 
    its name in the most recent game against them.</p>
 	</div>
 	
	<center>
		
  <center>
 	
 	<table border="0" cellpadding="10" cellspacing="10" style="border-collapse: collapse; margin:0 0 12 0" bordercolor="#111111">
  	<tr>
    <td valign="top">
      
	<%
	outline = ""
	
	sql = "WITH CTE1 AS "
	sql = sql & "( "
	sql = sql & "select homeaway, name_now, cast(sum(points) as dec(7,2))/SUM(p) as pointspergame, sum(p) as p, sum(w) as w, sum(d) as d, sum(l) as l "
	sql = sql & "from ( "
	sql = sql & " select homeaway, name_now, 1 as p, "
	sql = sql & " case when goalsfor > goalsagainst then 1 else 0 end as w, "
	sql = sql & " case when goalsfor = goalsagainst then 1 else 0 end as d, "
	sql = sql & " case when goalsfor < goalsagainst then 1 else 0 end as l, "
	sql = sql & " case when goalsfor > goalsagainst then 3 " 
	sql = sql & " when goalsfor = goalsagainst then 1 " 
	sql = sql & " when goalsfor < goalsagainst then 0 end as points "
	sql = sql & " from " & tableview 
	sql = sql & " join opposition on opposition = name_then join season on date between date_start and date_end "
	sql = sql & " where homeaway in ('H','A') " 
	sql = sql & "  and season_no between " & season_no1 & " and " & season_no2
	sql = sql & ") as sub "
	sql = sql & "group by homeaway, name_now "
	sql = sql & "having count(*) > 3 "
	sql = sql & "), "
	sql = sql & "CTE2 as "
	sql = sql & "( "
	sql = sql & "select rank() over(partition by homeaway order by homeaway, pointspergame desc) as rank, homeaway, pointspergame, name_now, p, w, d, l "
	sql = sql & "from CTE1 "
	sql = sql & ") "
	sql = sql & "select rank, homeaway, pointspergame, name_now, p, w, d, l "
	sql = sql & "from CTE2 "
	sql = sql & "order by homeaway desc, rank "

	rs.open sql,conn,1,2

	outline = outline & "<table id=""table1"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"">"
 	outline  = outline & "<tr>"
    outline  = outline & "<td colspan=""2""><p style=""margin: 3 0 3 0""><b>Success Ranking at Home</b></p></td>"
	outline  = outline & "<td>P</td>"
	outline  = outline & "<td>W</td>" 
	outline  = outline & "<td>D</td>" 
	outline  = outline & "<td>L</td>" 
	outline  = outline & "<td>Pts/<br>Game</td>"
	outline  = outline & "</tr>"
	  
	Do While Not rs.EOF and rs.Fields("homeaway") = "H" 
	
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td>" & rs.Fields("name_now") & "</td>" 
		outline  = outline & "<td>" & rs.Fields("p") & "</td>"
		outline  = outline & "<td>" & rs.Fields("w") & "</td>" 
		outline  = outline & "<td>" & rs.Fields("d") & "</td>" 
		outline  = outline & "<td>" & rs.Fields("l") & "</td>" 
		outline  = outline & "<td>" & round(rs.Fields("pointspergame"),3) & "</td>"
		outline  = outline & "</tr>"  
	
		rs.MoveNext 
	Loop 

	outline  = outline & "</table>"	
	
    outline = outline & "</td>"
	outline = outline & "<td valign=""top"">" 
	
	outline = outline & "<table id=""table2"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"">"
 	outline  = outline & "<tr>"
    outline  = outline & "<td colspan=""2""><p style=""margin: 3 0 3 0""><b>Success Ranking Away</b></p></td>"
	outline  = outline & "<td>P</td>"
	outline  = outline & "<td>W</td>" 
	outline  = outline & "<td>D</td>" 
	outline  = outline & "<td>L</td>" 
	outline  = outline & "<td>Pts/<br>Game</td>"
	outline  = outline & "</tr>"


	Do While Not rs.EOF
			
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td>" & rs.Fields("name_now") & "</td>" 
		outline  = outline & "<td>" & rs.Fields("p") & "</td>"
		outline  = outline & "<td>" & rs.Fields("w") & "</td>" 
		outline  = outline & "<td>" & rs.Fields("d") & "</td>" 
		outline  = outline & "<td>" & rs.Fields("l") & "</td>" 
		outline  = outline & "<td>" & round(rs.Fields("pointspergame"),3) & "</td>"
		outline  = outline & "</tr>" 
 
  		rs.MoveNext
	Loop
	
	rs.close
	
	outline  = outline & "</table>"
	
	outline = outline & "</td>"
	outline = outline & "<td valign=""top"">" 

	
sql = "WITH CTE1 AS "
	sql = sql & "( "
	sql = sql & "select name_now, cast(sum(points) as dec(7,2))/SUM(p) as pointspergame, sum(p) as p, sum(w) as w, sum(d) as d, sum(l) as l "
	sql = sql & "from ( "
	sql = sql & " select name_now, 1 as p, "
	sql = sql & " case when goalsfor > goalsagainst then 1 else 0 end as w, "
	sql = sql & " case when goalsfor = goalsagainst then 1 else 0 end as d, "
	sql = sql & " case when goalsfor < goalsagainst then 1 else 0 end as l, "
	sql = sql & " case when goalsfor > goalsagainst then 3 " 
	sql = sql & " when goalsfor = goalsagainst then 1 " 
	sql = sql & " when goalsfor < goalsagainst then 0 end as points "
	sql = sql & " from " & tableview 
	sql = sql & " join opposition on opposition = name_then join season on date between date_start and date_end "
	sql = sql & "  and season_no between " & season_no1 & " and " & season_no2
	sql = sql & ") as sub "
	sql = sql & "group by name_now "
	sql = sql & "having count(*) > 6 "
	sql = sql & "), "
	sql = sql & "CTE2 as "
	sql = sql & "( "
	sql = sql & "select rank() over(order by pointspergame desc) as rank, pointspergame, name_now, p, w, d, l "
	sql = sql & "from CTE1 "
	sql = sql & ") "
	sql = sql & "select rank, pointspergame, name_now, p, w, d, l "
	sql = sql & "from CTE2 "
	sql = sql & "order by rank "

	rs.open sql,conn,1,2

	outline = outline & "<table id=""table1"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"">"
 	outline  = outline & "<tr>"
    outline  = outline & "<td colspan=""2""><p style=""margin: 3 0 3 0""><b>Success Ranking Home & Away</b></p></td>"
	outline  = outline & "<td>P</td>"
	outline  = outline & "<td>W</td>" 
	outline  = outline & "<td>D</td>" 
	outline  = outline & "<td>L</td>" 
	outline  = outline & "<td>Pts/<br>Game</td>"
	outline  = outline & "</tr>"
	  
	Do While Not rs.EOF
	
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td>" & rs.Fields("name_now") & "</td>" 
		outline  = outline & "<td>" & rs.Fields("p") & "</td>"
		outline  = outline & "<td>" & rs.Fields("w") & "</td>" 
		outline  = outline & "<td>" & rs.Fields("d") & "</td>" 
		outline  = outline & "<td>" & rs.Fields("l") & "</td>" 
		outline  = outline & "<td>" & round(rs.Fields("pointspergame"),3) & "</td>"
		outline  = outline & "</tr>"  
	
		rs.MoveNext 
	Loop 
	
	rs.close
	
	outline  = outline & "</table>"
		
	outline  = outline & "</td></tr>"	
	outline  = outline & "</table>"
	
	response.write(outline)
	%>
	</td>
</tr>

<%
conn.close
%>	
	
</table>
</center><br>

<!--#include file="base_code.htm"-->
</body>

</html>