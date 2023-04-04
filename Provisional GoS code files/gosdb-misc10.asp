<%@ Language=VBScript %> 
<% Option Explicit %>

<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>GoS-DB Miscellaneous Report</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
.tables td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1}
.tables .right {text-align: right;}-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline, heading, season_no1, selected_s1, season_no2, selected_s2, season1opts, season2opts, selyears1, selyears2, lastyears
Dim selected_comp(3), competition, tableview, restrictions, yearssave, notfirst


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
    
	<p style="margin-top:6px; margin-bottom:0; text-align:center; font-size:13px">
    <b>Report 10: Goalscorers Ranked by Season</b></p>
    
 	<p style="margin-top: 12px; margin-bottom: 4px; ">This report lists goalscorers 
    for each selected season. Change the options below to display a range of seasons, or 
    for league or cup competitions only.</p>
  
 	<p style="margin-top: 4px; margin-bottom: 6px; ">The names revealed are also links to the 
    associated player page.</p>
  
    </td>
        
	<td width="260" valign="top"  align="justify">
	'Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch.
     
    </td>
    </tr>   
	</table>
	<center>
		
	<form style="font-size: 10px; padding: 0; margin: 0;" action="gosdb-misc10.asp" method="post" name="form1">
	
      
<%
competition = Request.Form("competition")
season_no1 = Cint(Request.Form("season1"))
season_no2 = Cint(Request.Form("season2"))

select case competition
	case "LG"
		selected_comp(1) = "selected"
		tableview = "v_match_all_league"
		heading = "League"
	case "FAC"
		selected_comp(2) = "selected"
		tableview = "v_match_FA"
		heading = "FA Cup"
	case "CUP"
		selected_comp(3) = "selected"
		tableview = "v_match_cups"
		heading = "All Cups"
	case else
		selected_comp(0) = "selected"
		tableview = "v_match_all"
		heading = "All Competitions"
 end select
 
 outline = "<select name=""competition"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ALL"" " & selected_comp(0) & ">All Competitions</option>"
 outline = outline & "<option value=""LG"" " & selected_comp(1) & ">League</option>" 
 outline = outline & "<option value=""FAC"" " & selected_comp(2) & ">FA Cup</option>"
 outline = outline & "<option value=""CUP"" " & selected_comp(3) & ">All Cups</option>" 
 outline = outline & "</select>"

 sql = "select season_no, years "
 sql = sql & "from season "
 rs.open sql,conn,1,2
 
 if season_no1 = 0 then season_no1 = Cint(rs.RecordCount)
 if season_no2 = 0 then season_no2 = Cint(rs.RecordCount)
 
 season1opts = ""
 season2opts = ""
   
 Do While Not rs.EOF
  if Cint(rs.Fields("season_no")) = season_no1 then 
    selected_s1 = "selected"
    if season_no1 > 1 then
    	selyears1 = rs.Fields("years") 
    end if
   else selected_s1 = ""
  end if
  season1opts = season1opts & "<option value=""" & rs.Fields("season_no") & """ " & selected_s1 & ">From " & rs.Fields("years") & "</option>"
  if Cint(rs.Fields("season_no")) = season_no2 then
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
 
 if season_no1 > 1 and season_no2 < Cint(rs.RecordCount) then heading = heading & " from " & selyears1 & " to " & selyears2
 if season_no1 = 1 and season_no2 < Cint(rs.RecordCount) then heading = heading & " from 1903-1904 to " & selyears2
 if season_no1 > 1 and season_no2 = Cint(rs.RecordCount) then heading = heading & " from " & selyears1 & " to " & lastyears
 
 rs.close
 
 outline = outline & " <select name=""season1"" style=""font-size: 10px"">" & season1opts & "</select>"  
 outline = outline & " <select name=""season2"" style=""font-size: 10px"">" & season2opts & "</select>"
 outline = outline & "<br><input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 9 0 0 0"" value=""Select options and click to redisplay"" name=""B1""></p>"  
 outline = outline & "</form>"
 
 outline = outline & "<p style=""margin-top:12px; margin-bottom:12px; text-align:center; font-size:13px; color:#006E32""><b>" & heading & "</b></p>"
 
 response.write(outline)
 %>
 	  
	<%
	outline = ""

   	sql = "with CTE1 as " 
	sql = sql & "( " 	
	sql = sql & "select years, player_id_spell1, surname, nullif(forename, initials) as firstname, " 
	sql = sql & " count(c.player_id) as goals, count(distinct b.date) as games, round(count(c.player_id)/cast(count(distinct b.date) as dec(7,3)),2) as pergame " 
	sql = sql & "from " & tableview & " a join season on date between date_start and date_end " 
	sql = sql & "join match_player b on a.date = b.date "  
	sql = sql & "left outer join match_goal c on b.player_id = c.player_id and b.date = c.date "  
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where season_no between " & season_no1 & " and " & season_no2
	sql = sql & "group by years, player_id_spell1, forename, surname, initials "
 	sql = sql & "), "
	sql = sql & "CTE2 as " 
	sql = sql & "( " 	
	sql = sql & "select player_id_spell1, count(distinct years) as seasons, " 
	sql = sql & " count(c.player_id) as goals, count(distinct b.date) as games, round(count(c.player_id)/cast(count(distinct b.date) as dec(7,3)),2) as pergame " 
	sql = sql & "from " & tableview & " a join season on date between date_start and date_end " 
	sql = sql & "join match_player b on a.date = b.date "  
	sql = sql & "left outer join match_goal c on b.player_id = c.player_id and b.date = c.date "  
	sql = sql & "join player d on b.player_id = d.player_id " 
	sql = sql & "group by player_id_spell1 "
	sql = sql & ") " 
	sql = sql & "select a.player_id_spell1, a.years, rank() over (partition by a.years order by a.goals desc) as rank, trim(isnull(rtrim(a.firstname),'') + ' ' + a.surname) as player, "
	sql = sql & " a.goals, a.games, a.pergame, seasons, b.goals as totgoals, b.games as totgames, b.pergame as totpergame "
	sql = sql & "from CTE1 a join CTE2 b on a.player_id_spell1 = b.player_id_spell1 "
	sql = sql & "where a.goals > 0 "
	sql = sql & "order by a.years, rank, a.surname "

	rs.open sql,conn,1,2
	

	outline = outline & "<table class=""tables"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"">"
	  
	Do While Not rs.EOF
	
		if rs.Fields("years") = yearssave then
			outline = outline & "<tr>"
			outline = outline & "<td></td>"
		  else
		  	if notfirst = "y" then outline = outline & "<tr><td colspan=""10"" height=""10px""> </td></tr>"
		  	notfirst = "y"
		  	outline = outline & "<tr>"
    		outline = outline & "<td colspan=""3""></td>"
    		outline = outline & "<td colspan=""3"">This Season</td>"
    		outline = outline & "<td colspan=""4"">PAFC Career</td>"
			outline = outline & "</tr>"
 			outline = outline & "<tr>"
 			outline = outline & "<td>Season</td>"
    		outline = outline & "<td>Rank</td>"
    		outline = outline & "<td>Player</td>"
			outline = outline & "<td>Goals</td>"
			outline = outline & "<td>Games</td>"
			outline = outline & "<td>Goals/<br>Game</td>"
			outline = outline & "<td>Sea-<br>sons</td>"  
			outline = outline & "<td>Goals</td>"
			outline = outline & "<td>Games</td>"
			outline = outline & "<td>Goals/<br>Game</td>"
			outline = outline & "</tr>"
			outline = outline & "<td>" & rs.Fields("years") & "</td>"
			yearssave = rs.Fields("years")
		end if
		
		outline = outline & "<td class=""right"">" & rs.Fields("rank") & "</td>"
		outline = outline & "<td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """>" & rs.Fields("player") & "</a></td>" 
		outline = outline & "<td class=""right"">" & rs.Fields("goals") & "</td>"
		outline = outline & "<td class=""right"">" & rs.Fields("games") & "</td>"
		outline = outline & "<td class=""right"">" & rs.Fields("pergame") & "</td>" 
		outline = outline & "<td class=""right"">" & rs.Fields("seasons") & "</td>" 
		outline = outline & "<td class=""right"">" & rs.Fields("totgoals") & "</td>"
		outline = outline & "<td class=""right"">" & rs.Fields("totgames") & "</td>" 
		outline = outline & "<td class=""right"">" & rs.Fields("totpergame") & "</td>"
		rs.MoveNext 
	Loop 

	rs.close
	outline = outline & "</table>"
	
response.write(outline)	

conn.close
%>	
</center><br>

<!--#include file="base_code.htm"-->
</body>

</html>