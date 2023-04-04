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
#table1 td {border: 1px solid #c0c0c0; margin: 0; white-space:nowrap; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1}
#table2 td {border: 1px solid #c0c0c0; margin: 0; white-space:nowrap; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1}

-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline, heading, playername
Dim selected_comp(3), competition, view, result, restrictions


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
    <b>Report 4: Top Substitutes</b></p>  
    </td>
        
	<td width="260" valign="top"  align="justify">
	<span style="font-size: 10px">'Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
     
    </td>
    </tr>   
	</table>
	<center>
		
	<form style="font-size: 10px; padding: 0; margin: 0;" action="gosdb-misc4.asp" method="post" name="form1">
	
<%
competition = Request.Form("competition")

select case competition
	case "FLG"
		selected_comp(1) = "selected"
		view = "v_match_FL"
		heading = "Football League"
	case "FAC"
		selected_comp(2) = "selected"
		view = "v_match_FA"
		heading = "FA Cup"
	case "FLC"
		selected_comp(3) = "selected"
		view = "v_match_LC"
		heading = "League Cup"
	case else
		selected_comp(0) = "selected"
		view = "v_match_all"
		heading = "All Competitions"
 end select
 
 outline = "<select name=""competition"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ALL"" " & selected_comp(0) & ">All Competitions</option>"
 outline = outline & "<option value=""FLG"" " & selected_comp(1) & ">Football League</option>" 
 outline = outline & "<option value=""FAC"" " & selected_comp(2) & ">FA Cup</option>" 
 outline = outline & "<option value=""FLC"" " & selected_comp(3) & ">League Cup</option>"
 outline = outline & "</select>"

 
 outline = outline & "<br><input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 9 0 0 0"" value=""Select option and click to redisplay"" name=""B1""></p>"  
 outline = outline & "</form>"
	
 response.write(outline)
 %>
 
 	<table border="0" style="border-collapse: collapse; margin:6 0 12 0" bordercolor="#111111">
  	<tr>
    <td valign="top" style="padding-right: 15px">
    <table id="table1" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr><td colspan="2"><b>ALL PLAYERS</b></td></tr>
      
	<%
	outline = ""
	
	sql = "select count(distinct player_id_spell1) as count " 
	sql = sql & "from " & view & " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	rs.open sql,conn,1,2		
	outline  = outline & "<tr><td>Players making any appearance</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	rs.close	
	
	sql = "select count(distinct player_id_spell1) as count " 
	sql = sql & "from " & view & " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where startpos > 0 "	
	rs.open sql,conn,1,2		
	outline  = outline & "<tr><td>Players who have started a game</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	rs.close	

	sql = "select count(distinct player_id_spell1) as count " 
	sql = sql & "from " & view & " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where startpos = 0 "
	rs.open sql,conn,1,2		
	outline  = outline & "<tr><td>Players who have appeared as a substitute</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	rs.close

	sql = "select count(*) as count " 
	sql = sql & "from " & view & " a join match_player b on a.date = b.date "
	sql = sql & "where startpos = 0 "
	rs.open sql,conn,1,2		
	outline  = outline & "<tr><td>Total player substitutions</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	rs.close
	
	response.write(outline)
	outline = ""
	%>
	<tr><td colspan="2">
      <p style="margin-top: 6"><b>Highest Substitute Appearances</b></td></tr>
	<%	
	sql = "select top 5 player_id_spell1, surname, forename, initials, count(*) as count " 
	sql = sql & "from " & view & " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where startpos = 0 "
	sql = sql & "group by player_id_spell1, surname, forename, initials "
	sql = sql & "order by count desc, surname "
	rs.open sql,conn,1,2
	Do While Not rs.EOF
		if not IsNull(rs.Fields("forename")) then 
			playername = trim(rs.Fields("forename")) & " " & trim(rs.Fields("surname")) 
	  	  else
			playername = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname"))
		end if		
		outline  = outline & "<tr><td>" & playername & "</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	  	rs.MoveNext
	Loop
	rs.close
	
	response.write(outline)
	outline = ""
	%>
	<tr><td colspan="2">
      <p style="margin-top: 6"><b>Top Substitute Goalscorers</b></td></tr>
	<%	
	sql = "select top 5 player_id_spell1, surname, forename, initials, count(*) as count " 
	sql = sql & "from " & view & " a join match_player b1 on a.date = b1.date "
	sql = sql & " join match_goal b2 on a.date = b2.date and b1.player_id = b2.player_id "
	sql = sql & " join player c on b1.player_id = c.player_id "
	sql = sql & "where startpos = 0 "
	sql = sql & "group by player_id_spell1, surname, forename, initials "
	sql = sql & "order by count desc, surname "
	rs.open sql,conn,1,2
	Do While Not rs.EOF
		if not IsNull(rs.Fields("forename")) then 
			playername = trim(rs.Fields("forename")) & " " & trim(rs.Fields("surname")) 
	  	  else
			playername = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname"))
		end if		
		outline  = outline & "<tr><td>" & playername & "</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	  	rs.MoveNext
	Loop
	rs.close

	outline  = outline & "</table>"
	response.write(outline)
	outline = ""
	%>
	
	</td>
    <td valign="top"  style="padding-left: 15px">
    <table id="table2" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr><td colspan="2"><b>CURRENT SQUAD</b></td></tr>
      
	<%
	sql = "select count(distinct player_id_spell1) as count " 
	sql = sql & "from " & view & " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where last_game_year = 9999 " 
	rs.open sql,conn,1,2		
	outline  = outline & "<tr><td>Players making any appearance</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	rs.close
		
	sql = "select count(distinct player_id_spell1) as count " 
	sql = sql & "from " & view & " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where startpos > 0 " 
	sql = sql & "  and last_game_year = 9999 "
	rs.open sql,conn,1,2		
	outline  = outline & "<tr><td>Players who have started a game</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	rs.close	
	
	sql = "select count(distinct player_id_spell1) as count " 
	sql = sql & "from " & view & " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where startpos = 0 "
	sql = sql & "  and last_game_year = 9999 "
	rs.open sql,conn,1,2		
	outline  = outline & "<tr><td>Players who have appeared as a substitute</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	rs.close

	sql = "select count(*) as count " 
	sql = sql & "from " & view & " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where startpos = 0 "
	sql = sql & "  and last_game_year = 9999 "
	rs.open sql,conn,1,2		
	outline  = outline & "<tr><td>Total player substitutions</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	rs.close
	
	response.write(outline)
	outline = ""
	%>
	<tr><td colspan="2">
      <p style="margin-top: 6"><b>Highest Substitute Appearances</b></td></tr>
	<%
	sql = "select top 5 player_id_spell1, surname, forename, initials, count(*) as count " 
	sql = sql & "from " & view & " a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where startpos = 0 "
	sql = sql & "  and last_game_year = 9999 "
	sql = sql & "group by player_id_spell1, surname, forename, initials "
	sql = sql & "order by count desc, surname "
	rs.open sql,conn,1,2
	Do While Not rs.EOF
		if not IsNull(rs.Fields("forename")) then 
			playername = trim(rs.Fields("forename")) & " " & trim(rs.Fields("surname")) 
	  	  else
			playername = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname"))
		end if		
		outline  = outline & "<tr><td>" & playername & "</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	  	rs.MoveNext
	Loop
	rs.close
	
	response.write(outline)
	outline = ""
	%>
	<tr><td colspan="2">
      <p style="margin-top: 6"><b>Top Substitute Goalscorers</b></td></tr>
	<%	
	sql = "select top 5 player_id_spell1, surname, forename, initials, count(*) as count " 
	sql = sql & "from " & view & " a join match_player b1 on a.date = b1.date "
	sql = sql & " join match_goal b2 on a.date = b2.date and b1.player_id = b2.player_id "
	sql = sql & " join player c on b1.player_id = c.player_id "
	sql = sql & "where startpos = 0 "
	sql = sql & "  and last_game_year = 9999 "
	sql = sql & "group by player_id_spell1, surname, forename, initials "
	sql = sql & "order by count desc, surname "
	rs.open sql,conn,1,2
	Do While Not rs.EOF
		if not IsNull(rs.Fields("forename")) then 
			playername = trim(rs.Fields("forename")) & " " & trim(rs.Fields("surname")) 
	  	  else
			playername = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname"))
		end if		
		outline  = outline & "<tr><td>" & playername & "</td><td align=""right"">" & rs.Fields("count") & "</td></tr>"		
	  	rs.MoveNext
	Loop
	rs.close
					
	outline  = outline & "</table>"
	response.write(outline)
	%>

<%
conn.close
%>	
</td></tr>

</table>
</center><br>

<!--#include file="base_code.htm"-->
</body>

</html>