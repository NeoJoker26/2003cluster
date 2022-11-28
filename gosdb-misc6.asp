<%@ Language=VBScript %> 
<% Option Explicit %>

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
#table1 td {border: 1px solid #c0c0c0; text-align:right; margin: 0; white-space:nowrap; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1}
-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline, heading, warning, bycolumn, bestworst, games, homeaway, homeawayclause1, homeawayclause2, orderby, selected_bycolumn(1), selected_bestworst(1), selected_games(20), selected_ha(2), thisseason, thisseasoncount
Dim competition, subheading, tablename, lastcomp, selected_comp(1)

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

competition = Request.Form("competition")
if competition = "" then competition = "league"	'not important for the next piece of code, but needed for later
 select case competition
	case "all" 
		selected_comp(1) = "selected"
		tablename = "match"
		heading = "First " & games & " Matches in all Competitions"
		subheading = "(Cup games allocated league-style points)"
 	case else
		selected_comp(0) = "selected"
		tablename = "[v_match_FL-39]"
		heading = "First " & games & " League Matches"
 end select

	sql = "select years, count(*) as count "
	sql = sql & "from " & tablename & " a join season on date >= date_start and date <= date_end "
	sql = sql & "where season_no = ( "
	sql = sql & " select max (season_no) from season "
	sql = sql & ") "
	sql = sql & homeawayclause1
	sql = sql & "group by years "
	rs.open sql,conn,1,2
	thisseasoncount = rs.Fields("count")
	thisseason = rs.Fields("years")
	rs.close
	
games = Request.Form("games")
lastcomp = Request.querystring("lastcomp")
if games = "" or lastcomp <> competition then games = thisseasoncount
if games < 2 then games = 2
if games > 20 then games = 20

warning = ""
if CInt(games) > CInt(thisseasoncount) then warning = "Warning: " & thisseason & "'s counts are for " & thisseasoncount & " matches"

homeaway = Request.Form("homeaway")
 select case homeaway
	case "ho" 
		selected_ha(1) = "selected"
		homeawayclause1 = "and homeaway = 'H' "
		homeawayclause2 = "where homeaway = 'H' " 
		heading = heading & " (Home)"
	case "ao" 
		selected_ha(2) = "selected"
		homeawayclause1 = "and homeaway = 'A' " 
		homeawayclause2 = "where homeaway = 'A' " 
		heading = heading & " (Away)"
 	case else
		selected_ha(0) = "selected"
		homeawayclause1 = ""
		homeawayclause2 = ""
 end select

	
bycolumn = Request.Form("bycolumn")
bestworst = Request.Form("bestworst")


select case bycolumn
	case "diff"
		orderby = "goaldiff"
		selected_bycolumn(1) = "selected"
	case else
		orderby = "modernpoints"
		selected_bycolumn(0) = "selected"
 end select
 
 select case bestworst
	case "worst"
		selected_bestworst(1) = "selected"
	case else
		orderby = orderby & " desc"
		selected_bestworst(0) = "selected"
	end select
 
selected_games(games) = "selected"

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
    <b>Report 6:  Best and Worst Starts to a Season</b></p> 
       
    </td>
        
	<td width="260" valign="top"  align="justify">
	'<span style="font-size: 10px">Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
     
    </td>
    </tr>   
	</table>
	    
	<table id="table2" border="0" cellpadding="0" cellspacing="0" width="380" style="border-collapse: collapse">
	
	<form style="font-size: 10px; padding: 0; margin: 0;" action="gosdb-misc6.asp?lastcomp=<%response.write(competition)%>" method="post" name="form1">
	     
<%
 outline = "<tr><td>Choose display order:</td>"
 outline = outline & "<td><select name=""bestworst"" style=""font-size: 10px"">"
 outline = outline & "<option value=""best"" " & selected_bestworst(0) & ">Best first</option>"
 outline = outline & "<option value=""worst"" " & selected_bestworst(1) & ">Worst first</option>"  
 outline = outline & "</select></td></tr>"
 
 outline = "<tr><td>Choose competition:</td>"
 outline = outline & "<td><select name=""competition"" style=""font-size: 10px"">"
 outline = outline & "<option value=""league"" " & selected_comp(0) & ">League Matches</option>"
 outline = outline & "<option value=""all"" " & selected_comp(1) & ">All Matches"  
 outline = outline & "</select></td></tr>"
 
 outline = outline & "<tr><td>Based on Modern Points (3 for a win applied to every season) or Goal Difference:</td>" 
 outline = outline & "<td><select name=""bycolumn"" style=""font-size: 10px"">"
 outline = outline & "<option value=""modpoints"" " & selected_bycolumn(0) & ">Modern Points</option>"
 outline = outline & "<option value=""diff"" " & selected_bycolumn(1) & ">Goal Difference</option>"  
 outline = outline & "</select></td></tr>"
 
 outline = outline & "<tr><td>Choose length of start (2-20 games; defaults to latest game if between 2 and 20):</td>"
 outline = outline & "<td><select name=""games"" style=""font-size: 10px"">"
 outline = outline & "<option value=""2"" " & selected_games(2) & ">2</option>"
 outline = outline & "<option value=""3"" " & selected_games(3) & ">3</option>"
 outline = outline & "<option value=""4"" " & selected_games(4) & ">4</option>"
 outline = outline & "<option value=""5"" " & selected_games(5) & ">5</option>"
 outline = outline & "<option value=""6"" " & selected_games(6) & ">6</option>" 
 outline = outline & "<option value=""7"" " & selected_games(7) & ">7</option>"
 outline = outline & "<option value=""8"" " & selected_games(8) & ">8</option>"
 outline = outline & "<option value=""9"" " & selected_games(9) & ">9</option>"
 outline = outline & "<option value=""10"" " & selected_games(10) & ">10</option>"
 outline = outline & "<option value=""11"" " & selected_games(11) & ">11</option>"
 outline = outline & "<option value=""12"" " & selected_games(12) & ">12</option>"
 outline = outline & "<option value=""13"" " & selected_games(13) & ">13</option>"
 outline = outline & "<option value=""14"" " & selected_games(14) & ">14</option>"
 outline = outline & "<option value=""15"" " & selected_games(15) & ">15</option>"
 outline = outline & "<option value=""16"" " & selected_games(16) & ">16</option>"
 outline = outline & "<option value=""17"" " & selected_games(17) & ">17</option>"
 outline = outline & "<option value=""18"" " & selected_games(18) & ">18</option>"
 outline = outline & "<option value=""19"" " & selected_games(19) & ">19</option>"
 outline = outline & "<option value=""20"" " & selected_games(20) & ">20</option>" 
 outline = outline & "</select></td></tr>"

 outline = outline & "<tr><td>Choose venue:</td>"
 outline = outline & "<td><select name=""homeaway"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ha"" " & selected_ha(0) & ">Home & Away</option>" 
 outline = outline & "<option value=""ho"" " & selected_ha(1) & ">Home only</option>"
 outline = outline & "<option value=""ao"" " & selected_ha(2) & ">Away only</option>"
 outline = outline & "</select></td></tr>"
 
 outline = outline & "<tr><td colspan=""2"" align=""center""><input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 9 0 0 0"" value=""Select options and click to redisplay"" name=""B1""></p>" 
 outline = outline & "</td></tr>"
 response.write(outline)
 %>
 
 </form>
 </table>
 
 <%
  response.write("<p style=""margin-top:18px; margin-bottom:0; text-align:center; font-size:12px; color:#006E32""><b>" & heading & "</b></p>")
  if subheading > "" then response.write("<p style=""margin-top:6px; margin-bottom:12px; text-align:center; font-size:11px;""><b>" & subheading & "</b></p>")
  if warning > "" then response.write("<p style=""margin-top:12px; margin-bottom:0; text-align:center; font-size:11px; color:#CC3300;""><b>" & warning & "</b></p>")

 %>
	
    <table id="table1" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; margin-top:12px;">
    <tr>
      <td><b>#</b></td>
      <td><b>Season</b></td>
      <td><b>W</b></td>
      <td><b>D</b></td>
      <td><b>L</b></td>
      <td><b>F</b></td>
      <td><b>A</b></td>
      <td><b>Goal<br>Diff</b></td>
      <td><b>Mod'n<br>Points</b></td>
      <td><b>Final<br>Pos'n</b></td>
    </tr>
	<%
	outline = ""
	sql = "WITH CTE1 as "
	sql = sql & "(select ROW_NUMBER() over (PARTITION by years order by date) as game, "
	sql = sql & "date, years, goalsfor, goalsagainst, "
	sql = sql & "case when goalsfor > goalsagainst then 1 else 0 end as wins, "
	sql = sql & "case when goalsfor = goalsagainst then 1 else 0 end as draws, "
	sql = sql & "case when goalsfor < goalsagainst then 1 else 0 end as defeats, "
	sql = sql & "case when goalsfor > goalsagainst then 3 "
	sql = sql & "	 when goalsfor = goalsagainst then 1 "
	sql = sql & "	 else 0 "
	sql = sql & "	 end as modernpoints, "
	sql = sql & "endpos, promrel "
	sql = sql & "from " & tablename & " a "
	sql = sql & " join season e on a.date >= e.date_start and a.date <= e.date_end "
	sql = sql & homeawayclause2
	sql = sql & "), "
	sql = sql & "CTE2 as "
	sql = sql & "(select years, sum(goalsfor) as goalsfor, sum(goalsagainst) as goalsagainst, "
	sql = sql & "			  sum(goalsfor) - sum(goalsagainst) as goaldiff, "
	sql = sql & "			  sum(wins) as wins, sum(draws) as draws, sum(defeats) as defeats, "
	sql = sql & "			  sum(modernpoints) as modernpoints, endpos, promrel "
	sql = sql & "from CTE1 "
	sql = sql & "where game <= " & games
	sql = sql & " group by years, endpos, promrel"
	sql = sql & ") "
	sql = sql & "select rank() over (order by " & orderby & ") as rank, years, goalsfor, goalsagainst, "
	sql = sql & "			  goaldiff, "
	sql = sql & "			  wins, draws, defeats, "
	sql = sql & "			  modernpoints, endpos, promrel "
	sql = sql & "from CTE2 "
	sql = sql & "order by rank, years "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		outline  = outline & "<tr"
		if rs.Fields("years") = thisseason then outline = outline & " style=""background-color:c8e0c7"""
		outline  = outline & ">"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td>" & rs.Fields("years") & "</td>"
		outline  = outline & "<td>" & rs.Fields("wins") & "</td>"
		outline  = outline & "<td>" & rs.Fields("draws") & "</td>"
		outline  = outline & "<td>" & rs.Fields("defeats") & "</td>"
		outline  = outline & "<td>" & rs.Fields("goalsfor") & "</td>"
		outline  = outline & "<td>" & rs.Fields("goalsagainst") & "</td>"
		outline  = outline & "<td>" & rs.Fields("goaldiff") & "</td>"
		outline  = outline & "<td>" & rs.Fields("modernpoints") & "</td>"
		outline  = outline & "<td>" & rs.Fields("endpos")
		
		Select case rs.Fields("promrel")
			case "P" 
				outline  = outline & " <img src=""images/promote.gif"" border=""0"">"
			case "R" 
				outline  = outline & " <img src=""images/relegate.gif"" border=""0"">"
			case else
				outline  = outline & " <img src=""images/dummy.gif"" border=""0"" width=""11"" height=""1"">"
		End Select
		
		outline  = outline & "</td></tr>"
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)

conn.close
%>	
	
</table>
</center><br>

<!--#include file="base_code.htm"-->
</body>

</html>