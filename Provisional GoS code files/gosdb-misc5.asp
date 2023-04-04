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
#table3 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1}
#table4 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:2 padding-right:2; padding-top:1; padding-bottom:1}
#table5 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:2; padding-right:2; padding-top:1; padding-bottom:1}
#table6 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:2; padding-right:2; padding-top:1; padding-bottom:1}
#table7 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:2; padding-right:2; padding-top:1; padding-bottom:1}
-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline, dobcount, playercount, playername, age, workdate, workdiff
Dim selected_comp(3), competition, heading, tablehead, LFCvalue, compcode


Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%
%>

  <table border="0" cellspacing="0" style="border-collapse: collapse; margin:0 auto" cellpadding="0" width="980">
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
    <b>Report 5: Youngest and Oldest</b></p> 
    
    <%
    sql = "select 'A', count(distinct a.player_id) as playercount "
	sql = sql & "from player a join match_player b on a.player_id = b.player_id "
	sql = sql & "where spell = 1 "
	sql = sql & "and dob is not null "
	sql = sql & "union all "
	sql = sql & "select 'B', count(distinct a.player_id) "
	sql = sql & "from player a join match_player b on a.player_id = b.player_id " 
	sql = sql & "where spell = 1 "
	sql = sql & "order by 1 "
	rs.open sql,conn,1,2
	dobcount = rs.Fields("playercount")
    rs.MoveNext
    playercount = rs.Fields("playercount")
    rs.close
	%> 
    
    </td>
        
	<td width="260" valign="top"  align="justify">
	'<span style="font-size: 10px">Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
     
    </td>
    </tr>   
	</table>
	<center>
	
	<form style="font-size: 10px; padding: 0; margin: 0;" action="gosdb-misc5.asp" method="post" name="form1">
	
      
<%
competition = Request.Form("competition")

compcode = ""
tablehead = ""

select case competition
	case "FLG"
		selected_comp(1) = "selected"
		LFCvalue = "'F'"
		heading = "Football League"
		tablehead = " in the Football League"
	case "FAC"
		selected_comp(2) = "selected"
		LFCvalue = "'C'"
		compcode = "'FAC'"		
		heading = "FA Cup"
		tablehead = " in the FA Cup"
	case "CUP"
		selected_comp(3) = "selected"
		LFCvalue = "'C'"
		heading = "Any Cup Competition"
		tablehead = " in any Cup"
	case else
		selected_comp(0) = "selected"
		LFCvalue = "'L','F','C'"
		heading = "All Competitions"
 end select
 
 outline = "<select name=""competition"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ALL"" " & selected_comp(0) & ">All Competitions</option>"
 outline = outline & "<option value=""FLG"" " & selected_comp(1) & ">Football League</option>"
 outline = outline & "<option value=""FAC"" " & selected_comp(2) & ">FA Cup</option>" 
 outline = outline & "<option value=""CUP"" " & selected_comp(3) & ">Any Cup</option>" 
 outline = outline & "</select>"

 outline = outline & "<br><input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 9 0 0 0"" value=""Select and click to redisplay"" name=""B1""></p>"  
 outline = outline & "</form>"
 
 outline = outline & "<p style=""margin-top:6px; margin-bottom:6px; text-align:center; font-size:14px; color:#006E32""><b>" & heading & "</b></p>"
 
 response.write(outline)
 %>
	
	<p style="margin-top: 9; margin-bottom: 4">
    <span style="font-size: 11px;">Note: GoS-DB holds birth dates for <%response.write(dobcount)%> of <%response.write(playercount)%> players, almost all of the 
    missing ones being for some who appeared before 1920 and in the years close 
    to World War II. </span> </p> 
		
	<p style="margin-top: 0; margin-bottom: 0">
    <span style="font-size: 11px">This report does not include those players, so the rankings 
    shown here could be inaccurate.</span> </p> 
		
	<p style="margin:12 200; ">
    <font color="#008000">
    <span style="font-size: 11px; font-weight: 700">Scroll down for youngest and 
    oldest goalscorers</span></font></p> 
		
 	<table border="0" cellpadding="10" cellspacing="10" style="border-collapse: collapse; margin:0 0 6 0" bordercolor="#111111">
  	<tr>
    <td valign="top">
    <table id="table1" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
 	<tr><td colspan="5"><b>Youngest Players on Debut<%response.write(tablehead)%></b></td></tr>
    <tr><td><b>#</b></td><td><b>Player</b></td><td><b>Yr</b></td><td align="right">
      <b>Dy</b></td><td><b>First Match</b></td></tr>
	</tr>
      
	<%
	outline = ""
	sql = "select top 50 rank() over (order by datediff(day,dob,date)) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age "
	sql = sql & "from  player a join match_player b on a.player_id = b.player_id "
	sql = sql & "where date = (select min(b1.date) "
	sql = sql & "			   from player a1 "
	sql = sql & "			   join match_player b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  "
	sql = sql & "			   where a1.player_id_spell1 = a.player_id_spell1 "
	sql = sql & "  				 and LFC in (" & LFCvalue & ") "
	if compcode > "" then sql = sql & "	and compcode = (" & compcode & ") "
	sql = sql & "			   ) " 
	sql = sql & "and dob is not null "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		
		if not IsNull(rs.Fields("forename")) then 
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
		  else
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if

		workdiff = AgeSub(rs.Fields("dob"),rs.Fields("date"))
		age = split(workdiff,",")
	  
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """>" & playername & "</a></td>" 
		outline  = outline & "<td>" & age(0) & "</td>"
		outline  = outline & "<td>" & age(1) & "</td>"
		outline  = outline & "<td>" & Day(rs.Fields("date")) & " " & MonthName(Month(rs.Fields("date")),True) & " " & Year(rs.Fields("date")) & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)
	%>
	</table>
	</td>
	<td valign="top"> 
    <table id="table2" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
 	<tr><td colspan="5"><b>Oldest Players on Debut<%response.write(tablehead)%></b></td></tr>
    <tr>
      <td><b>#</b></td><td><b>Player</b></td><td><b>Yr</b></td><td align="right">
      <b>Dy</b></td><td><b>First Match</b></td>
    </tr>
	<%
	outline = ""
	sql = "select top 50 rank() over (order by datediff(day,dob,date) desc) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age "
	sql = sql & "from  player a join match_player b on a.player_id = b.player_id "
	sql = sql & "where date = (select min(b1.date) "
	sql = sql & "			   from player a1 "
	sql = sql & "			   join match_player b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  "
	sql = sql & "			   where a1.player_id_spell1 = a.player_id_spell1 "
	sql = sql & "  				 and LFC in (" & LFCvalue & ") "
	if compcode > "" then sql = sql & "	and compcode = (" & compcode & ") "
	sql = sql & "			   ) " 
	sql = sql & "and dob is not null "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		
		if not IsNull(rs.Fields("forename")) then 
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
		  else
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if

		workdiff = AgeSub(rs.Fields("dob"),rs.Fields("date"))
		age = split(workdiff,",")
	  
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """>" & playername & "</a></td>" 
		outline  = outline & "<td>" & age(0) & "</td>"
		outline  = outline & "<td>" & age(1) & "</td>"
		outline  = outline & "<td>" & Day(rs.Fields("date")) & " " & MonthName(Month(rs.Fields("date")),True) & " " & Year(rs.Fields("date")) & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)
	%>
	</table>
	</td>
	<td valign="top">
    <table id="table3" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
 	<tr><td colspan="5"><b>Oldest Players to Appear<%response.write(tablehead)%></b></td></tr>
    <tr>
      <td><b>#</b></td><td><b>Player</b></td><td><b>Yr</b></td><td align="right">
      <b>Dy</b></td><td><b>Last Match</b></td>
    </tr>
	<%
	outline = ""
	sql = "select top 50 rank() over (order by datediff(day,dob,date) desc) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age "
	sql = sql & "from  player a join match_player b on a.player_id = b.player_id "
	sql = sql & "where date = (select max(b1.date) "
	sql = sql & "			   from player a1 "
	sql = sql & "			   join match_player b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  "
	sql = sql & "			   where a1.player_id_spell1 = a.player_id_spell1 "
	sql = sql & "  				 and LFC in (" & LFCvalue & ") "
	if compcode > "" then sql = sql & "	and compcode = (" & compcode & ") "
	sql = sql & "			   ) " 
	sql = sql & "and dob is not null "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		
		if not IsNull(rs.Fields("forename")) then 
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
		  else
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if

		workdiff = AgeSub(rs.Fields("dob"),rs.Fields("date"))
		age = split(workdiff,",")
	  
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """>" & playername & "</a></td>" 
		outline  = outline & "<td>" & age(0) & "</td>"
		outline  = outline & "<td>" & age(1) & "</td>"
		outline  = outline & "<td>" & Day(rs.Fields("date")) & " " & MonthName(Month(rs.Fields("date")),True) & " " & Year(rs.Fields("date")) & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)
	%>
	</table>
	</td></tr>
	</table>
	
		
 	<table border="0" cellpadding="10" cellspacing="10" style="border-collapse: collapse; margin:6px auto" bordercolor="#111111"  width="980">
  	<tr>
    <td valign="top">
    <table id="table4" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
 	<tr><td colspan="5"><b>Youngest Scorers<%response.write(tablehead)%></b></td></tr>
    <tr><td><b>#</b></td><td><b>Player</b></td><td><b>Yr</b></td><td align="right">
      <b>Dy</b></td><td><b>First Goal</b></td></tr>
	</tr>
      
	<%
	outline = ""
	sql = "select top 50 rank() over (order by datediff(day,dob,date)) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age, count(*) as goalcount "
	sql = sql & "from  player a join match_goal b on a.player_id = b.player_id "
	sql = sql & "where date = (select min(b1.date) "
	sql = sql & "			   from player a1 "
	sql = sql & "			   join match_goal b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  "
	sql = sql & "			   where a1.player_id_spell1 = a.player_id_spell1 "
	sql = sql & "  				 and LFC in (" & LFCvalue & ") "
	if compcode > "" then sql = sql & "	and compcode = (" & compcode & ") "
	sql = sql & "			   ) " 
	sql = sql & "and dob is not null "
	sql = sql & "group by player_id_spell1, surname, forename, initials, dob, date "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		
		if not IsNull(rs.Fields("forename")) then 
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
		  else
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if

		workdiff = AgeSub(rs.Fields("dob"),rs.Fields("date"))
		age = split(workdiff,",")
	  
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """>" & playername & "</a>"
		if rs.Fields("goalcount") > 1 then outline = outline & " * " 
		outline  = outline & "</td><td>" & age(0) & "</td>"
		outline  = outline & "<td>" & age(1) & "</td>"
		outline  = outline & "<td>" & Day(rs.Fields("date")) & " " & MonthName(Month(rs.Fields("date")),True) & " " & Year(rs.Fields("date")) & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)
	%>
	</table>
	</td>
	<td valign="top">
    <table id="table5" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
 	<tr><td colspan="5" style="white-space:normal"><b>Youngest Scorers on Debut<%response.write(tablehead)%></b></td></tr>
    <tr><td><b>#</b></td><td><b>Player</b></td><td><b>Yr</b></td><td align="right">
      <b>Dy</b></td><td><b>First Goal</b></td></tr>
	</tr>
      
	<%
	outline = ""
	sql = "select top 50 rank() over (order by datediff(day,dob,date)) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age, count(*) as goalcount "
	sql = sql & "from  player a join match_goal b on a.player_id = b.player_id "
	sql = sql & "where date = (select min(b1.date) "
	sql = sql & "			   from player a1 "
	sql = sql & "			   join match_player b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  "
	sql = sql & "			   where a1.player_id_spell1 = a.player_id_spell1 "
	sql = sql & "  				 and LFC in (" & LFCvalue & ") "
	if compcode > "" then sql = sql & "	and compcode = (" & compcode & ") "
	sql = sql & "			   ) " 
	sql = sql & "and dob is not null "
	sql = sql & "group by player_id_spell1, surname, forename, initials, dob, date "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		
		if not IsNull(rs.Fields("forename")) then 
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
		  else
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if

		workdiff = AgeSub(rs.Fields("dob"),rs.Fields("date"))
		age = split(workdiff,",")
		
		
	  
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """>" & playername & "</a>"
		if rs.Fields("goalcount") > 1 then outline = outline & " * " 
		outline  = outline & "</td><td>" & age(0) & "</td>"
		outline  = outline & "<td>" & age(1) & "</td>"
		outline  = outline & "<td>" & Day(rs.Fields("date")) & " " & MonthName(Month(rs.Fields("date")),True) & " " & Year(rs.Fields("date")) & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)
	%>
	</table>
	</td>
	<td valign="top"> 
    <table id="table6" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
 	<tr><td colspan="5" style="white-space:normal"><b>Youngest Scorers on Starting Debut<%response.write(tablehead)%></b></td></tr>
    <tr>
      <td><b>#</b></td><td><b>Player</b></td><td><b>Yr</b></td><td align="right">
      <b>Dy</b></td><td><b>First Goal</b></td>
    </tr>
	<%
	outline = ""
	sql = "select top 50 rank() over (order by datediff(day,dob,date)) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age, count(*) as goalcount "
	sql = sql & "from  player a join match_goal b on a.player_id = b.player_id "
	sql = sql & "where date = (select min(b1.date) "
	sql = sql & "			   from player a1 "
	sql = sql & "			   join match_player b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  "	
	sql = sql & "			   where a1.player_id_spell1 = a.player_id_spell1 "
	sql = sql & "  				 and LFC in (" & LFCvalue & ") "
	sql = sql & "			     and startpos > 0 "
	if compcode > "" then sql = sql & "	and compcode = (" & compcode & ") "
	sql = sql & "			   ) " 
	sql = sql & "and dob is not null "
	sql = sql & "group by player_id_spell1, surname, forename, initials, dob, date "

	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		
		if not IsNull(rs.Fields("forename")) then 
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
		  else
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if

		workdiff = AgeSub(rs.Fields("dob"),rs.Fields("date"))
		age = split(workdiff,",")
	  
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """>" & playername & "</a>"
		if rs.Fields("goalcount") > 1 then outline = outline & " * " 
		outline  = outline & "</td><td>" & age(0) & "</td>"
		outline  = outline & "<td>" & age(1) & "</td>"
		outline  = outline & "<td>" & Day(rs.Fields("date")) & " " & MonthName(Month(rs.Fields("date")),True) & " " & Year(rs.Fields("date")) & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)
	%>
	</table>
	</td>
	<td valign="top"> 
    <table id="table7" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
 	<tr><td colspan="5"><b>Oldest Scorers<%response.write(tablehead)%></b></td></tr>
    <tr>
      <td><b>#</b></td><td><b>Player</b></td><td><b>Yr</b></td><td align="right">
      <b>Dy</b></td><td><b>Last Goal</b></td>
    </tr>
	<%
	outline = ""
	sql = "select top 50 rank() over (order by datediff(day,dob,date) desc) as rank, a.player_id_spell1, surname, forename, initials, dob, date, datediff(day,dob,date) as age, count(*) as goalcount "
	sql = sql & "from  player a join match_goal b on a.player_id = b.player_id "
	sql = sql & "where date = (select max(b1.date) "
	sql = sql & "			   from player a1 "
	sql = sql & "			   join match_goal b1 on a1.player_id = b1.player_id join v_match_all c1 on b1.date = c1.date  "
	sql = sql & "			   where a1.player_id_spell1 = a.player_id_spell1 "
	sql = sql & "  				 and LFC in (" & LFCvalue & ") "
	if compcode > "" then sql = sql & "	and compcode = (" & compcode & ") "
	sql = sql & "			   ) " 
	sql = sql & "and dob is not null "
	sql = sql & "group by player_id_spell1, surname, forename, initials, dob, date "
	
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		
		if not IsNull(rs.Fields("forename")) then 
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
		  else
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if

		workdiff = AgeSub(rs.Fields("dob"),rs.Fields("date"))
		age = split(workdiff,",")
	  
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """>" & playername & "</a>"
		if rs.Fields("goalcount") > 1 then outline = outline & " * " 
		outline  = outline & "</td><td>" & age(0) & "</td>"
		outline  = outline & "<td>" & age(1) & "</td>"
		outline  = outline & "<td>" & Day(rs.Fields("date")) & " " & MonthName(Month(rs.Fields("date")),True) & " " & Year(rs.Fields("date")) & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)
	%>
	</table>

	</td></tr>
	</table>
	<p style="margin: 0 0 12 0; font-size: 11px;">* scored more than once in the game</p>


<%
conn.close

Function AgeSub(date1,date2)
	workdate = day(date1) & "-" & month(date1) & "-" & year(date2)
	if datediff("d",workdate,date2) < 0 then workdate = day(date1) & "-" & month(date1) & "-" & year(date2) - 1
	AgeSub = datediff("yyyy",date1,workdate) & "," & datediff("d",workdate,date2)
End Function
%>	


<!--#include file="base_code.htm"-->
</body>

</html>