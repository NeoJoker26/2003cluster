<%@ Language=VBScript %> 
<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>GoS-DB Miscellaneous Report</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
#table1 th, #table2 th, #table3 th {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding: 6px 4px}
#table1 td, #table2 td, #table3 td {border: 1px solid #c0c0c0; text-align:right; margin: 0; white-space:nowrap; padding: 2px 4px}
-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline, playername, temp, startdate, enddate 

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%
%>

  <table border="0" cellspacing="0" style="border-collapse: collapse; margin:0 auto; padding:15px" width="980">
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
	<p style="margin-top:12; margin-bottom:0; text-align:center; font-size:18px; color:#006E32">MISCELLANEOUS REPORTS</p>     
	<p style="margin-top:6; margin-bottom:0; text-align:center; font-size:13px">
    <b>Report 12: Consecutive Appearances</b></p>     
    </td>
        
	<td width="260" valign="top"  align="justify">
	<span style="font-size: 10px">'Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
     
    </td>
    </tr>   
	</table>
	<center>
	
		
 	<table border="0" style="border-collapse: collapse; margin:10px 0 6 0" bordercolor="#111111">
  	<tr>
    <td valign="top" style="padding-right:20px">
    <table id="table1" style="border-collapse: collapse">
 	<tr><th colspan="4"><b>Consecutive Starting Appearances<br>in All Competitions (50 and above)</b></th></tr>
    <tr><td><b>#</b></td><td style="text-align:left"><b>Player</b></td><td><b>App</b></td><td style="text-align:left"><b>Dates</b></td></tr>
	</tr>
      
	<%
	outline = ""
	sql = "select rank() over (order by consec_count desc) as rank, a.player_id, surname, forename, consec_count, start_date, end_date "
	sql = sql & "from consecutive_appears a join player b on a.player_id = b.player_id "
	sql = sql & "where consec_count >= 50 "
	sql = sql & "order by consec_count desc, start_date "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		
		if not IsNull(rs.Fields("forename")) then 
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
		  else
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if
	  
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td style=""text-align:left""><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id") & """>" & playername & "</a></td>" 
		outline  = outline & "<td>" & rs.Fields("consec_count") & "</td>"
		temp = split(FormatDateTime(rs.Fields("start_date"),1)," ")
		startdate = temp(0) & " " & left(temp(1),3) & " " & temp(2)
		temp = split(FormatDateTime(rs.Fields("end_date"),1)," ")
		enddate = temp(0) & " " & left(temp(1),3) & " " & temp(2)
		outline  = outline & "<td style=""text-align:left"">" & startdate & "<br>to " & enddate & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)
	%>
	</table>
	</td>
	<td valign="top" style="padding-right:20px"> 
    <table id="table2" style="border-collapse: collapse">
    <tr><th colspan="4"><b>Consecutive Starting Appearances<br>in League Matches (50 and above)</b></th></tr>
    <tr><td><b>#</b></td><td style="text-align:left"><b>Player</b></td><td><b>App</b></td><td style="text-align:left"><b>Dates</b></td></tr>
	<%
	outline = ""
	sql = "select rank() over (order by l_consec_count desc) as rank, a.player_id, surname, forename, l_consec_count, l_start_date, l_end_date "
	sql = sql & "from consecutive_appears a join player b on a.player_id = b.player_id "
	sql = sql & "where l_consec_count >= 50 "
	sql = sql & "order by l_consec_count desc, l_start_date "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		
		if not IsNull(rs.Fields("forename")) then 
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
		  else
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if
	  
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td style=""text-align:left""><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id") & """>" & playername & "</a></td>" 
		outline  = outline & "<td>" & rs.Fields("l_consec_count") & "</td>"
		temp = split(FormatDateTime(rs.Fields("l_start_date"),1)," ")
		startdate = temp(0) & " " & left(temp(1),3) & " " & temp(2)
		temp = split(FormatDateTime(rs.Fields("l_end_date"),1)," ")
		enddate = temp(0) & " " & left(temp(1),3) & " " & temp(2)
		outline  = outline & "<td style=""text-align:left"">" & startdate & "<br>to " & enddate & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)
	%>
	</table>
	</td>
	<td valign="top">
	<div style="width:310px">
	<p class="style1bold" style="margin: 0 0 4px 0;">A Note About Pat Jones</p>
	<p class="style1" style="margin: 0 0 30px 0;">The left-back's astonishing record of 279 consecutive Football League appearances 
	for Argyle would have been eclipsed by an even more impressive feat -
	291 in all competitions instead of 175 - had he not been injured for one game, 
	the FA Cup third round replay at Wolves in January 1950.</p>
	</div>
 	<table id="table3" style="border-collapse: collapse">
	<tr><th colspan="5"><b>Consecutive Goal-Scoring Appearances<br>in All Competitions (3 and above)</b></th></tr>
    <tr><td style="text-align:left"><b>Player</b></td><td><b>App</b></td><td><b>Gls</b></td><td style="text-align:left"><b>Dates</b></td></tr>
	<%
	outline = ""
	sql = "select rank() over (order by goal_consec_count desc) as rank, a.player_id, surname, forename, goal_consec_count, goal_count, goal_start_date, goal_end_date "
	sql = sql & "from consecutive_appears a join player b on a.player_id = b.player_id "
	sql = sql & "where goal_consec_count >= 3 "
	sql = sql & "order by goal_consec_count desc, goal_start_date "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		
		if not IsNull(rs.Fields("forename")) then 
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
		  else
		  	playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if
	  
		outline  = outline & "<tr>"
		outline  = outline & "<td style=""text-align:left""><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id") & """>" & playername & "</a></td>" 
		outline  = outline & "<td>" & rs.Fields("goal_consec_count") & "</td>"
		outline  = outline & "<td>" & rs.Fields("goal_count") & "</td>"
		temp = split(FormatDateTime(rs.Fields("goal_start_date"),1)," ")
		startdate = temp(0) & " " & left(temp(1),3) & " " & temp(2)
		temp = split(FormatDateTime(rs.Fields("goal_end_date"),1)," ")
		enddate = temp(0) & " " & left(temp(1),3) & " " & temp(2)
		outline  = outline & "<td style=""text-align:left"">" & startdate & "<br>to " & enddate & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	conn.close
	response.write(outline)
	%>
	
	</table>

	</td></tr>
	</table>
	

<!--#include file="base_code.htm"-->
</body>

</html>