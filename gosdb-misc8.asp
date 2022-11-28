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
Dim conn,sql,rs, outline, heading1, heading2, bycolumn, bestworst, yearcount, homeaway, homeawayclause1, homeawayclause2, orderby, selected_bycolumn(3), selected_bestworst(1)
Dim selected_yearcount(20), selected_ha(2), yearrange, work, style1, style2, style3, style4

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

	
yearcount = Request.Form("yearcount")
if yearcount = "" then yearcount = 1


homeaway = Request.Form("homeaway")
select case homeaway
	case "ho" 
		selected_ha(1) = "selected"
		homeawayclause1 = "and homeaway = 'H' "
		homeawayclause2 = "where homeaway = 'H' " 
		heading2 = "for Home League Matches only"
	case "ao" 
		selected_ha(2) = "selected"
		homeawayclause1 = "and homeaway = 'A' " 
		homeawayclause2 = "where homeaway = 'A' " 
		heading2 = "for Away League Matches only"
 	case else
		selected_ha(0) = "selected"
		homeawayclause1 = ""
		homeawayclause2 = ""
		heading2 = "for all League Matches"
end select

bycolumn = Request.Form("bycolumn")
bestworst = Request.Form("bestworst")

style1 = ""
style2 = ""
style3 = ""
style4 = ""

select case bycolumn
	case "diff"
		orderby = "goaldiff"
		selected_bycolumn(1) = "selected"
		style3 = "style=""background-color:f0f0f0""" 
	case "wins"
		orderby = "wins"
		selected_bycolumn(2) = "selected"
		style1 = "style=""background-color:f0f0f0""" 
	case "defeats"
		orderby = "defeats"
		selected_bycolumn(3) = "selected"
		style2 = "style=""background-color:f0f0f0""" 
	case else
		orderby = "modernpoints"
		selected_bycolumn(0) = "selected"
		style4 = "style=""background-color:f0f0f0""" 
end select
 
select case bestworst
	case "worst"
		selected_bestworst(1) = "selected"
	case else
		orderby = orderby & " desc"
		selected_bestworst(0) = "selected"
end select
 
selected_yearcount(yearcount) = "selected"
if yearcount = 1 then 
	heading1 = "Figures accumulated over a Single Calendar Year"
  else
  	heading1 = "Figures accumulated over " & yearcount & " Calendar Years" 
end if 
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
    <b>Report 8: League Results in a Calendar Year</b></p> 
          
    </td>
        
	<td width="260" valign="top"  align="justify">
	'<span style="font-size: 10px">Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
     
    </td>
    </tr>   
	</table>
	    
	<table id="table2" border="0" cellpadding="0" cellspacing="0" width="320" style="border-collapse: collapse">
	
	<form style="font-size: 10px; padding: 0; margin: 0;" action="gosdb-misc8.asp" method="post" name="form1">
	     
<%
 outline = "<tr><td>Choose display order:</td>"
 outline = outline & "<td><select name=""bestworst"" style=""font-size: 10px"">"
 outline = outline & "<option value=""best"" " & selected_bestworst(0) & ">Greatest first</option>"
 outline = outline & "<option value=""worst"" " & selected_bestworst(1) & ">Least first</option>"  
 outline = outline & "</select></td></tr>"
 
 outline = outline & "<tr><td>Sort criteria (Modern Points means 3 for a win for every season):</td>" 
 outline = outline & "<td><select name=""bycolumn"" style=""font-size: 10px"">"
 outline = outline & "<option value=""modpoints"" " & selected_bycolumn(0) & ">Modern Points</option>"
 outline = outline & "<option value=""diff"" " & selected_bycolumn(1) & ">Goal Difference</option>"
  outline = outline & "<option value=""wins"" " & selected_bycolumn(2) & ">Wins</option>"  
 outline = outline & "<option value=""defeats"" " & selected_bycolumn(3) & ">Defeats</option>"  
 outline = outline & "</select></td></tr>"
 
 outline = outline & "<tr><td>Calendar year range (leave at one for a single year view):</td>"
 outline = outline & "<td><select name=""yearcount"" style=""font-size: 10px"">"
 outline = outline & "<option value=""1"" " & selected_yearcount(1) & ">1</option>"
 outline = outline & "<option value=""2"" " & selected_yearcount(2) & ">2</option>"
 outline = outline & "<option value=""3"" " & selected_yearcount(3) & ">3</option>"
 outline = outline & "<option value=""4"" " & selected_yearcount(4) & ">4</option>" 
 outline = outline & "<option value=""5"" " & selected_yearcount(5) & ">5</option>"
 outline = outline & "<option value=""6"" " & selected_yearcount(6) & ">6</option>"
 outline = outline & "<option value=""7"" " & selected_yearcount(7) & ">7</option>"
 outline = outline & "<option value=""8"" " & selected_yearcount(8) & ">8</option>"
 outline = outline & "<option value=""9"" " & selected_yearcount(9) & ">9</option>"
 outline = outline & "<option value=""10"" " & selected_yearcount(10) & ">10</option>"
 
 outline = outline & "</select></td></tr>"

 outline = outline & "<tr><td>Venue:</td>"

 outline = outline & "<td><select name=""homeaway"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ha"" " & selected_ha(0) & ">Home & Away</option>" 
 outline = outline & "<option value=""ho"" " & selected_ha(1) & ">Home only</option>"
 outline = outline & "<option value=""ao"" " & selected_ha(2) & ">Away only</option>"
 outline = outline & "</select></td></tr>"
 
 outline = outline & "<tr><td colspan=""2"" align=""center""><input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 9 0 12 0"" value=""Select options and click to redisplay"" name=""B1""></p>" 
 outline = outline & "</td></tr>"
 response.write(outline)
 %>
 
 </form>
 </table>
 
 <%
  response.write("<p style=""margin-top:0; margin-bottom:6; text-align:center; font-size:12px; color:#006E32""><b>" & heading1 & "</b></p>")
  response.write("<p style=""margin-top:0; margin-bottom:12; text-align:center; font-size:12px; color:#006E32""><b>" & heading2 & "</b></p>")
 %>
 
 <p style="margin-top:0; margin-bottom:12; text-align:center; font-size:11px">
 Note: for a fair comparison, only years that included Football<br>League fixtures each side of a summer break have been included.</b></p>

	
    <table id="table1" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
    <tr>
      <td><b>#</b></td>
      <td><b>Season</b></td>
      <td><b>P</b></td>
      <td <%response.write(style1)%>><b>W</b></td>
      <td><b>D</b></td>
      <td <%response.write(style2)%>><b>L</b></td>
      <td><b>F</b></td>
      <td><b>A</b></td>
      <td <%response.write(style3)%>><b>Goal<br>Diff</b></td>
      <td <%response.write(style4)%>><b>Mod'n<br>Points</b></td>
    </tr>
	<%
	outline = ""
	sql = "WITH CTE1 as "
	sql = sql & "(select distinct cast(year(date) as varchar) + '-' + cast(year(date) + " & yearcount-1 & " as varchar) as years, year(date) as startyear, year(date) + " & yearcount-1 & " as endyear "
	sql = sql & "from [v_match_FL-39] a "
	sql = sql & "where year(date) <> 1920 " 
	sql = sql & "  and year(date) <> 1946 " 
	sql = sql & "  and year(date) + " & yearcount-1 & " < " & year(Date)
	sql = sql & "  and year(date) + " & yearcount-1 & " not between 1939 and 1946 "
	sql = sql & "), "
	sql = sql & "CTE2 as "
	sql = sql & "(select years, 1 as played, goalsfor, goalsagainst, "
	sql = sql & "case when goalsfor > goalsagainst then 1 else 0 end as wins, "
	sql = sql & "case when goalsfor = goalsagainst then 1 else 0 end as draws, "
	sql = sql & "case when goalsfor < goalsagainst then 1 else 0 end as defeats, "
	sql = sql & "case when goalsfor > goalsagainst then 3 "
	sql = sql & "	 when goalsfor = goalsagainst then 1 "
	sql = sql & "	 else 0 "
	sql = sql & "	 end as modernpoints "
	sql = sql & "from [v_match_FL-39] join CTE1 on year(date) between startyear and endyear "
	sql = sql & homeawayclause2
	sql = sql & "), "
	sql = sql & "CTE3 as "
	sql = sql & "(select years, sum(played) as played, sum(goalsfor) as goalsfor, sum(goalsagainst) as goalsagainst, "
	sql = sql & "			  sum(goalsfor) - sum(goalsagainst) as goaldiff, "
	sql = sql & "			  sum(wins) as wins, sum(draws) as draws, sum(defeats) as defeats, "
	sql = sql & "			  sum(modernpoints) as modernpoints "
	sql = sql & "from CTE2 "
	sql = sql & " group by years "
	sql = sql & ") "
	sql = sql & "select rank() over (order by " & orderby & ") as rank, years, played, goalsfor, goalsagainst, "
	sql = sql & "			  goaldiff, "
	sql = sql & "			  wins, draws, defeats, "
	sql = sql & "			  modernpoints "
	sql = sql & "from CTE3 "
	sql = sql & "order by rank, years "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		yearrange = rs.Fields("years")
		work = split(yearrange,"-")
		if work(0) = work(1) then yearrange = work(0)	'one year only 
		outline  = outline & "<td>" & yearrange & "</td>"
		outline  = outline & "<td>" & rs.Fields("played") & "</td>"
		outline  = outline & "<td " & style1 & ">" & rs.Fields("wins") & "</td>"
		outline  = outline & "<td>" & rs.Fields("draws") & "</td>"
		outline  = outline & "<td " & style2 & ">" & rs.Fields("defeats") & "</td>"
		outline  = outline & "<td>" & rs.Fields("goalsfor") & "</td>"
		outline  = outline & "<td>" & rs.Fields("goalsagainst") & "</td>"
		outline  = outline & "<td " & style3 & ">" & rs.Fields("goaldiff") & "</td>"
		outline  = outline & "<td " & style4 & ">" & rs.Fields("modernpoints") & "</td>"
	
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