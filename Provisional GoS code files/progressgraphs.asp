<%@ Language=VBScript %>
<% Option Explicit %>

<!DOCTYPE html>

<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Greens on Screen</title>

<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--

.thisdate {
	background-color: #e0f0e0;
	} 

.nohover a:hover { background-color: transparent; }

-->
</style>

<script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script>
<script>
$(document).ready(function(){

    $('select').on('change',function() {
        $(".clock").append('<img style="position:relative; left:3px; top:4px; border:0; margin:0; padding:0; height:14px; " src="images/loading.gif">');
    });

});
</script>
 
</head>

<body><!--#include file="top_code.htm"-->

</style>
<%

Dim id1, output, output1, output2, headline, headrow, baserow, players, match, target_date, target_matchno, target_season, latest_date, linkparm, arraybit, num, n0, n1, n2, result(60,15), n3, n4, n5, n6, n7, pos1, pos2, square1, square2, bkground
Dim subhold(5), subind, scorers, scorer, scorename, goaltime, numgoals, graphic, title1, title2, cardhold, matchno, homeno, season(50,70,2), contentlink(60,2), points(50,1), thhlt, ptsbar, sort, lowsquad, highsquad, step, validmatches, barheight, latestmatch, latestpoints, totpoints1, totpoints2, hometotattend, homecount, maint
Dim position(24,50), lpos, bordercolour, lastpointsheight, thispointsheight, thispointsborder, temp, teams_in_div, pos_promote, pos_promote_playoff, pos_relegate_playoff, pos_relegate, endpos, selected, posnote

Dim season_start, season_end, season_no, pointsforawin, currentseason, division, divisionpart, divisionparts, tier

Dim conn, sql, rs, rs1, sqlhold
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

target_date = Request.QueryString("date")
target_season = Request.Form("season")

output = "" 

sql = "select max(date) as maxdate "
sql = sql & "from match a join competition b on a.compcode = b.compcode "
sql = sql & "where lfc = 'F' " 
rs.open sql,conn,1,2
	latest_date = rs.Fields("maxdate")
rs.close

sql = "select season_no, years, division, date_start, date_end, endpos, tier, teams_in_div, pos_promote, pos_promote_playoff, pos_relegate_playoff, pos_relegate "
sql = sql & "from season " 
	if target_season > "" then
		sql = sql & "where season_no = " & target_season
	  elseif target_date > "" then
		sql = sql & "where date_start <= '" & target_date & "' "
		sql = sql & "  and date_end >= '" & target_date & "' "
	  else
		sql = sql & "where season_no = (select max(season_no) from season) "
end if

rs.open sql,conn,1,2

season_no = rs.Fields("season_no")
tier = rs.Fields("tier")
teams_in_div = rs.Fields("teams_in_div")
pos_promote = rs.Fields("pos_promote")
pos_promote_playoff = rs.Fields("pos_promote_playoff")
pos_relegate_playoff = rs.Fields("pos_relegate_playoff")
pos_relegate = rs.Fields("pos_relegate")
endpos = rs.Fields("endpos")
if isnull(rs.Fields("endpos")) then currentseason = "Y"

output = output & "<h3 style=""margin:20px 0 0"">" & rs.Fields("division") & " " & rs.Fields("years") & "</h3>"

season_start = rs.Fields("date_start")
season_end = rs.Fields("date_end")
rs.close

'Offer choice of another season if we have not come from a match page

if target_date = "" then

	output = output & "<form style=""font-size: 10px; padding: 0; margin: 0 auto;"" action=""progressgraphs.asp"" method=""post"" name=""form1"">"
	output = output & "or select another season: <select name=""season"" style=""margin: 3px 0 0; font-size: 10px"" onchange=""this.form.submit()"">"
	
	sql = "select distinct season_no, years "
	sql = sql & "from v_match_season " 
	sql = sql & "where year(date_start) >= 1920 and year(date_start) <> 1945 "
	sql = sql & "  and totpoints is not null " 
	sql = sql & "order by years "

	rs.open sql,conn,1,2
	Do While Not rs.EOF
		if rs.Fields("season_no") = season_no then
			selected = "selected"
		   else
		   	selected = ""
		end if
		output = output & "<option value=""" & rs.Fields("season_no") & """" & selected & ">" & rs.Fields("years") & "</option>"					
		rs.MoveNext
	Loop
	rs.close
	
	output = output & "</select><span class=""clock""></span>"  
	output = output & "</form>"

end if

output = output & "<div id=""graph1"">"

' Initialise league points array 

for n1 = 0 to UBound(points,1)
	for n2 = 0 to UBound(points,2)
		points(n1,n2) = 0   'initialise
    next
next

' Initialise league position array 

for n1 = 0 to UBound(position,1)
	for n2 = 0 to UBound(position,2)
		position(n1,n2) = 0   'initialise
    next
next


if currentseason = "Y" then
	sqlhold = "union all " 
	sqlhold = sqlhold & "select a.date, a.opposition, name_abbrev, lfc, homeaway, 'Home Park', NULL, NULL, NULL, NULL, competition, NULL "
	sqlhold = sqlhold & "from season_this a join competition b on a.compcode = b.compcode " 
	sqlhold = sqlhold & "join opposition c on a.opposition = c.name_then "
	sqlhold = sqlhold & "where not exists (select * from match c where c.date = a.date) "
	sqlhold = sqlhold & "  and homeaway = 'H' "
	sqlhold = sqlhold & "  and lfc = 'F' "
	sqlhold = sqlhold & "  and right(rtrim(shortcomp),2) <> 'PO' "		'Avoid the play-off results
	sqlhold = sqlhold & "  and date between '" & season_start & "' and '" & season_end & "'"
	sqlhold = sqlhold & "union all " 
	sqlhold = sqlhold & "select a.date, a.opposition, name_abbrev, lfc, homeaway, ground_name, NULL, NULL, NULL, NULL, competition, NULL "
	sqlhold = sqlhold & "from season_this a join competition b on a.compcode = b.compcode " 
	sqlhold = sqlhold & "join opposition c on a.opposition = c.name_then "
	sqlhold = sqlhold & "join venue d on a.opposition = d.club_name_then and a.date between d.first_game and d.last_game "   
	sqlhold = sqlhold & "where not exists (select * from match c where c.date = a.date) "
	sqlhold = sqlhold & "  and homeaway in ('A', 'N') "
	sqlhold = sqlhold & "  and lfc = 'F' "
	sqlhold = sqlhold & "  and right(rtrim(shortcomp),2) <> 'PO' "		'Avoid the play-off results
	sqlhold = sqlhold & "  and date between '" & season_start & "' and '" & season_end & "'"
  else
	sqlhold = ""
end if

sql = "select date, opposition, NULL as name_abbrev, lfc, homeaway, 'Home Park' as ground_name, goalsfor, goalsagainst, totpoints, position, competition, attendance "
sql = sql & "from v_match_season a " 
sql = sql & "where season_no = " & season_no
sql = sql & "  and homeaway = 'H' " 
sql = sql & "  and lfc = 'F' "
sql = sql & "  and right(rtrim(rtrim(shortcomp)),2) <> 'PO' "		'Avoid the play-off results
sql = sql & "union all "
sql = sql & "select date, opposition, NULL, lfc, homeaway, ground_name, goalsfor, goalsagainst, totpoints, position, competition, attendance "
sql = sql & "from v_match_season a "
sql = sql & "join venue b on a.opposition = b.club_name_then and a.date between b.first_game and b.last_game "  
sql = sql & "where season_no = " & season_no
sql = sql & "  and homeaway in ('A', 'N') "
sql = sql & "  and lfc = 'F' "
sql = sql & "  and right(rtrim(shortcomp),2) <> 'PO' "		'Avoid the play-off results
sql = sql & sqlhold
sql = sql & "order by lfc desc, date"
rs.open sql,conn,1,2

matchno = 0
homeno = 0

Do While Not rs.EOF

		if rs.Fields("homeaway") = "H" then homeno = homeno + 1
		matchno = matchno + 1
					
		 	result(matchno,0) = matchno
		 	result(matchno,1) = rs.Fields("homeaway")
			result(matchno,4) = rs.Fields("totpoints")  
			result(matchno,5) = rs.Fields("position")   
			result(matchno,9) = rs.Fields("attendance")	
		 	result(matchno,10) = rs.Fields("competition")
		 	result(matchno,10) = Replace(result(matchno,10),"'","\\\'")	'deal with an apostrophe in the competition name 
		 	result(matchno,11) = rs.Fields("date")
		 	result(matchno,13) = rs.Fields("ground_name") 
		 	result(matchno,13) = Replace(result(matchno,13),"'","")
			result(matchno,14) = rs.Fields("name_abbrev")
			
			if rs.Fields("date") = target_date then target_matchno = matchno
			
			sql = "with cte as (select row_number() over(order by date) as match, totpoints, position "
			sql = sql & "from v_match_season "    
			sql = sql & "where season_no = " & season_no - 1
			sql = sql & "	and lfc = 'F' "
			sql = sql & ") "
			sql = sql & "select totpoints, position "
			sql = sql & "from cte "
			sql = sql & "where match = " & matchno
			rs1.open sql,conn,1,2
			if not rs1.eof then
				result(matchno,6) = rs1.Fields("totpoints")	'total points at this stage last season
				if currentseason = "Y" then result(matchno,7) = rs1.Fields("position")	'position at this stage last season, but only for current season
			end if
			rs1.close
			
			sql = "with cte as (select row_number() over(order by date) as match, date, attendance "
			sql = sql & "from v_match_season "    
			sql = sql & "where season_no = " & season_no - 1
			sql = sql & "and homeaway = 'H' "
			sql = sql & ") "
			sql = sql & "select sum(attendance)/count(*) as aveatt "
			sql = sql & "from cte "
			sql = sql & "where match <= " & homeno
			rs1.open sql,conn,1,2
			
			if rs1.RecordCount > 0 then result(matchno,15) = rs1.Fields("aveatt") 
			rs1.close

		 	if rs.Fields("homeaway") = "A" then
		 		result(matchno,12) = rs.Fields("opposition") & " v Argyle"
				if not isnull(rs.Fields("goalsfor")) then result(matchno,8) = rs.Fields("opposition") & " " & rs.Fields("goalsagainst") & " Argyle " & rs.Fields("goalsfor")
		 	  else
		 		result(matchno,12) = "Argyle v " & rs.Fields("opposition")
		 		if not isnull(rs.Fields("goalsfor")) then result(matchno,8) = "Argyle " & rs.Fields("goalsfor") & " " & rs.Fields("opposition") & " " &rs.Fields("goalsagainst") 
		 	end if
 	 	 		
 		 	if rs.Fields("goalsfor") > rs.Fields("goalsagainst") then
		 		result(matchno,3) = 3 
		 	  elseif rs.Fields("goalsfor") = rs.Fields("goalsagainst") then
				result(matchno,3) = 1
			  else 
			  	result(matchno,3) = 0
		 	end if
						
			' Initialise content link
			contentlink(matchno,0) = ""
			contentlink(matchno,1) = "closed"
	
			if result(matchno,8) > "" then
			 	if target_date = "" then
			 		linkparm = "gosdb-match.asp?date=" & result(matchno,11)
			 	  else
			 	  	linkparm = "#"
			 	end if
				contentlink(matchno,0) = "<div class=""nohover""><a href=""" & linkparm & """>"
			  	contentlink(matchno,1) = "open"
			  else 
			  	contentlink(matchno,0) = "<div class=""nohover"">"
			end if
		

			'set this and last year's points totals
			
			if result(matchno,4) > "" then points(matchno,1) = result(matchno,4)
			if result(matchno,6) > "" then points(matchno,0) = result(matchno,6)
		
			
			'set this and last year's position (this year indicated by 2, last year by 1)
			
			lpos = result(matchno,5)
			if lpos > "" then position(lpos,matchno) = position(lpos,matchno) + 2 
			lpos = result(matchno,7)
			if lpos > "" then position(lpos,matchno) = position(lpos,matchno) + 1
			
		rs.MoveNext
	Loop
	rs.close
	 		
	 		
' Build Points display

output = output & "<h3 style=""margin:18px 0 0"">Points Progress</h3>"
output = output & "<p class=""style2"" style=""text-align: center; margin:0; "">Rest on a folder for basic details; click for the match page</p>"
	
validmatches = 0
homecount = 0
hometotattend = 0
headrow	= ""
baserow = ""

output = output & "<img style=""display: inline-block; vertical-align: bottom; margin-bottom: 16px;"" src=""images/gridaxis2.gif"">"

output = output & "<table><tr>"

For matchno = 1 to 50
	if result(matchno,1) = "H" then
		ptsbar = "homebar_"
	   else 
	    ptsbar = "awaybar_"
	end if
	if result(matchno,8) > "" and contentlink(matchno,1) = "open" then
		temp = split(FormatDateTime(result(matchno,11),1))
		if left(temp(0),1) = "0" then temp(0) = mid(temp(0),2,1)
		temp(1) = left(temp(1),3) 
		output1 = "<th title=""" & temp(0) & " " & temp(1) & " " & temp(2) & ": " & result(matchno,8) & """>" 
		output1 = output1 & matchno
		baserow = baserow + output1 + "</th>" 
		output2 = "<br>" & contentlink(matchno,0)
		if target_date = "" then output2 = output2 & "<img border=""0"" src=""images/" & contentlink(matchno,1) & ".gif"">"
		output2 = output2 & "<img border=""0"" vspace=""2"" src=""images/" & ptsbar & result(matchno,3) & ".gif"">"	
		output2 = output2 & "</a></div></th>"	
		headrow = headrow + output1 + output2
		validmatches = validmatches + 1
	 else
	 if result(matchno,12) > "" then
	 	temp = split(FormatDateTime(result(matchno,11),1))
		if left(temp(0),1) = "0" then temp(0) = mid(temp(0),2,1)
		temp(1) = left(temp(1),3) 
		output1 = "<th title=""" & temp(0) & " " & temp(1) & " " & temp(2) & ": " & result(matchno,12) & """>"
		output1 = output1 & matchno
		baserow = baserow + output1 + "</th>"
		output2 = "<br>" & contentlink(matchno,0)
		if target_date = "" then output2 = output2 & "<img border=""0"" src=""images/" & contentlink(matchno,1) & ".gif"">"	
		output2 = output2 & "<img border=""0"" vspace=""2"" src=""images/dummbar_0.gif"">"
		output2 = output2 & "</a></div></th>"	
		headrow = headrow + output1 + output2
		validmatches = validmatches + 1
	 end if
	end if
next

output = output & headrow
if currentseason = "Y" then output = output & "<th style=""border-top-style:none;border-right-style:none;""></th>"	'to accommodate a prediction column

output = output & "</tr><tr>"

for matchno = 1 to validmatches
 	bkground = ""
	
	if target_date < latest_date and matchno = target_matchno then bkground = " class=""thisdate"""
	
	if result(matchno,6) = 0 then
		lastpointsheight = 1
	  else
	  	lastpointsheight = result(matchno,6)*15/5
	end if
	if result(matchno,4) = 0 then
		thispointsheight = 1
		thispointsborder = 0
	  else
	  	thispointsheight = result(matchno,4)*15/5 - 2
	  	thispointsborder = 1
	end if
	
	output = output & "<td" & bkground & "><img src=""images/pointdumm1.gif"" width=""2px""><img src=""images/pointlast2.gif"" width=""5px"" height=""" & lastpointsheight & """ title=""Last season, points: " & result(matchno,6) & """>"
	if result(matchno,4) > "" then
		output = output & "<img src=""images/pointthis2.gif"" border=""" & thispointsborder & """ width=""4px"" height=""" & thispointsheight & """ title=""This season, points: " & result(matchno,4) & """><img src=""images/pointdumm1.gif"" width=""2px"">"
	  	latestmatch = matchno
	  	latestpoints = result(matchno,4)
	  else
	  	output = output & "<img src=""images/pointdumm1.gif"" width=""6px"">"
	end if	
	output = output & "</td>"
next

'Time to do predictions
if currentseason = "Y" and latestmatch > 9 then
	output = output + "<td style=""border-right-style:none;"">"
	totpoints1 = Round(latestpoints + (validmatches-latestmatch)*(result(latestmatch,4)-result(latestmatch-6,4))/6)
	output = output + "<img style=""margin-bottom:" & 15/5*totpoints1 - 3 & "px;"" src=""images/predict1.gif""></a>" 
	totpoints2 = Round(latestpoints + (validmatches-latestmatch)*latestpoints/latestmatch)
	output = output + "<img style=""margin-bottom:" & 15/5*totpoints2 - 3 & "px;"" src=""images/predict2.gif""></a>" 
	output = output + "</td>"
end if

output = output + "</tr>"

' Finish off with the final line

output = output + "<tr style=""border-top: 1px solid #c0c0c0"">" & baserow
if currentseason = "Y" then output = output + "<th></th>"		'because of the prediction column
output = output + "</tr>"

output = output & "</table>"

output = output & "<img style=""display: inline-block; vertical-align: bottom; margin-bottom: 16px;"" src=""images/gridaxis2.gif"">"

output = output & "<p class=""style2"" style=""text-align: center; margin:6px 0 2px""><b>Season:&nbsp;" 
output = output & "<img border=""0"" src=""images/thissquare1.gif"" align=""baseline""> this&nbsp;"
output = output & "<img border=""0"" src=""images/lastsquare1.gif"" align=""baseline""> last</b></p>"

if season_no <= 70 then output = output & "<p class=""style2"" style=""text-align: center; margin:0; "">[Two points for a win before 1981-82]"

if currentseason = "Y" and latestmatch > 9 then

	output = output & "<p class=""style1"" style=""text-align: center; margin: 0 0 2px;"">Don't take the <span style=""color:red"">red dots</span> (top right) too seriously!</p>"

	output = output & "<p class=""style1"" style=""text-align: center; margin: 0;""><img style=""vertical-align: text-top; margin-right: 2px"" src=""images/predict1.gif"">Based on latest 6 games: " & totpoints1 & " points"
	output = output & "<img style=""vertical-align: text-top;  margin-left: 20px; margin-right: 2px"" src=""images/predict2.gif"">Based on season so far: " & totpoints2 & " points"
	
end if

output = output & "</div>"


output = output & "<div id=""graph2"">"

if target_date > "" then output = output & headline	'don't include a second headline for dedicated graph page

output = output & "<table>"


' Build Position display

output = output & "<h3 style=""margin:18px 0 0"">Position Progress</h3>"
  
output = output + "<tr><th></th>"
	
if target_date > ""	then		'don't include a second full header row for dedicated graph page
	output = output + headrow & "</tr>"
  else
	output = output & baserow & "<th>End</th></tr>"
end if

posnote = ""

for n1 = 0 to teams_in_div
    
    bordercolour = ""
	if n1 = pos_promote then bordercolour = " style = ""border-bottom: 2px solid green;""" 
	if n1 = pos_promote_playoff then bordercolour = " style = ""border-bottom: 2px dashed green;"""
	if n1 = pos_relegate_playoff - 1 then bordercolour = " style = ""border-bottom: 2px dashed red;""" 
 	if n1 = pos_relegate - 1 then bordercolour = " style = ""border-bottom: 2px solid red;""" 
 
    
	for n2 = 0 to validmatches+1		'+1 to accomodate the end-of-season position
		if n2 = 0 then
			if n1 = 0 then 
				output = output + "<tr><td " & bordercolour & "><p style=""margin: 0 5px 0 3px; text-align: right; font-size: 9px;"">Pos</td>"
			else
				output = output + "<tr><td " & bordercolour & "><p style=""margin: 0 5px 0 3px; text-align: right; font-size: 9px;"">" & n1 & "</td>"
			end if	
		  elseif n2 = validmatches+1 and n1 = endpos then output = output + "<td" & bkground & bordercolour & "><img src=""images/thissquare1.gif""></td>"
		  else
			if n1 = 0 then
				output = output + "<td><img src=""images/nullsquare1.gif""></td>"
			  else
			  	bkground = ""
			  	if target_date < latest_date and n2 = target_matchno then bkground = " class=""thisdate"" "
				select case position(n1,n2)
				case 3
					output = output & "<td" & bkground & bordercolour & "><img src=""images/bothsquare1.gif""></td>"
				case 2
					output = output + "<td" & bkground & bordercolour & "><img src=""images/thissquare1.gif""></td>"
					if n2 = validmatches and n1 <> endpos then posnote = "Note: The end position is lower than match " & validmatches & " because other teams played after Argyle's final game." 
				case 1
					output = output + "<td" & bkground & bordercolour & "><img src=""images/lastsquare1.gif""></td>"
				case else
					output = output + "<td" & bkground & bordercolour & "><img src=""images/nullsquare1.gif""></td>"	
				end select
			end if
		end if
		if n2 = validmatches + 1 then
			output = output + "</tr>"
		end if
    next

next


' Finish off with the final line

output = output & "<td></td>"
output = output & baserow
output = output & "<th>End</th></tr>"

output = output & "</table>"

if currentseason = "Y" then
	output = output & "<p class=""style2"" style=""text-align: center; margin-top:6px 0""><b>Season:&nbsp;" 
	output = output & "<img border=""0"" src=""images/thissquare1.gif"" align=""baseline""> this&nbsp;"
	output = output & "<img border=""0"" src=""images/lastsquare1.gif"" align=""baseline""> last&nbsp;"
	output = output & "<img border=""0"" src=""images/bothsquare1.gif"" align=""baseline""> both</b></p>"
  elseif posnote > "" then
  	output = output & "<p class=""style1"" style=""text-align: center; margin-top:6px 0"">" & posnote & "</p>"
end if

output = output & "</div>"

response.write(output)
%><%'="a" %><%
%>

<!--#include file="base_code.htm"-->

</body>

</html>