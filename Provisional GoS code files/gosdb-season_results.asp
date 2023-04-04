<%
Dim conn, sql, rs, rsgoals, rslineup, rssubbedsubs, goalscorers(10,1)
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rsgoals = Server.CreateObject("ADODB.Recordset")
Set rslineup = Server.CreateObject("ADODB.Recordset")
Set rssubbedsubs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

sort = Request.QueryString("sort")
years = Request.QueryString("years")
part3940 = Request.QueryString("part")

output = "<p style=""font-size: 14px; font-weight:700; color:green; margin: 0 0 6px; text-align: center;"">"
if part3940 = "D2" then
    output = output & " SEASON RESULTS (Abandoned Division 2 only)</b></p>"
  	orderby = "order by date "
  	limitdate = " and date <= '1939-09-02' " 
  elseif part3940 = "SWRL" then
    output = output & " SEASON RESULTS (South West Regional League only)</b></p>"
  	orderby = "order by date " 
  	limitdate = " and date > '1939-09-02' "
  else
  	output = output & " SEASON RESULTS</b></p>"
   	orderby = "order by date " 
   	limitdate = "" 
end if

output = output & "<p class=""style1"" style=""margin: 0; text-align: center;"">Click on "																	
output = output & "<img border=""0"" src=""images/more.png""> for more match details."  
output = output & "</p>"

output = output & "<table class=""restable"">"  
output = output & "<tr><th id=""sort1"" class=""sort"">Date<br><img src=""images/sort.gif"" border=""0"" hspace=""1"" vspace=""2""></th>"
output = output & "<th id=""sort2"" class=""sort"">Comp<br><img src=""images/sort.gif"" border=""0"" hspace=""1"" vspace=""2""></th>"
output = output & "<th>H/A</th>"
output = output & "<th>Opposition</th>"
output = output & "<th>Attend</th>"
output = output & "<th>Score</th>"
output = output & "<th align=""left"">Scorers</th>"
output = output & "<th align=""left"">Team</th></tr>"

if sort = 2 then orderby = "order by shortcomp, date " 
    
sql = "select competition, shortcomp, subcomp, LFC, date, "
sql = sql & "left(datename(weekday,date),3) + '<br>' + left(datename(month,date),3) + ' ' + datename(day,date) as form_date, "
sql = sql & "homeaway, opposition, opposition_qual, goalsfor, goalsagainst, attendance, a.notes "
sql = sql & "from v_match_all a join season b on date between date_start and date_end "
sql = sql & "where years = '" & years & "'" 
sql = sql & limitdate
sql = sql & orderby

rs.open sql,conn,1,2

tagno = 1

Do While Not rs.EOF

	datehold = rs.Fields("date")
	
	att = rs.Fields("attendance")	
	if len(att) > 3 then att = left(rs.Fields("attendance"),len(rs.Fields("attendance"))-3) & "," & right(rs.Fields("attendance"),3)

	output  = output & "<tr>"
	output  = output & "<td nowrap class=""match""><a href=""gosdb-match.asp?date=" & datehold & """><img style=""vertical-align: text-top"" src=""images/more.png"">" & rs.Fields("form_date") & "</a></td>"
	if rs.Fields("LFC") = "C"  then
		output  = output & "<td nowrap" & cellclass & " title=""Cup: " & rs.Fields("competition") & """>" & rs.Fields("shortcomp")
		if not isnull(rs.Fields("subcomp")) then output =  output & "<br>/" & rs.Fields("subcomp")
		output = output  & "</td>"
	  else
		output  = output & "<td nowrap" & cellclass & " title=""League: " & rs.Fields("competition") & """>" & rs.Fields("shortcomp") & "</td>" 
	end if
	if rs.Fields("homeaway") = "H" then output  = output & "<td align=""center"">Home</td>"
	if rs.Fields("homeaway") = "A" then output  = output & "<td align=""center"">Away</td>"
	if rs.Fields("homeaway") = "N" then output  = output & "<td align=""center"">Neutral<br>Venue</td>"
	output  = output & "<td " & cellclass & ">" & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & "</td>"
	output  = output & "<td " & cellclass & " align=""right"">" & att & "</td>"
	output  = output & "<td " & cellclass & " align=""center"">" & rs.Fields("goalsfor") & "-" & rs.Fields("goalsagainst") & "</td>"
	
		'Get team line-up
		teamline1 = ""
		teamline2 = ""
		teamsurnames = ""
		agecount = 0
		totage = 0
		
		sql = "select b.surname as start_surname, b.forename as start_forename, b.initials as start_initials, a.startpos, b.dob as startdob, c.player_id as sub_playerid, d.surname as sub_surname, d.forename as sub_forename, d.initials as sub_initials, d.dob as subdob "
		sql = sql & "from match_player a join player b on a.player_id = b.player_id "
		sql = sql & " left outer join match_player c on (a.date = c.date and a.replaced_by = c.player_id) left outer join player d on c.player_id = d.player_id  "
		sql = sql & "where a.date = '" & datehold & "' "
		sql = sql & "  and a.startpos > 0 "
		sql = sql & "order by a.startpos "

		rslineup.open sql,conn,1,2

		Do While Not rslineup.EOF
			if rslineup.Fields("start_forename") > "" then
				teamline1 = teamline1 & trim(rslineup.Fields("start_forename")) & " " & trim(rslineup.Fields("start_surname"))
				teamline2 = teamline2 & trim(rslineup.Fields("start_forename")) & " " & trim(rslineup.Fields("start_surname"))
				elseif rslineup.Fields("start_initials") > "" then
					teamline1 = teamline1 & left(rslineup.Fields("start_initials"),1) & ". " & trim(rslineup.Fields("start_surname"))
					teamline2 = teamline2 & left(rslineup.Fields("start_initials"),1) & ". " & trim(rslineup.Fields("start_surname"))
				else 
					teamline1 = teamline1 & trim(rslineup.Fields("start_surname"))
					teamline2 = teamline2 & trim(rslineup.Fields("start_surname"))  	 
			end if

			startage = ""
			subage = ""
			
			if not isnull(rslineup.Fields("startdob")) then
				agecount = agecount + 1
				totage = totage + DateDiff("d",rslineup.Fields("startdob"),datehold)
				startage = DateDiff("yyyy",rslineup.Fields("startdob"),datehold)
				if datevalue(year(datehold) & "-" & month(rslineup.Fields("startdob")) & "-" & day(rslineup.Fields("startdob"))) > datevalue(datehold) then startage = startage - 1 'haven't reached birthday this year yet
				startage = " <span style=""color:#606060"">" & startage & "</span>"
			end if
			if not isnull(rslineup.Fields("subdob")) then 
				subage = DateDiff("yyyy",rslineup.Fields("subdob"),datehold)
				if datevalue(year(datehold) & "-" & month(rslineup.Fields("subdob")) & " - " & day(rslineup.Fields("subdob"))) > datevalue(datehold) then subage = subage - 1 'haven't reached birthday this year yet
				subage = " <span style=""color:#606060"">" & subage & "</span>"
			end if

			teamline2 = teamline2 & startage
			teamsurnames = teamsurnames & trim(rslineup.Fields("start_surname")) & " "
			
			sub_surname = trim(rslineup.Fields("sub_surname")) 
			sub_forename = trim(rslineup.Fields("sub_forename")) 
			sub_initials = left(rslineup.Fields("sub_initials"),1)
			sub_playerid = rslineup.Fields("sub_playerid")
			 
			Do While sub_surname > ""
				if sub_forename > "" then
					teamline1 = teamline1 & " (" & sub_forename & " " & sub_surname & ")"
					teamline2 = teamline2 & " (" & sub_forename & " " & sub_surname & subage & ")"
					elseif sub_initials > "" then
						teamline1 = teamline1 & " (" & sub_initials & ". " & sub_surname & ")"
						teamline2 = teamline2 & " (" & sub_initials & ". " & sub_surname & subage & ")"
					else 
						teamline1 = teamline1 & " (" & sub_surname & ")"
						teamline2 = teamline2 & " (" & sub_surname & subage & ")"
				end if
				teamsurnames = teamsurnames & sub_surname & " "
				sub_surname = "" 
				
				'Find if this sub has himself been subbed
				sql = "select c.player_id as sub_playerid, d.surname as sub_surname, d.forename as sub_forename, d.initials as sub_initials, d.dob as subdob "
				sql = sql & "from match_player a join player b on a.player_id = b.player_id "
				sql = sql & " left outer join match_player c on (a.date = c.date and a.replaced_by = c.player_id) left outer join player d on c.player_id = d.player_id  "
				sql = sql & "where a.date = '" & datehold & "' "
				sql = sql & "  and a.player_id = " & sub_playerid

				rssubbedsubs.open sql,conn,1,2
				if not rssubbedsubs.EOF then
					sub_surname = trim(rssubbedsubs.Fields("sub_surname")) 
					sub_forename = trim(rssubbedsubs.Fields("sub_forename")) 
					sub_initials = left(rssubbedsubs.Fields("sub_initials"),1)
					sub_playerid = rssubbedsubs.Fields("sub_playerid")
					subage = ""
					if not isnull(rssubbedsubs.Fields("subdob")) then 
						subage = DateDiff("yyyy",rssubbedsubs.Fields("subdob"),datehold)
						if datevalue(year(datehold) & "-" & month(rssubbedsubs.Fields("subdob")) & "-" & day(rssubbedsubs.Fields("subdob"))) > datevalue(datehold) then subage = subage - 1 'haven't reached birthday this year yet
						subage = " <span style=""color:#606060"">" & subage & "</span>"				
					end if
				end if
				rssubbedsubs.close
			Loop
			
			teamline1 = teamline1 & ", "
			teamline2 = teamline2 & ", "
			rslineup.MoveNext
		Loop
		teamline1 = left(teamline1,len(teamline1)-2)  & "."  'remove last comma and space and add full stop
		teamline2 = left(teamline2,len(teamline2)-2)  & "."  'remove last comma and space and add full stop

		rslineup.close

		'Check for unattributable subs

		'Get older-type subs (no indication of who subbed for)
		sql = "select surname, forename, initials, startpos, dob "
		sql = sql & "from match_player a join player b on a.player_id = b.player_id "
		sql = sql & "where date = '" & datehold & "' "
		sql = sql & "  and startpos = 0 and sub_replacing is null "
		sql = sql & "order by startpos "

		rslineup.open sql,conn,1,2
		
		if rslineup.RecordCount > 0 then
			if rslineup.RecordCount = 1 then 
				teamline1  = teamline1 & " Sub: "
				teamline2  = teamline2 & " Sub: "
			  else 
				teamline1  = teamline1 & " Subs: "
				teamline2  = teamline2 & " Subs: "
			end if
			Do While Not rslineup.EOF
				if rslineup.Fields("forename") > "" then
					teamline1 = teamline1 & trim(rslineup.Fields("forename")) & " " & trim(rslineup.Fields("surname"))
					teamline2 = teamline2 & trim(rslineup.Fields("forename")) & " " & trim(rslineup.Fields("surname"))
				  	elseif rslineup.Fields("initials") > "" then
					 	teamline1 = teamline1 & left(rslineup.Fields("initials"),1) & ". " & trim(rslineup.Fields("surname"))
					 	teamline2 = teamline21 & left(rslineup.Fields("initials"),1) & ". " & trim(rslineup.Fields("surname"))
				  	else 
				  		teamline1 = teamline1 & trim(rslineup.Fields("surname"))
				  		teamline2 = teamline2 & trim(rslineup.Fields("surname"))  	 
				end if
				teamsurnames = teamsurnames & trim(rslineup.Fields("surname")) & " "
				subage = ""
				if not isnull(rslineup.Fields("dob")) then
					subage = DateDiff("yyyy",rslineup.Fields("dob"),datehold)
					if datevalue(year(datehold) & "-" & month(rslineup.Fields("dob")) & "-" & day(rslineup.Fields("dob"))) > datevalue(datehold) then subage = subage - 1 'haven't reached birthday this year yet
					subage = " <span style=""color:#606060"">" & subage & "</span>"
				end if  
				teamline1 = teamline1 & ", "
				teamline2 = teamline2 & subage & ", "
				rslineup.MoveNext
			Loop
			teamline1 = left(teamline1,len(teamline1)-2)  & "."  'remove last comma and space and add full stop
			teamline2 = left(teamline2,len(teamline2)-2)  & "."  'remove last comma and space and add full stop

		end if
				
		rslineup.close
		
		if agecount = 11 then teamline2 = teamline2 & " Average age of starting XI: <span style=""color:#606060"">" & Round(totage/(11*365.25),2) & "</span>"
			
		teamline1 = teamline1 & " [<a href=""javascript:Toggle('x" & tagno & "','y" & tagno & "');"">Show ages</a>]"
		teamline2 = teamline2 & " [<a href=""javascript:Toggle('y" & tagno & "','x" & tagno & "');"">Hide ages</a>]"

	output  = output & "<td " & cellclass & ">"	
				
		'Get goal scorers
		
		for i = 0 to 9
			goalscorers(i,0) = " "
			goalscorers(i,1) = 0
		next
		
		sql = "select surname, forename, initials, time, seqno "
		sql = sql & "from match_goal a join player b on a.player_id = b.player_id "
		sql = sql & "where date = '" & datehold & "' "
		sql = sql & "order by seqno "

		rsgoals.open sql,conn,1,2
		
		if rsgoals.RecordCount > 0 then  
		
			Do While Not rsgoals.EOF
			
				'Check if goalscorer's surname appears more than once in team list. If not, no need to list forname in goalscorer list
				thisname = split(teamsurnames, trim(rsgoals.Fields("surname"))) 
				if Ubound(thisname) = 1 then
						goalscorer = trim(rsgoals.Fields("surname"))
					else
						if rsgoals.Fields("forename") > "" then
							goalscorer =  trim(rsgoals.Fields("forename")) & " " & trim(rsgoals.Fields("surname"))
							elseif rsgoals.Fields("initials") > "" then
					 			goalscorer = left(rsgoals.Fields("initials"),1) & ". " & trim(rsgoals.Fields("surname"))
		  					else goalscorer = trim(rsgoals.Fields("surname")) 
						end if
				end if
				for i = 0 to 10
					if goalscorers(i,0) = " " then 
						goalscorers(i,0) = goalscorer
						goalscorers(i,1) = 1
						exit for
					end if
					if goalscorers(i,0) = goalscorer then 
						goalscorers(i,1) = goalscorers(i,1) + 1
						exit for
					end if		
				next
				
				rsgoals.MoveNext
			Loop
			
			i = 0
			Do Until goalscorers(i,0) = " "
			
				output = output & goalscorers(i,0)
				if goalscorers(i,1) > 1 then
					output = output & " (" & goalscorers(i,1) & "), "
					else output = output & ", "
				end if
				i = i + 1
			Loop
			if i > 0 then output = left(output,len(output)-2)  	'remove last comma and space
			
		end if
				
		rsgoals.close
		
	output  = output & "</td>"
		
	output  = output & "<td " & cellclass & ">"
	output  = output & "<span id=""x" & tagno & """>" & teamline1 & "</span>" & "<span id=""y" & tagno & """ style=""display:none;"">" & teamline2 & "</span>"
	
	if not IsNull(rs.Fields("notes")) then
		output = output & "<p style=""margin:4 0 4 0; text-align:left; line-height:1.2;""><u>Note</u>: " & rs.Fields("notes") & "</p>"
	end if				
	output  = output & "</td>"
		
	output  = output & "</tr>"
	
	tagno = tagno + 1 
	
	rs.MoveNext
	
Loop


rs.close
conn.close


output = output & "</tbody></table>"


response.write(output)
%><%'="a" %>