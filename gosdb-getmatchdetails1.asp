<%
response.expires = -1

' Set Session Locale to English(UK) to solve date differences when swapping from Fasthosts to Tsohost
' [Tsohost defaults to 1033, which is English(US)]
Session.LCID=2057

querydate = Request.QueryString("q")
displaydate = FormatDateTime(querydate,1)
mnths = DateDiff("m",querydate,Date)
yrs = Int(mnths/12)
remmnths = mnths - 12*yrs
select case yrs
 case 0
 	yrs = ""
 case 1	
  	yrs = yrs & " year "
 case else 
 	yrs = yrs & " years "
 end select
 select case remmnths
 case 0
 	remmnths = ""
 case 1	
  	remmnths = remmnths & " month "
 case else 
 	remmnths = remmnths & " months "
 end select 
work1 = split(displaydate," ")
outline = "<b>" & WeekDayName(WeekDay(querydate)) & ", " & work1(0) & " " & left(work1(1),3) & " " & work1(2) & "</b>"
if yrs > "" then outline = outline & " [" & yrs & remmnths & " ago]"
outline = outline & "<br>" 

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rsgoals = Server.CreateObject("ADODB.Recordset")
Set rslineup = Server.CreateObject("ADODB.Recordset")
Set rssubbedsubs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_read.inc"--><%

'Get further match details
sql = "select competition, subcomp, homeaway, opposition, opposition_qual, goalsfor, goalsagainst, aet_ind, pensfor, pensagainst, notes, ground_name "
sql = sql & "from match a"
sql = sql & " join competition b on a.compcode = b.compcode "
sql = sql & " join opposition on opposition = name_then "
sql = sql & " left outer join venue on name_then = club_name_then and date between first_game and last_game "
sql = sql & "where date = '" & querydate & "' "

rs.open sql,conn,1,2

outline = outline & "<b>" & rs.Fields("competition")
if rs.Fields("subcomp") > "" then outline = outline & " " & rs.Fields("subcomp")
outline = outline & "<br>"
if rs.Fields("homeaway") = "H" then 
	outline = outline & "Argyle " & rs.Fields("goalsfor") & " " & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " " & rs.Fields("goalsagainst")
	if ucase(rs.Fields("aet_ind")) = "Y" then outline = outline & "</b> after extra time"
	if not isnull(rs.Fields("pensfor")) then
		outline = outline & "<br><b>Penalties:</b> Argyle " & rs.Fields("pensfor") & " " & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " " & rs.Fields("pensagainst")
	end if
	venue = "Home Park"
	if querydate = "1961-03-18" then venue = "Plainmoor"
  else
	outline = outline & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " " & rs.Fields("goalsagainst") & " " & "Argyle " & rs.Fields("goalsfor")
	if ucase(rs.Fields("aet_ind")) = "Y" then outline = outline & "</b> after extra time"
	if not isnull(rs.Fields("pensfor")) then
		outline = outline & "<br><b>Penalties:</b> " & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " " & rs.Fields("pensagainst") & " Argyle " & rs.Fields("pensfor")
	end if
	venue = rs.Fields("ground_name")
end if
outline = outline & "</b><br>"
outline = outline & "<b>Venue: </b>" & venue & "<br>"
if rs.Fields("notes") > "" then outline = outline & "<b>Note: </b>" &rs.Fields("notes") & "<br>"

rs.close
					
'Get team line-up
sql = "select b.surname as start_surname, b.forename as start_forename, b.initials as start_initials, a.startpos, c.player_id as sub_playerid, d.surname as sub_surname, d.forename as sub_forename, d.initials as sub_initials "
sql = sql & "from match_player a join player b on a.player_id = b.player_id "
sql = sql & " left outer join match_player c on (a.date = c.date and a.replaced_by = c.player_id) left outer join player d on c.player_id = d.player_id  "
sql = sql & "where a.date = '" & request.querystring("q") & "' "
sql = sql & "  and a.startpos > 0 "
sql = sql & "order by a.startpos "

rslineup.open sql,conn,1,2
		
outline  = outline & "<b>Team:</b> "
Do While Not rslineup.EOF
	if rslineup.Fields("start_forename") > "" then
		outline = outline & trim(rslineup.Fields("start_forename")) & " " & trim(rslineup.Fields("start_surname"))
		elseif rslineup.Fields("start_initials") > "" then
			outline = outline & left(rslineup.Fields("start_initials"),1) & ". " & trim(rslineup.Fields("start_surname"))
		else outline = outline & trim(rslineup.Fields("start_surname"))  	 
	end if

			sub_surname = trim(rslineup.Fields("sub_surname")) 
			sub_forename = trim(rslineup.Fields("sub_forename")) 
			sub_initials = left(rslineup.Fields("sub_initials"),1)
			sub_playerid = rslineup.Fields("sub_playerid")
			sub_brackets = "" 
			 
			Do While sub_surname > ""
				if sub_forename > "" then
					outline = outline & " (" & sub_forename & " " & sub_surname
					elseif sub_initials > "" then
						outline = outline & " (" & sub_initials & ". " & sub_surname
					else outline = outline & " (" & sub_surname
				end if
				sub_brackets = sub_brackets &  ")"
				sub_surname = "" 
				
				'Find if thus sub has himself been subbed
				sql = "select c.player_id as sub_playerid, d.surname as sub_surname, d.forename as sub_forename, d.initials as sub_initials "
				sql = sql & "from match_player a join player b on a.player_id = b.player_id "
				sql = sql & " left outer join match_player c on (a.date = c.date and a.replaced_by = c.player_id) left outer join player d on c.player_id = d.player_id  "
				sql = sql & "where a.date = '" & request.querystring("q") & "' "
				sql = sql & "  and a.player_id = " & sub_playerid

				rssubbedsubs.open sql,conn,1,2
				if not rssubbedsubs.EOF then
					sub_surname = trim(rssubbedsubs.Fields("sub_surname")) 
					sub_forename = trim(rssubbedsubs.Fields("sub_forename")) 
					sub_initials = left(rslineup.Fields("sub_initials"),1)
					sub_playerid = rssubbedsubs.Fields("sub_playerid")
				end if
				rssubbedsubs.close
			Loop
			
			outline = outline & sub_brackets & ", "

	rslineup.MoveNext
Loop
outline = left(outline,len(outline)-2)  & "."  'remove last comma and space and add full stop

rslineup.close

'Check for unattributable subs

'Get older-type subs (no indication of who subbed for)
sql = "select surname, forename, initials, startpos "
sql = sql & "from match_player a join player b on a.player_id = b.player_id "
sql = sql & "where date = '" & request.querystring("q") & "' "
sql = sql & "  and startpos = 0 and sub_replacing is null "
sql = sql & "order by startpos "

rslineup.open sql,conn,1,2

if rslineup.RecordCount > 0 then
	if rslineup.RecordCount = 1 then 
		outline  = outline & " Sub: "
		else outline  = outline & " Subs: "
	end if
	Do While Not rslineup.EOF
		if rslineup.Fields("forename") > "" then
			outline = outline & trim(rslineup.Fields("forename")) & " " & trim(rslineup.Fields("surname"))
		  	elseif rslineup.Fields("initials") > "" then
			 	outline = outline & left(rslineup.Fields("initials"),1) & ". " & trim(rslineup.Fields("surname"))
		  	else outline = outline & trim(rslineup.Fields("surname"))  	 
		end if
		outline = outline & ", "
		rslineup.MoveNext
	Loop
	outline = left(outline,len(outline)-2)  & "."  'remove last comma and space and add full stop
end if
				
rslineup.close

'Get goal scorers
sql = "select surname, forename, initials, time, seqno, a.player_id "
sql = sql & "from match_goal a join player b on a.player_id = b.player_id "
sql = sql & "where date = '" & request.querystring("q") & "' "
sql = sql & "order by seqno "

rsgoals.open sql,conn,1,2
		
if rsgoals.RecordCount > 0 then  
		
	outline  = outline & "<br><b>Goals:</b> "
	
	Dim goals(9,2)
	
	Do While Not rsgoals.EOF
		if rsgoals.Fields("forename") > "" then
				scorer = trim(rsgoals.Fields("forename")) & " " & trim(rsgoals.Fields("surname"))
			elseif rsgoals.Fields("initials") > "" then
			 	scorer = left(rsgoals.Fields("initials"),1) & ". " & trim(rsgoals.Fields("surname"))
		  	else 
		  		scorer = trim(rsgoals.Fields("surname"))
		end if
		
		for i = 0 to 9
			if goals(i,0) = rsgoals.Fields("player_id") then
				goals(i,2) = goals(i,2) + 1
				exit for
			  elseif goals(i,0) = "" then
			  	goals(i,0) = rsgoals.Fields("player_id")
			  	goals(i,1) = scorer
			  	goals(i,2) = 1
			  	exit for
			 end if
		next
		if i = 9 then outline = outline & "A problem has occurred, please report this to Steve."
		rsgoals.MoveNext
	Loop
	
	for i = 0 to 9
		if goals(i,0) = "" then exit for
		outline = outline & goals(i,1) 
		if goals(i,2) > 1 then outline = outline & " (" & goals(i,2) & ")"
	  	outline = outline & ", "
	next
	
	outline = left(outline,len(outline)-2)  	'remove last comma and space
			
end if
				
rsgoals.close
		
response.write(outline)
%>