<%
response.expires = -1

' Set Session Locale to English(UK) to solve date differences when swapping from Fasthosts to Tsohost
' [Tsohost defaults to 1033, which is English(US)]
Session.LCID=2057


Dim lastgamedate, teamfound, laststarters(10,1), thisstarters(10,1), intoteam, outofteam, report, reporttext

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
Set rsgoals = Server.CreateObject("ADODB.Recordset")
Set rslineup = Server.CreateObject("ADODB.Recordset")
Set rssubbedsubs = Server.CreateObject("ADODB.Recordset")
Set rsmiles = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_read.inc"--><%

	' Get match details
	
	matchdate = Request.QueryString("date")

	sql = "select season_no, division, a.date, opposition, opposition_qual, name_then_short, lfc, homeaway, goalsfor, goalsagainst, pensfor, pensagainst, totpoints, position, max_prog_image_no, notes, managers, "
	sql = sql & "competition, subcomp, attendance, visitors, non_playing_subs, opp_team, opp_goals, headline, report, report_published, report_acknowledge, "
	sql = sql & "d.ground_name as away_ground_name, d.ground_name_trad, d.round_trip, e.ground_name as home_ground_name, referee, ogscorers  "
	sql = sql & "from v_match_season a left outer join match_extra b on a.date = b.date join opposition c on a.opposition = c.name_then "
	sql = sql & "join v_manager_horiz on a.date between earliest_from_date and coalesce(latest_to_date,'9999-12-31') "
	sql = sql & "left outer join venue d on a.opposition = d.club_name_then and a.date between d.first_game and d.last_game " 
	sql = sql & "left outer join venue e on 'Plymouth Argyle' = e.club_name_then and a.date between e.first_game and e.last_game "
	sql = sql & "where a.date = '" & matchdate & "' "

	rs.open sql,conn,1,2
	
	if rs.RecordCount > 0 then 
					
		homeawayhold = rs.Fields("homeaway")
		clubhold = rs.Fields("opposition")
		goalsagainsthold = rs.Fields("goalsagainst")
		if isnull(rs.Fields("name_then_short")) then
			clubshorthold = rs.Fields("opposition")
		  else
		  	clubshorthold = rs.Fields("name_then_short")
		end if
		       	
       	output = "<div id=""matchdetails"">"
       	output = output & "<img class=""close"" style=""margin: 0 -12px 6px 6px; float: right; border: 0"" src=""images/close.png"">"
       	output = output & "<h1>Match of the Day: " & FormatDateTime(rs.Fields("date"),1) & "</h1>"'
       	output = output & "<h2>" & rs.Fields("competition")
       	if not IsNull(rs.Fields("subcomp")) then output = output & " " & trim(rs.Fields("subcomp"))
       	lfc = rs.Fields("lfc")
   		output = output & "</h2>"

       
		''''' Result '''''
       
        if rs.Fields("homeaway") = "H" then
			output = output & "<p class=""score"">Argyle &nbsp;" & rs.Fields("goalsfor") & " - " & rs.Fields("goalsagainst") & "&nbsp; " & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & "</p>"
 		  	if not isnull(rs.Fields("pensfor")) then
		  	   	output = output & "<p class=""penalties"">Penalties: Argyle " & rs.Fields("pensfor") & " - " & rs.Fields("pensagainst") & " " & rs.Fields("name_then_short") & "</p>"
			end if
		  else 
			output = output & "<p class=""score"">" & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " &nbsp;" & rs.Fields("goalsagainst") & " - " & rs.Fields("goalsfor") & "&nbsp; Argyle" & "</p>"
			if not isnull(rs.Fields("pensfor")) then
			   	output = output & "<p class=""penalties"">Penalties: " & rs.Fields("name_then_short") & " " & rs.Fields("pensagainst") & " - " & rs.Fields("pensfor") & " Argyle" & "</p>"	
			end if		   		  
		end if
		
		''''' Venue etc '''''

	    output = output & "<p style=""margin:9px 0 3px;"">"
	    output = output & "<span class=""bold"">Venue: </span><span class=""venue"">" 
    	if rs.Fields("homeaway") = "H" then
    		output = output & rs.Fields("home_ground_name")
       	  else 
    		output = output & rs.Fields("away_ground_name")
    		if not isnull(rs.Fields("ground_name_trad")) then output = output & " (aka " & rs.Fields("ground_name_trad") & ")"
		end if
		output = output & "</span>"
		
		if not IsNull(rs.Fields("attendance")) then 
			output = output & "<span class=""bold"">Attendance: </span>" 
			output = output & "<span class=""attendance"">" & FormatNumber(rs.Fields("attendance"),0,,-1) & "</span>"
		end if
		if not IsNull(rs.Fields("visitors")) then 
			output = output & "<span class=""bold"">Visitors: </span>" 
			output = output & "<span class=""visitors"">" & rs.Fields("visitors") & "</span>"
		end if
		if not IsNull(rs.Fields("referee")) then 
			output = output & "<span style=""white-space: nowrap""><span class=""bold"">Referee: </span>" 
			output = output & rs.Fields("referee") & "</span>"
		end if
	
       	output = output & "</p>"
       	
      	output = output & "<p style=""margin:3px 0 9px;"">"
      	if not IsNull(rs.Fields("totpoints")) then
	    	output = output & "<span class=""bold"">Total Points: </span><span class=""totpoints"">"
    		output = output & rs.Fields("totpoints")
			output = output & "</span>"
		end if
      	if not IsNull(rs.Fields("position")) then
	    	output = output & "<span class=""bold"">Position: </span>"
    		output = output & rs.Fields("position")
		end if	
	
       	output = output & "</p>"
       	    	
       	if not IsNull(rs.Fields("notes")) then output = output & "<p style=""margin:9px 0;""><span class=""bold"">Special note: </span>" & trim(rs.Fields("notes")) & "</p>"
 	
 		
		''''' Teams '''''

		outteam  = ""
		
		Call Getteam(matchdate)
		
		'Check the first cell of the team line-up array to see if a team has been found
		
		if teamfound = "y" then
				
			output_argyle = output_argyle & "<p class=""team""><span class=""style1bold"">ARGYLE:</span> " & outteam 
				
			if not IsNull(rs.Fields("non_playing_subs")) then
				nonplayingsubs = rs.Fields("non_playing_subs")
				if right(nonplayingsubs,1) <> "." then nonplayingsubs = nonplayingsubs & "."
				nonplayingsubs = "<p class=""font11px"" style=""margin:2px 0 0;""><span class=""style1bold"">Non-playing substitutes:</span> " & nonplayingsubs 
  				output_argyle = output_argyle & nonplayingsubs
	  		end if
	  	
	  		output_argyle = output_argyle & "</p>"

			if isdate(lastgamedate) then	'Check that it's not the first ever game
				'output_argyle = output_argyle & "<p class=""style1"" style=""margin:4px 0 0;""><b>Starting lineup changes <span style=""margin-left: 150px"">Manager: </span></b>" & rs.Fields("managers") & "</p>"
				output_argyle = output_argyle & "<p class=""style1"" style=""margin:4px 0 0;""><b>Starting lineup changes</b></p>"   
				output_argyle = output_argyle & "<p class=""style1"" style=""margin:0 0 6px;"">"
				if intoteam = "" then
					output_argyle = output_argyle & "None"
		   	  	  else
					output_argyle = output_argyle & "In: " & intoteam
					output_argyle = output_argyle & "<br>Out: " & outofteam
				end if		
	  			output_argyle = output_argyle & "</p>"
	  		end if
	  		
     	end if
     	     
		''''' Goals '''''

		outgoals = ""
		
		if not IsNull(rs.Fields("ogscorers")) then ogscorers = split(rs.Fields("ogscorers"),",")
		
		Call Getgoals(matchdate)
		
	  	if outgoals > "" then goals_argyle = "<span class=""style1"">Argyle:</span> " & outgoals & "<br>"

	  	if not IsNull(rs.Fields("opp_team")) then	
  			opposition = rs.Fields("opp_team")
  			opposition = replace(opposition,".",". ")
  			opposition = replace(opposition,"[y]","<img src=""images/yellowcard.gif"" hspace=""2"" height=""7"" width=""7"" align=""bottom"">")  
  			opposition = replace(opposition,"[yr]","<img src=""images/yelredcard.gif"" hspace=""2"" height=""8"" width=""10"" align=""bottom"">")
  			opposition = replace(opposition,"[r]","<img src=""images/redcard.gif"" hspace=""2"" height=""7"" width=""7"" align=""bottom"">")
  			opposition = "<span style=""white-space: nowrap;"">" & replace(opposition, ", " , ",</span> <span style=""white-space: nowrap;"">") & "</span>"
			output_opposition = output_opposition & "<p class=""team""><span class=""bold"">" & Ucase(clubshorthold) & "</span>: " & opposition & "</p>"
		end if 
		
		if not IsNull(rs.Fields("opp_goals")) then goals_opposition = "<span class=""style1"">" & clubshorthold & ":</span> " & rs.Fields("opp_goals") & "<br>"
		
		if rs.Fields("max_prog_image_no") > 0 then
			output = output & "<img style=""float:right; margin:0 0 6px 12px; max-width:280px; max-height:280px;"" src=""gosdb/photos/programmes/" & rs.Fields("date") & "-1.jpg"">"
		end if	
				
		if homeawayhold = "H" then
			output = output & output_argyle & output_opposition
			if goals_argyle > "" or goals_opposition > "" then 
				output = output & "<p class=""style1bold goals"">GOALS</p>"
				output = output & goals_argyle & goals_opposition
			end if
		  else
			output = output & output_opposition & output_argyle
			if goals_argyle > "" or goals_opposition > "" then 
				output = output & "<p class=""style1bold goals"">GOALS</p>"
				output = output & goals_opposition & goals_argyle		
			end if
		end if 
	       		    
		if not isnull(rs.Fields("report_published")) then 
		
			reporttext = ""
			
			if rs.Fields("headline") > " " then
				headline = replace(rs.Fields("headline"),"’","'") 
				headline = replace(headline,"|"," ") 
				reporttext = "<span class=""style1boldgrey"">" & headline
				if right(headline,1) = "!" or right(headline,1) = "?" then
					reporttext = reporttext & "</span> "
				   else
					reporttext = reporttext & ".</span> "
				end if   
			end if			
			reporttext = reporttext & replace(replace(rs.Fields("report"),"|p|","</p><p style=""margin:6px 0 0"">"),"’","'")
			if rs.Fields("report_acknowledge") = "A" then
				report = "<div id=""report""><p style=""margin:15px 0 0"">" & reporttext  
				report = report & "</p><p style=""margin:12px 0 0"">[Summary from <span style=""font-style:italic"">Plymouth Argyle, The Modern Era, A Complete Record</span> by Andy Riddle, with the author's kind permission.]</p></div>"
			  elseif rs.Fields("report_acknowledge") = "H" then
				report = "<div id=""report""><p style=""margin:15px 0 0"">" & reporttext  
				report = report & "</p><p style=""margin:12px 0 0"">[Extracted from <span style=""font-style:italic"">Harley Lawer's Argyle Classics</span>, with the author's kind permission.]</p></div>"
			  elseif rs.Fields("report_acknowledge") = "L" then
				report = "<div id=""report""><p style=""margin:15px 0 0"">" & reporttext  
				report = report & "</p><p style=""margin:12px 0 0"">[With thanks to Alec Hepburn for his account of the game.]</p></div>"
			  elseif rs.Fields("report_acknowledge") = "J" then
				report = "<div id=""report""><p style=""margin:15px 0 0"">" & reporttext  
				report = report & "</p><p style=""margin:12px 0 0"">[With thanks to John Eales for his account of the game.]</p></div>"
			  else
				report = "<div id=""report""><p style=""margin:15px 0 0"">" & reporttext & "</p></div>"
			end if 
		
			if rs.Fields("report_published") = "Y" then
				output2 = output2 & report
			  elseif phase = "review"	then
			    output2 = output2 & report & "<p style=""color:red; font-weight:bold;"">After review, don't forget to use your back button (possibly twice) to amend or publish.</p>"
    		end if
    		
    	  else
    	  	
			output2 = output2 & "<br><br><br>" 	'No report, so add a few lines to avoid formatting issues for the charts and table
    	
    	end if
		
		
		'This section displays the milestones and any match material
		
		output_milestones = "<div id=""milestones"">"
	
			Call Getmilestones(matchdate)
			
		output_milestones = output_milestones & "</div>"	'Finish off the milestones div
		
	
	
     	
		'Now put all the right-hand div components together
						
		output = output & output_milestones & output2 

      	output = output & "<p style=""margin: 12px auto""><a style=""font-weight:bold"" href=""gosdb-match.asp?date=" & rs.Fields("date") & """>Go to the Full Match Page</a></p></div>"	
      	
      	rs.close
      	
	  else
	
		output = "No match found"
	
	end if	
		
conn.close

response.write(output)


'-------------------------------------------------------------------------------------------------

Function Getteam(matchdate)

	Dim laststarters(10,1), thisstarters(10,1)
	i = 0
	j = 0

	'Get lineup for previous game (where a lineup exists)

		sql = "with cte1 as ( "
		sql = sql & "select row_number() over(order by a.date) as gameno, a.date "
		sql = sql & "from match a join match_player b on a.date = b.date "
		sql = sql & "where startpos = 1 "
		sql = sql & ") "
		sql = sql & "select b.date as lastgamedate "
		sql = sql & "from cte1 a join cte1 b on a.gameno = b.gameno+1 "
		sql = sql & "where a.date = '" & matchdate & "' "
		rslineup.open sql,conn,1,2

		if not rslineup.EOF then lastgamedate = rslineup.Fields("lastgamedate")
	
		rslineup.close
	
		if isdate(lastgamedate) then 	'Check that it's not the first ever game
		
			sql = "select b.player_id_spell1, b.surname as start_surname, b.forename as start_forename "
			sql = sql & "from match_player a "
			sql = sql & "join player b on a.player_id = b.player_id "
			sql = sql & "where a.date = '" & lastgamedate & "' "
			sql = sql & "  and a.startpos > 0 "
			sql = sql & "order by a.startpos "
			rslineup.open sql,conn,1,2
	
			Do While Not rslineup.EOF
				laststarters(i,0) = rslineup.Fields("player_id_spell1")
				laststarters(i,1) = trim(rslineup.Fields("start_forename")) & " " & trim(rslineup.Fields("start_surname"))
				i = i + 1
				rslineup.MoveNext
			Loop
 
			rslineup.close
		
		end if
	
	'Now get lineup for this game

		sql = "select b.player_id_spell1, b.surname as start_surname, b.forename as start_forename, a.card as start_card, d.surname as sub_surname, d.forename as sub_forename, c.sub_time, c.card as sub_card, c.player_id as sub_playerid, d.player_id_spell1 as sub_playerid_spell1 "
		sql = sql & "from match_player a "
		sql = sql & "join player b on a.player_id = b.player_id "
		sql = sql & "join player b1 on b.player_id_spell1 = b1.player_id "
		sql = sql & " left outer join match_player c on (a.date = c.date and a.replaced_by = c.player_id) left outer join player d on c.player_id = d.player_id  "
		sql = sql & "where a.date = '" & matchdate & "' "
		sql = sql & "  and a.startpos > 0 "
		sql = sql & "order by a.startpos "
		rslineup.open sql,conn,1,2
		
		Do While Not rslineup.EOF
		
			teamfound = "y"
		
			thisstarters(j,0) = rslineup.Fields("player_id_spell1")
			thisstarters(j,1) = trim(rslineup.Fields("start_forename")) & " " & trim(rslineup.Fields("start_surname"))
			j = j + 1
			
			playerid = rslineup.Fields("player_id_spell1")
						
			outteam = outteam & "<span style=""white-space: nowrap;""><a target=""_blank"" href=""gosdb-players2.asp?pid=" & playerid & """>"
			outteam = outteam & trim(rslineup.Fields("start_forename")) & " " & trim(rslineup.Fields("start_surname"))
			outteam = outteam & "</a>"
			
			if not IsNull(rslineup.Fields("start_card")) then outteam = outteam & "[" & rtrim(rslineup.Fields("start_card")) & "]" 

			sub_surname = trim(rslineup.Fields("sub_surname")) 
			sub_forename = trim(rslineup.Fields("sub_forename")) 
			sub_playerid = rslineup.Fields("sub_playerid")
			sub_playerid_spell1 = rslineup.Fields("sub_playerid_spell1")
			sub_time = trim(rslineup.Fields("sub_time"))
			if isnumeric(sub_time) then
				'if sub_time > 90 then sub_time = "90+" & sub_time-90
			end if
			sub_card = rslineup.Fields("sub_card")
			sub_brackets = "" 
			 
			Do While sub_surname > ""
			
				outteam = outteam & " (<a target=""_blank"" href=""gosdb-players2.asp?pid=" & sub_playerid_spell1 & """>"
				outteam = outteam & sub_forename & " " & sub_surname
				outteam = outteam & "</a>"
				
				if not IsNull(sub_card) then outteam = outteam & "[" & rtrim(sub_card) & "]" 
				if not IsNull(sub_time) then outteam = outteam & " " & sub_time
				sub_brackets = sub_brackets & ")"
				sub_surname = "" 
				
				'Find if this sub has himself been subbed
				sql = "select c.player_id as sub_playerid, d.player_id_spell1 as sub_playerid_spell1, d.surname as sub_surname, d.forename as sub_forename, c.sub_time, c.card as sub_card "
				sql = sql & "from match_player a join player b on a.player_id = b.player_id "
				sql = sql & " left outer join match_player c on (a.date = c.date and a.replaced_by = c.player_id) left outer join player d on c.player_id = d.player_id  "
				sql = sql & "where a.date = '" & matchdate & "' "
				sql = sql & "  and a.player_id = " & sub_playerid
				rssubbedsubs.open sql,conn,1,2

				if not rssubbedsubs.EOF then
					sub_surname = trim(rssubbedsubs.Fields("sub_surname")) 
					sub_forename = trim(rssubbedsubs.Fields("sub_forename")) 
					sub_playerid = rssubbedsubs.Fields("sub_playerid")
					sub_playerid_spell1 = rssubbedsubs.Fields("sub_playerid_spell1")
					sub_time = trim(rssubbedsubs.Fields("sub_time"))
					'if sub_time > 90 then sub_time = "90+" & sub_time-90
					sub_card = rssubbedsubs.Fields("sub_card")
				end if
				rssubbedsubs.close
			Loop
			
			outteam = outteam & sub_brackets & ",</span> "

			rslineup.MoveNext
		Loop
		
		if len(outteam) > 9 then outteam = left(outteam,len(outteam)-9)  & ".</span>"  'remove last comma and space and add full stop
		rslineup.close
		

		'Check for unattributable subs

		'Get older-type subs (no indication of who subbed for)
		sql = "select player_id_spell1, surname, forename, initials, startpos "
		sql = sql & "from match_player a join player b on a.player_id = b.player_id "
		sql = sql & "where date = '" & matchdate & "' "
		sql = sql & "  and startpos = 0 and sub_replacing is null "
		sql = sql & "order by surname "

		rslineup.open sql,conn,1,2
		
		if rslineup.RecordCount > 0 then
			if rslineup.RecordCount = 1 then 
				outteam  = outteam & " Sub: "
			  else 
				outteam  = outteam & " Subs: "
			end if
			
			Do While Not rslineup.EOF
				outteam = outteam & "<span style=""white-space: nowrap;""><a target=""_blank"" href=""gosdb-players2.asp?pid=" & rslineup.Fields("player_id_spell1") & """>"
				if rslineup.Fields("forename") > "" then
					outteam = outteam & trim(rslineup.Fields("forename")) & " " & trim(rslineup.Fields("surname"))
				  elseif rslineup.Fields("initials") > "" then
					outteam = outteam & left(rslineup.Fields("initials"),1) & ". " & trim(rslineup.Fields("surname"))
				  else 
				  	outteam = outteam & trim(rslineup.Fields("surname"))	 
				end if
				outteam = outteam & "</a>,</span> "
				rslineup.MoveNext
			Loop
			
			if len(outteam) > 9 then outteam = left(outteam,len(outteam)-9)  & ".</span>"  'remove last comma, /span and space, and add full stop and /span

		end if
				
		rslineup.close
		

		outteam = Replace(outteam,"[y]","<img src=""images/yellowcard.gif"" hspace=""2"" height=""7"" width=""7"" align=""bottom"">") 
		outteam = Replace(outteam,"[yr]","<img src=""images/yelredcard.gif"" hspace=""2"" height=""8"" width=""10"" align=""bottom"">")  
		outteam = Replace(outteam,"[r]","<img src=""images/redcard.gif"" hspace=""2"" height=""7"" width=""7"" align=""bottom"">")
		
		if isdate(lastgamedate) then	'Check that it's not the first ever game

			for i = 0 to 10
				foundplayer = 0
				for j = 0 to 10
					if laststarters(i,0) = thisstarters(j,0) then 
						foundplayer = "1"
						exit for
					end if
				next
				if foundplayer = 0 then 
					outofteam = outofteam & "<a target=""_blank"" href=""gosdb-players2.asp?pid=" & laststarters(i,0) & """>"
					outofteam = outofteam & laststarters(i,1) & "</a>, "
				end if
			next	
		
			if outofteam > "" then outofteam  = left(outofteam, len(outofteam)-2) & "."		'replace final comma and space with full stop
		
			for j = 0 to 10
				foundplayer = 0
				for i = 0 to 10
					if thisstarters(j,0) = laststarters(i,0) then 
						foundplayer = "1"
						exit for
					end if
				next
				if foundplayer = 0 then 
					intoteam = intoteam & "<a target=""_blank"" href=""gosdb-players2.asp?pid=" & thisstarters(j,0) & """>"
					intoteam = intoteam & thisstarters(j,1) & "</a>, "
				end if
			next	
		
			if intoteam > "" then intoteam  = left(intoteam, len(intoteam)-2) & "."		'replace final comma and space with full stop 
			
		end if 
		
End Function  'Getteam
 
Function Getgoals(matchdate)
		 
		  sql = "select b.surname, b.forename, b.initials, time, seqno, pen_ind, a.player_id, b.player_id_spell1 "
		  sql = sql & "from match_goal a "
		  sql = sql & "join player b on a.player_id = b.player_id "
		  sql = sql & "join player b1 on b.player_id_spell1 = b1.player_id "
		  sql = sql & "where date = '" & matchdate & "' "
		  sql = sql & "order by seqno "

		  rsgoals.open sql,conn,1,2
		
		  if rsgoals.RecordCount > 0 then 
		 	
		  	Dim goals(12,9,1)
	
			Do While Not rsgoals.EOF

				scorer = trim(rsgoals.Fields("forename")) & " " & trim(rsgoals.Fields("surname"))
							
			for i = 0 to 12
				
					if isnull(rsgoals.Fields("time")) then		' no goal times stored, so just list players in order
						if goals(i,0,0) = "" then
							goals(i,0,0) = rsgoals.Fields("player_id_spell1")
			  				goals(i,1,0) = scorer
			  				exit for
			  			end if	
				
					 else										' goal time stored, so group goals together for each player
						if goals(i,0,0) = rsgoals.Fields("player_id_spell1") then
							for j = 3 to 9
								if goals(i,j,0) = "" then
			  						goals(i,j,0) = rsgoals.Fields("time") 
					  				if rsgoals.Fields("pen_ind") = "Y" then goals(i,j,1) = "pen"
									exit for
								end if 	
							next		
							exit for
			  	 	 	  elseif goals(i,0,0) = "" then
			  				goals(i,0,0) = rsgoals.Fields("player_id_spell1")
			  				goals(i,1,0) = scorer
	  						goals(i,2,0) = rsgoals.Fields("time")
			  				if rsgoals.Fields("pen_ind") = "Y" then goals(i,2,1) = "pen"
			  				exit for
			 			end if
			 		
			 		end if	
				next
				if i = 12 or j = 9 then outgoals = outgoals & "A problem has occurred, please report this to Steve."
				rsgoals.MoveNext
			Loop

			i = 0
			k = 0
			
			Do while goals(i,0,0) > "" 
				
				if goals(i,0,0) = 9000 then				'player_id=9000, so an own-goal
					outgoals = outgoals & rtrim(ogscorers(k)) & " (o.g.)" 'own-goal scorer
					k = k + 1				  
				  else
					outgoals = outgoals & "<a target=""_blank"" href=""gosdb-players2.asp?pid=" & goals(i,0,0) & """>" & goals(i,1,0) & "</a>"	'goalscorer
				end if

				if goals(i,2,0) > "" then			'if a goal time is available, then put times in brackets

					if goals(i,0,0) = 9000 then							'player_id=9000, so an own-goal
						outgoals = left(outgoals,len(outgoals)-1) & " "	'replace the closing bracket from (o.g.) with a blank
					  else	
						outgoals = outgoals & " ("
					end if
					
					j = 2
					do while goals(i,j,0) > ""
						if goals(i,j,1) > "" then outgoals = outgoals & goals(i,j,1) & " "
						if lfc <> "C" and goals(i,j,0) > 90 then goals(i,j,0) = "90+" & goals(i,j,0)-90  'don't do this for cup games as it could be extra time
						outgoals = outgoals & goals(i,j,0) & ", "
						j = j + 1
					loop
					outgoals = left(outgoals,len(outgoals)-2) & "), "  	'remove last comma and replace with close bracket, comma and space
					
				 else								'no goal times, so just list players in order stored   
						outgoals = outgoals & ", "
				 
				end if
				i = i + 1
			loop
			
			outgoals = left(outgoals,len(outgoals)-2) 				 	'remove last comma and space			
	
		  end if
				
		  rsgoals.close

End Function  'Getgoals


Function Getmilestones(matchdate)
		 
		  sql = "select type, seq, quantity, surname, forename, initials, opposition, milestone_details "
		  sql = sql & "from match_milestone "
		  sql = sql & "where date = '" & matchdate & "' "
		  sql = sql & "order by seq, surname, forename "

		  rsmiles.open sql,conn,1,2
		
		  if rsmiles.RecordCount > 0 then 
		  
		  	output_milestones = output_milestones & "<p style=""margin: 10px 0 2px;"" class=""style1bold"">Match Milestones</p>"
			output_milestones = output_milestones & "<ul>"
			
			Do While Not rsmiles.EOF
		
				if IsNull(rsmiles.Fields("forename")) then
					factname = rtrim(rsmiles.Fields("initials")) & " " & rtrim(rsmiles.Fields("surname"))
	  			  else
	  				factname = rtrim(rsmiles.Fields("forename")) & " " & rtrim(rsmiles.Fields("surname"))	  			
				end if
				
				Select Case rsmiles.Fields("quantity") mod 10
					Case 1 
						quantity = formatNumber(rsmiles.Fields("quantity"),0,,,-1) & "st"
					Case 2 
						quantity = formatNumber(rsmiles.Fields("quantity"),0,,,-1) & "nd"
					Case else 
						quantity = formatNumber(rsmiles.Fields("quantity"),0,,,-1) & "th"
				End Select
				
				output_milestones = output_milestones & "<li>"
				
				Select Case rtrim(rsmiles.Fields("type"))
					Case "CM"
						output_milestones = output_milestones & factname & "PAFC's " & quantity & " match in all competitions" 
					Case "CMH"
						output_milestones = output_milestones & factname & "PAFC's " & quantity & " match at home"
					Case "CMA"
						output_milestones = output_milestones & factname & "PAFC's " & quantity & " away match"
					Case "CMFL"
						output_milestones = output_milestones & factname & "PAFC's " & quantity & " Football League match" 
					Case "CMFLH"
						output_milestones = output_milestones & factname & "PAFC's " & quantity & " Football League match at home"
					Case "CMFLA"
						output_milestones = output_milestones & factname & "PAFC's " & quantity & " away match in the Football League"
					Case "CG"
						output_milestones = output_milestones & factname & " scored PAFC's " & quantity & " goal in all competitions" 
					Case "CGH"
						output_milestones = output_milestones & factname & " scored PAFC's " & quantity & " goal at home" 
					Case "CGA"
						output_milestones = output_milestones & factname & " scored PAFC's " & quantity & " away goal" 
					Case "CGFL"
						output_milestones = output_milestones & factname & " scored PAFC's " & quantity & " Football League goal" 
					Case "CGFLH"
						output_milestones = output_milestones & factname & " scored PAFC's " & quantity & " Football League goal at home" 
					Case "CGFLA"
						output_milestones = output_milestones & factname & " scored PAFC's " & quantity & " away goal in the Football League"
					Case "CGO"
						output_milestones = output_milestones & factname & " scored PAFC's " & quantity & " goal against " & rsmiles.Fields("opposition")
					Case "CMFAC"
						output_milestones = output_milestones & factname & "PAFC's " & quantity & " FA Cup match" 
					Case "CMFACH"
						output_milestones = output_milestones & factname & "PAFC's " & quantity & " FA Cup match at home"
					Case "CMFACA"
						output_milestones = output_milestones & factname & "PAFC's " & quantity & " away match in the FA Cup"
					Case "CGFAC"
						output_milestones = output_milestones & factname & " scored PAFC's " & quantity & " FA Cup goal" 
					Case "CGFACH"
						output_milestones = output_milestones & factname & " scored PAFC's " & quantity & " FA Cup goal at home" 
					Case "CGFACA"
						output_milestones = output_milestones & factname & " scored PAFC's " & quantity & " away goal in the FA Cup"
					Case "M1M"
						output_milestones = output_milestones & "Manager " & factname & "'s first match in charge"
					Case "MLM"
						output_milestones = output_milestones & "Manager " & factname & "'s " & quantity & " and last match in charge"
					Case "MM"
						output_milestones = output_milestones & "Manager " & factname & "'s " & quantity & " match in charge"
					Case "O1M"
						output_milestones = output_milestones & "PAFC's first game against " & rsmiles.Fields("opposition")
					Case "O1MFL"
						output_milestones = output_milestones & "PAFC's first game against " & rsmiles.Fields("opposition") & " in the Football League"
					Case "OM"
						output_milestones = output_milestones & "PAFC's " & quantity & " game against " & rsmiles.Fields("opposition")
					Case "OMFL"
						output_milestones = output_milestones & "PAFC's " & quantity & " game against " & rsmiles.Fields("opposition") & " in the Football League"
					Case "P1M"
						output_milestones = output_milestones & factname & "'s Argyle debut"
					Case "P1S"
						output_milestones = output_milestones & factname & "'s first start for Argyle"
					Case "POM"
						output_milestones = output_milestones & factname & "'s only game for Argyle"
					Case "PLM"
						output_milestones = output_milestones & factname & "'s last game for Argyle"
					Case "PM"
						output_milestones = output_milestones & factname & "'s " & quantity & " game for Argyle"
					Case "PMFL"
						output_milestones = output_milestones & factname & "'s " & quantity & " Football League game for Argyle"
					Case "PMFAC"
						output_milestones = output_milestones & factname & "'s " & quantity & " FA Cup game for Argyle"
					Case "PS"
						output_milestones = output_milestones & factname & "'s " & quantity & " start for Argyle"
					Case "PSFL"
						output_milestones = output_milestones & factname & "'s " & quantity & " Football League start for Argyle"
					Case "PSFAC"
						output_milestones = output_milestones & factname & "'s " & quantity & " FA Cup start for Argyle"
					Case "P1G"
						output_milestones = output_milestones & factname & "'s first goal for Argyle"	
					Case "POG"
						output_milestones = output_milestones & factname & "'s only goal for Argyle"
					Case "PLG"
						output_milestones = output_milestones & factname & "'s last goal for Argyle"
					Case "PG"
						output_milestones = output_milestones & factname & "'s " & quantity & " goal for Argyle"
					Case "PGFL"
						output_milestones = output_milestones & factname & "'s " & quantity & " Football League goal for Argyle"
					Case "PGFAC"
						output_milestones = output_milestones & factname & "'s " & quantity & " FA Cup goal for Argyle"	
					Case "MANUAL"
						output_milestones = output_milestones & rsmiles.Fields("milestone_details")					
				End Select
				
				output_milestones = output_milestones & "</li>"
				
				rsmiles.MoveNext
			Loop
			
			output_milestones = output_milestones & "</ul>"

		end if
				
		rsmiles.close

End Function  'Getmilestones



%>