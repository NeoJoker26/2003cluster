<%
response.expires = -1

' Set Session Locale to English(UK) to solve date differences when swapping from Fasthosts to Tsohost
' [Tsohost defaults to 1033, which is English(US)]
Session.LCID=2057


Dim lastgamedate, teamfound, laststarters(10,1), thisstarters(10,1), intoteam, outofteam, report, reporttext, videomark, shortrange

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
Set rsgoals = Server.CreateObject("ADODB.Recordset")
Set rslineup = Server.CreateObject("ADODB.Recordset")
Set rssubbedsubs = Server.CreateObject("ADODB.Recordset")
Set rsmiles = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_read.inc"--><%

pass = Request.QueryString("pass")
phase = Request.QueryString("phase")


select case pass

	case 1
		
		decade = Request.QueryString("decade")
		start_year = left(decade,4)
		end_year = left(decade,2) & mid(decade,6,2) 
		output = "<ul class=""nav""><li class=""header"")>Seasons</li>"
		
		'Get seasons
		sql = "select years, sum(count) as videocount "
		sql = sql & "from (select distinct years, 0 as count "
		sql = sql & "		from v_match_season "
		sql = sql & "		where left(years,4) between '" & start_year & "' and '" & end_year &"' "
		sql = sql & "		union all "
		sql = sql & "		select years, 1 "
		sql = sql & "		from v_match_season a left join event_control b on date = event_date "
		sql = sql & "		where left(years,4) between '" & start_year & "' and '" & end_year & "' "
		sql = sql & "		  and event_published = 'Y' and ((event_type = 'M' and material_type = 'Y') or event_type = 'V') "
		sql = sql & "		) x " 
		sql = sql & "group by years order by years "

		rs.open sql,conn,1,2
		
		Do While Not rs.EOF
			shortrange = left(rs.Fields("years"),5) & right(rs.Fields("years"),2)
			output = output & "<li id=""s" & left(rs.Fields("years"),4) & """ class=""cell boxwidth"">" & shortrange
			if rs.Fields("videocount") > 0 then output = output & "<span style=""float:right; letter-spacing:-1px; padding:0 1px; margin:1px 1px 0 0; font-size:9px; color:green; background-color:white;"">" & rs.Fields("videocount") & "</span>"			
			output = output & "</li>"
			rs.MoveNext
		Loop	
		
		output = output & "</ul>"
		
	case 2
		
		season = Request.QueryString("season")
		if mid(season,6,2) = "00" then
			season = left(season,5) & "2000"
		  else
		  	season = left(season,5) & left(season,2) & mid(season,6,2)
		end if
 
		output = "<table"
		if left(season,4) = "1944" then output = output & " style=""width:355px""" 
		output = output & ">"
		x = 0
		y = 0
		lastmonth = ""
		dim grid (12,20)
		grid(0,0) = "<td class=""header"">Matches</td>"
		
		'Get matches
		sql = "select distinct '1' as type, date, day(date) as day, left(datename(m,date),3) as month, year(date) as year, NULL as notplayedtype, material_type as video_exists "
		sql = sql & "from match a join season b "
		sql = sql & "  on date >= date_start and date <= date_end "
		sql = sql & "left join event_control c "
		sql = sql & "  on date = event_date and event_published = 'Y' and ((event_type = 'M' and material_type = 'Y') or event_type = 'V') "
		sql = sql & "where years = '" & season & "' "
		
		sql = sql & " union all "
		sql = sql & "select '2' as type, date, day(date) as day, left(datename(m,date),3) as month, year(date) as year, NULL, NULL "
		sql = sql & "from season_this a join season b "
		sql = sql & "  on date >= date_start and date <= date_end "
		sql = sql & "where years = '" & season & "' "
		sql = sql & "  and not exists (select date from match x where x.date = a.date) "
		
		sql = sql & " union all "
		sql = sql & "select '3' as type, date, day(date) as day, left(datename(m,date),3) as month, year(date) as year, not_played_type, NULL "
		sql = sql & "from match_not_played a join season b "
		sql = sql & "  on date >= date_start and date <= date_end "
		sql = sql & "where years = '" & season & "' "
		
		sql = sql & "order by date "

		rs.open sql,conn,1,2
		
		Do While Not rs.EOF
			if rs.Fields("month") <> lastmonth then 
				x = x + 1
				y = 0
				lastmonth = rs.Fields("month")
			end if
			cellqual = ""
			if rs.Fields("type") = 1 then 
				cellgrey = ""
			  elseif rs.Fields("type") = 2 then
			    cellgrey = "cellgrey"
			  else
			    cellgrey = ""
			    if rs.Fields("notplayedtype") = "P" then
			    	cellqual = " [P]"
			      elseif rs.Fields("notplayedtype") = "A" then
			    	cellqual = " [A]"
			      elseif rs.Fields("notplayedtype") = "C" then
			    	cellqual = " [C]"
			    end if
			end if
			videomark = ""
			if rs.Fields("video_exists") = "Y" then videomark = "<img style=""float:right; border:0; padding:1px 2px 0 0;"" src=""images/video7x12.gif"">"
   			grid(x,y) = "<td id=""" & rs.Fields("year") & rs.Fields("month") & rs.Fields("day") & """ class=""cell " & cellgrey & """>" & rs.Fields("month") & " " & rs.Fields("day") & cellqual & videomark & "</td>"
			if y > max_y then max_y = y
			y = y + 1
			rs.MoveNext
		Loop
		
		max_x = x
		
		for y = 0 to max_y
		  output = output & "<tr>"
			for x = 0 to max_x
				if grid(x,y) = "" then 
					output = output & "<td></td>"
				  else
					output = output & grid(x,y)	
				end if	
			next
		  output = output & "</tr>"
		next	
		output = output & "</table>"

		
	case 3
		
		season = Request.QueryString("season")
		if mid(season,6,2) = "00" then
			season = left(season,5) & "2000"
		  else
		  	season = left(season,5) & left(season,2) & mid(season,6,2)
		end if

		
		matchdate = Request.QueryString("date")	'It's in the form 2015May8
		'convert matchdate to conventional form
			matchmonth = mid(matchdate,5,3)
			Select Case matchmonth 
				Case "Jan" monthno = "01" 
				Case "Feb" monthno = "02" 
				Case "Mar" monthno = "03" 
				Case "Apr" monthno = "04" 
				Case "May" monthno = "05" 
				Case "Jun" monthno = "06" 
				Case "Jul" monthno = "07" 
				Case "Aug" monthno = "08" 
				Case "Sep" monthno = "09" 
				Case "Oct" monthno = "10" 
				Case "Nov" monthno = "11" 
				Case "Dec" monthno = "12" 
			End Select 
			matchdate = left(matchdate,4) & "-" & monthno & "-" & right("0" & mid(matchdate,8),2) 
		
	' Get match details

	sql = "select season_no, division, a.date, opposition, opposition_qual, name_then_short, lfc, homeaway, goalsfor, goalsagainst, aet_ind, pensfor, pensagainst, totpoints, position, max_prog_image_no, notes, managers, "
	sql = sql & "competition, subcomp, attendance, visitors, non_playing_subs, opp_team, opp_goals, shootout_argyle, shootout_opp, headline, report, report_published, report_acknowledge, "
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
		totpointshold = rs.Fields("totpoints")
		positionhold = rs.Fields("position")

		output_left = "<div id=""left-side"" style=""float:left; width:23%; margin-top:8px;"">"
		output_left = output_left & "<a href=""gosdb.asp""><img style=""float:left; border:0; margin-right:3px;"" src=""images/gosdb-small.jpg""></a>"
		output_left = output_left & "<span style=""color: #404040; font-size: 15px; font-weight: 700"">Match<br>View</span>"

       	output_right = "<div id=""right-side"" style=""float:right; width:77%;"">"
       	       	
        	output_right = output_right & "<div id=""right-side2"" style=""float:right; margin:12px 9px 0 12px"">"
        	
					sql = "with cte as ( "
					sql = sql & "select row_number() over (order by date) as rownum, date "
					sql = sql & "from match "
					sql = sql & "where opposition = '" & replace(clubhold,"'","''") & "' "			'Double-up any apostrophes inside the club name (e.g. Lovell's)
					sql = sql & ") "
					sql = sql & "select prevdate=prev.date, thisdate=cte.date, nextdate=next.date "
					sql = sql & "from cte "
					sql = sql & "left outer join cte prev on prev.rownum = cte.rownum - 1 "
					sql = sql & "left outer join cte next on next.rownum = cte.rownum + 1 "
					sql = sql & "where cte.date = '" & matchdate & "' "
			
					rs1.open sql,conn,1,2

 		 			output_right = output_right & "<ul class=""nav2"">"
					if not IsNull(rs1.Fields("prevdate")) then output_right = output_right & "<a href=""gosdb-match.asp?date=" & rs1.Fields("prevdate") & """><li class=""cell"">Prev " & clubshorthold & "</li></a>"
					if not IsNull(rs1.Fields("nextdate")) then output_right = output_right & "<a href=""gosdb-match.asp?date=" & rs1.Fields("nextdate") & """><li class=""cell"">Next " & clubshorthold & "</li></a>"
					team = Replace(clubhold," ","%20")
					team = Replace(team,"&","%26")
					output_right = output_right & "<a target=""_blank"" href=""gosdb-results.asp?team=" & team & """><li class=""cell"">All " & clubshorthold & "</li></a>"
					output_right = output_right & "<a target=""_blank"" href=""gosdb-season.asp?years=" & season & """><li class=""cell"">" & left(season,5) & mid(season,8,2) & " Season</li></a></ul>"
						
					rs1.close
        	
        	output_right = output_right & "</div>"
        	
        	if rs.Fields("max_prog_image_no") > 0 then
        	
        		output_right = output_right & "<div id=""right-side3"" style=""clear:both; float:right; max-width:36%; margin:12px 0 12px 12px; border:1px solid #c0c0c0; padding:3px; text-align:center; "">"
        			
        		output_right = output_right & "<p style=""margin:4px 0;"">PROGRAMME EXTRACTS</p>" 
        			
        		for n = 1 to rs.Fields("max_prog_image_no")

        			output_right = output_right & "<a class=""highslide"" onclick=""return hs.expand(this,{slideshowGroup: 'programmes', captionId: 'progcaption', allowSizeReduction: 'true', fullExpandOpacity: 80})"" " 
        			output_right = output_right & "href=""gosdb/photos/programmes/" & rs.Fields("date") & "-" & n & ".jpg"">"
        			output_right = output_right & "<img style=""margin:2px 2px 0; vertical-align:top; max-width:190px; max-height:150px; "" src=""gosdb/photos/programmes/" & rs.Fields("date") & "-" & n & ".jpg"">"
        			output_right = output_right & "</a>"
        			       				
        		next
        	
				output_right = output_right & "<p style=""font-size:10px; margin:5px auto; max-width: 95%; line-height:1.2; "">Click to expand. Magnify again by clicking on the bottom-right corner of the expanded image (especially useful for 1974-77).</p>"        			
       		
        		if rs.Fields("max_prog_image_no") = 1 then output_right = replace(output_right, "an image", "the image")
   					
        		output_right = output_right & "</div>" 
        	
			end if	
        	
       	output_right = output_right & "<h1>" & FormatDateTime(rs.Fields("date"),1) & "</h1>"'
       	output_right = output_right & "<h2>" & rs.Fields("competition")
       	if not IsNull(rs.Fields("subcomp")) then output_right = output_right & " " & trim(rs.Fields("subcomp"))
       	lfc = rs.Fields("lfc")
   		output_right = output_right & "</h2>"

       
		''''' Result '''''
       
        if rs.Fields("homeaway") = "H" then
			output_right = output_right & "<p class=""score"">Argyle &nbsp;" & rs.Fields("goalsfor") & " - " & rs.Fields("goalsagainst") & "&nbsp; " & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & "</p>"
			if ucase(rs.Fields("aet_ind")) = "Y" then output_right = output_right & "<p class=""aet"">After extra time</p>"
 		  	if not isnull(rs.Fields("pensfor")) then
		  	   	output_right = output_right & "<p class=""penalties"">Penalties: Argyle " & rs.Fields("pensfor") & " - " & rs.Fields("pensagainst") & " " & rs.Fields("name_then_short") & "</p>"
			end if
		  else 
			output_right = output_right & "<p class=""score"">" & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " &nbsp;" & rs.Fields("goalsagainst") & " - " & rs.Fields("goalsfor") & "&nbsp; Argyle" & "</p>"
			if ucase(rs.Fields("aet_ind")) = "Y" then output_right = output_right & "<p class=""aet"">After extra time</p>"
			if not isnull(rs.Fields("pensfor")) then
			   	output_right = output_right & "<p class=""penalties"">Penalties: " & rs.Fields("name_then_short") & " " & rs.Fields("pensagainst") & " - " & rs.Fields("pensfor") & " Argyle" & "</p>"	
			end if		   		  
		end if
		
		''''' Venue etc '''''

	    output_right = output_right & "<p style=""margin:9px 0 3px;"">"
	    output_right = output_right & "<span class=""bold"">Venue: </span><span class=""venue"">" 
    	if rs.Fields("homeaway") = "H" then
    		output_right = output_right & rs.Fields("home_ground_name")
       	  else 
    		output_right = output_right & rs.Fields("away_ground_name")
    		if not isnull(rs.Fields("ground_name_trad")) then output_right = output_right & " (aka " & rs.Fields("ground_name_trad") & ")"
		end if
		output_right = output_right & "</span>"
		
		if not IsNull(rs.Fields("attendance")) then 
			output_right = output_right & "<span class=""bold"">Attendance: </span>" 
			output_right = output_right & "<span class=""attendance"">" & FormatNumber(rs.Fields("attendance"),0,,-1) & "</span>"
		end if
		if not IsNull(rs.Fields("visitors")) then 
			output_right = output_right & "<span class=""bold"">Visitors: </span>" 
			output_right = output_right & "<span class=""visitors"">" & rs.Fields("visitors") & "</span>"
		end if
		if not IsNull(rs.Fields("referee")) then 
			output_right = output_right & "<span style=""white-space: nowrap""><span class=""bold"">Referee: </span>" 
			output_right = output_right & rs.Fields("referee") & "</span>"
		end if
	
       	output_right = output_right & "</p>"
       	
      	output_right = output_right & "<p style=""margin:3px 0 9px;"">"
      	if not IsNull(totpointshold) then
	    	output_right = output_right & "<span class=""bold"">Total Points: </span><span class=""totpoints"">"
    		output_right = output_right & totpointshold
			output_right = output_right & "</span>"
		end if
      	if not IsNull(positionhold) then
	    	output_right = output_right & "<span class=""bold"">Position: </span>"
    		output_right = output_right & positionhold
		end if	
	
       	output_right = output_right & "</p>"
       	
       	if not IsNull(rs.Fields("notes")) then output_right = output_right & "<p style=""margin:9px 0;""><span class=""bold"">Note: </span>" & trim(rs.Fields("notes")) & "</p>"
    	    	
 		
		''''' Teams '''''

		outteam  = ""
		photolist_players = ""
		
		Call Getteam(matchdate)
		
		'Check the first cell of the team line-up array to see if a team has been found
		
		if teamfound = "y" then
		
			output_argyle = output_argyle & "<p class=""team""><span class=""style1bold"">ARGYLE</span>: " & outteam 
				
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
		photolist_scorers = ""
		
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
		
		if not IsNull(rs.Fields("shootout_argyle")) then
			shootout_heading = "<p class=""style1bold"" style=""margin: 6px 0 0"">Penalty Shootout</p>"  
			shootout_argyle = "<span class=""style1"">Argyle:</span> " & rs.Fields("shootout_argyle") & "<br>"
			shootout_opposition = "<span class=""style1"">" & clubshorthold & ":</span> " & rs.Fields("shootout_opp") & "<br>"
		end if
				
		if homeawayhold = "H" then
			output_right = output_right & output_argyle & output_opposition
			if goals_argyle > "" or goals_opposition > "" then 
				output_right = output_right & "<p class=""style1bold goals"">GOALS</p>"
				output_right = output_right & goals_argyle & goals_opposition & shootout_heading & shootout_argyle & shootout_opposition
			end if
			if not IsNull(rs.Fields("shootout_argyle")) then output_right = output_right & shootout_heading & shootout_argyle & shootout_opposition
		  else
			output_right = output_right & output_opposition & output_argyle
			if goals_argyle > "" or goals_opposition > "" then 
				output_right = output_right & "<p class=""style1bold goals"">GOALS</p>"
				output_right = output_right & goals_opposition & goals_argyle
			end if
			if not IsNull(rs.Fields("shootout_argyle")) then output_right = output_right & shootout_heading & shootout_opposition & shootout_argyle
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
				report = report & "</p><p style=""margin:9px 0 15px"">[Summary from <span style=""font-style:italic"">Plymouth Argyle, The Modern Era, A Complete Record</span> by Andy Riddle, with the author's kind permission.]</p></div>"
			  elseif rs.Fields("report_acknowledge") = "H" then
				report = "<div id=""report""><p style=""margin:15px 0 0"">" & reporttext  
				report = report & "</p><p style=""margin:9px 0 15px"">[Extracted from <span style=""font-style:italic"">Harley Lawer's Argyle Classics</span>, with the author's kind permission.]</p></div>"
			  elseif rs.Fields("report_acknowledge") = "L" then
				report = "<div id=""report""><p style=""margin:15px 0 0"">" & reporttext  
				report = report & "</p><p style=""margin:9px 0 15px"">[Report thanks to Alec Hepburn]</p></div>"
			  elseif rs.Fields("report_acknowledge") = "J" then
				report = "<div id=""report""><p style=""margin:15px 0 0"">" & reporttext  
				report = report & "</p><p style=""margin:9px 0 15px"">[Report thanks to John Eales]</p></div>"
			  elseif rs.Fields("report_acknowledge") = "P" then
				report = "<div id=""report""><p style=""margin:15px 0 0"">" & reporttext  
				report = report & "</p><p style=""margin:9px 0 15px"">[A summary of PAFC's match report, with thanks]</p></div>"
			  else
				report = "<div id=""report""><p style=""margin:15px 0 0"">" & reporttext & "</p></div>"
			end if 
		
			if rs.Fields("report_published") = "Y" then
				output_right2 = output_right2 & report
			  elseif phase = "review"	then
			    output_right2 = output_right2 & report & "<p style=""color:red; font-weight:bold;"">After review, don't forget to use your back button (possibly twice) to amend or publish.</p>"
    		end if
    		
    	  else
    	  	
			output_right2 = output_right2 & "<br><br><br>" 	'No report, so add a few lines to avoid formatting issues for the charts and table
    	
    	end if
	
		if rs.Fields("lfc") = "F" then
		
			output_right2 = output_right2 & "<ul id=""chartbuttons"">"
			if not IsNull(totpointshold) then output_right2 = output_right2 & "<li id=""A" & matchdate & """ class=""cell"">Points Progress</li>"
			if not IsNull(positionhold) then output_right2 = output_right2 & "<li id=""B" & matchdate & """ class=""cell"">Position Progress</li>"
			output_right2 = output_right2 & "<li id=""C" & matchdate & """ class=""cell"">League Table Plus</li>"
			output_right2 = output_right2 & "</ul>"
		
		end if		

		rs.close
		
		
	'This section displays the milestones and any match material
		
		output_milestones = "<div id=""milestones"">"
	
			Call Getmilestones(matchdate)
			
		output_milestones = output_milestones & "</div>"	'Finish off the milestones div
		
	
	'This section displays the photo and audio links
	
		output_material = ""
		undofloat_class = ""
		
		Call Getmaterial(matchdate)
		
     	
	'Now put all the right-hand div components together
						
		output_right = output_right & output_material & output_milestones & output_right2 
		
		if output_material = "" then output_right = replace(output_right, "<div id=""report"">", "<div id=""report-wider"">") 
		    	
      	output_right = output_right & "</div>"		'Finish off right-hand div

      	 
	 
 	'Now finish off left-hand Div
 	
		if teamfound = "y" then 	'only add a photo when a lineup has been found
			
			if len(photolist_scorers) > 0 then
				photolist_scorers = left(photolist_scorers,len(photolist_scorers)-1)		'drop final separator (^)
				scorerdetails = split(photolist_scorers,"^")
				if ubound(scorerdetails) = 0 then	'only one goal
					scorer = split(scorerdetails(0),",")
					image_id = scorer(0)	
					caption = "Goalscorer <span style=""white-space: nowrap;"">" & scorer(1) & " " & scorer(2) & "</span>"
			  	  else		  		
					randomize
					r = int(rnd*(ubound(scorerdetails)+1)+1)
					scorer = split(scorerdetails(r-1),",")
					image_id = scorer(0)	
					caption = "One of the goalscorers, <span style=""white-space: nowrap;"">" & scorer(1) & " " & scorer(2) & "</span>"
				end if 
		  	  else
		  		if len(photolist_players) > 0 then
		  			photolist_players = left(photolist_players,len(photolist_players)-1)		'drop final separator (^)
					playerdetails = split(photolist_players,"^")
					if goalsagainsthold = 0  then	'no goals scored, but also a clean sheet, so feature the goalkeeper
						scorer = split(playerdetails(0),",")
						image_id = scorer(0)	
						caption = "A clean sheet from <span style=""white-space: nowrap;"">" & scorer(1) & " " & scorer(2) & "</span>"
			  	  	  else
						randomize
						r = int(rnd*(ubound(playerdetails)+1)+1)
						scorer = split(playerdetails(r-1),",")
						image_id = scorer(0)	
						caption = "At random, <span style=""white-space: nowrap;"">" & scorer(1) & " " & scorer(2) & "</span>"
					end if
				end if
			end if 
			
			output_left = output_left & "<img style=""border:0; margin: 0 auto; height:90%; width:90%;"" src=""gosdb/photos/players/" & image_id & ".jpg"">"
			output_left = output_left & "<p class=""caption font11px"">" & caption & "</p>" 
		
		end if
		
		output_left = output_left & "</div>"	'finish off left-side div
		output = output_left & output_right

	else
	
		'No match found, so close and try this season's fixtures
		
		rs.close 
	
		sql = "select season_no, division, a.date, opposition, opposition_qual, name_then_short, homeaway, competition, subcomp "
		sql = sql & "from season_this a "
		sql = sql & "join season b on a.date between b.date_start and b.date_end "
		sql = sql & "join competition b1 on a.compcode = b1.compcode "
		sql = sql & "join opposition c on a.opposition = c.name_then left outer join venue d on a.opposition = d.club_name_then and a.date between d.first_game and d.last_game "
		sql = sql & "where a.date = '" & matchdate & "' "

		rs.open sql,conn,1,2

		if rs.RecordCount > 0 then 
	
			output_left = "<div id=""left-side"" style=""float:left; width:25%; margin-top:8px;"">"
			output_left = output_left & "<a href=""gosdb.asp""><img style=""float:left; border:0; margin-right:3px;"" src=""images/gosdb-small.jpg""></a>"
			output_left = output_left & "<span style=""color: #404040; font-size: 15px; font-weight: 700"">Match<br>View</span>"
			output_left = output_left & "</div>"

       		output_right = "<div id=""right-side"" style=""float:right; width:75%;"">"
       	
       		output_right = output_right & "<div id=""right-side2"" style=""float:right; margin:12px 9px 0 12px"">"
        	
			sql = "select max(date) as prevdate "
			sql = sql & "from match "
			sql = sql & "where opposition = '" & replace(rs.Fields("opposition"),"'","''") & "' "				'Double-up any apostrophes inside the club name (e.g. Lovell's)
		
			rs1.open sql,conn,1,2

 			output_right = output_right & "<ul class=""nav"">"
			if not IsNull(rs1.Fields("prevdate")) then output_right = output_right & "<a href=""gosdb-match.asp?date=" & rs1.Fields("prevdate") & """><li class=""cell"">Prev " & rs.Fields("name_then_short") & "</li></a>"
			team = Replace(rs.Fields("opposition")," ","%20")
			team = Replace(team,"&","%26")
			output_right = output_right & "<a target=""_blank"" href=""gosdb-results.asp?team=" & team & """><li class=""cell"">All " & rs.Fields("name_then_short") & "</li></a>"
			output_right = output_right & "<a target=""_blank"" href=""gosdb-season.asp?years=" & season & """><li class=""cell"">" & left(season,5) & mid(season,8,2) & " Season</li></a></ul>"
			
			rs1.close
        	
        	output_right = output_right & "</div>"
       	        	
	       	output_right = output_right & "<h1>" & FormatDateTime(rs.Fields("date"),1) & "</h1>"'
    	   	output_right = output_right & "<h2>" & rs.Fields("competition")
       		if not IsNull(rs.Fields("subcomp")) then output_right = output_right & " " & trim(rs.Fields("subcomp"))
   			output_right = output_right & "</h2>"
	
			''''' Fixture '''''
       
        	if rs.Fields("homeaway") = "H" then
				output_right = output_right & "<p class=""score"">Argyle v " & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & "</p>"
		  	  else 
				output_right = output_right & "<p class=""score"">" & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " v Argyle" & "</p>"   		  
			end if
		
			rs.close
		
			if datediff("d",matchdate,date) >= 0 then 
		
				output_right = output_right & "<p style=""margin: 24px 0;"">The match result and details have yet to be added." 
		
				'This section displays any photo and audio links
	
				output_material = ""
				undofloat_class = "undofloat"
		
				Call Getmaterial(matchdate)
		
				output_right = output_right & output_material & "</div>"		'Finish off right-hand div
			
		  	  else
		  	
		  		output_right = output_right & "<p style=""margin: 24px 0;"">This match has not been played." 
		
			end if
		
			output = output_left & output_right
			
		  else
		  
		  	'No fixture found in current season, so close and try postponed and abandoned games
		
			rs.close 
	
			sql = "select season_no, division, a.date, opposition, opposition_qual, homeaway, name_then_short, competition, subcomp, not_played_type, date_played, day(date_played) as day, left(datename(m,date_played),3) as month, details "
			sql = sql & "from match_not_played a "
			sql = sql & "join season b on a.date between b.date_start and b.date_end "
			sql = sql & "join competition b1 on a.compcode = b1.compcode "
			sql = sql & "join opposition c on a.opposition = c.name_then left outer join venue d on a.opposition = d.club_name_then and a.date between d.first_game and d.last_game "
			sql = sql & "where a.date = '" & matchdate & "' "

			rs.open sql,conn,1,2

			if rs.RecordCount > 0 then 
	
				output_left = "<div id=""left-side"" style=""float:left; width:25%; margin-top:8px;"">"
				output_left = output_left & "<a href=""gosdb.asp""><img style=""float:left; border:0; margin-right:3px;"" src=""images/gosdb-small.jpg""></a>"
				output_left = output_left & "<span style=""color: #404040; font-size: 15px; font-weight: 700"">Match<br>View</span>"
				output_left = output_left & "</div>"

       			output_right = "<div id=""right-side"" style=""float:right; width:75%;"">"
       	
       			output_right = output_right & "<div id=""right-side2"" style=""float:right; margin:12px 9px 0 12px"">"
        	
				sql = "select max(date) as prevdate "
				sql = sql & "from match "
				sql = sql & "where opposition = '" & replace(rs.Fields("opposition"),"'","''") & "' "				'Double-up any apostrophes inside the club name (e.g. Lovell's)
						
				rs1.open sql,conn,1,2

 				output_right = output_right & "<ul class=""nav"">"
				if not IsNull(rs1.Fields("prevdate")) then output_right = output_right & "<a href=""gosdb-match.asp?date=" & rs1.Fields("prevdate") & """><li class=""cell"">Prev " & rs.Fields("name_then_short") & "</li></a>"
				team = Replace(rs.Fields("opposition")," ","%20")
				team = Replace(team,"&","%26")
				output_right = output_right & "<a target=""_blank"" href=""gosdb-results.asp?team=" & team & """><li class=""cell"">All " & rs.Fields("name_then_short") & "</li></a>"
				output_right = output_right & "<a target=""_blank"" href=""gosdb-season.asp?years=" & season & """><li class=""cell"">" & left(season,5) & mid(season,8,2) & " Season</li></a></ul>"
			
				rs1.close
        	
        		output_right = output_right & "</div>"
       	        	
	       		output_right = output_right & "<h1>" & FormatDateTime(rs.Fields("date"),1) & "</h1>"'
    	   		output_right = output_right & "<h2>" & rs.Fields("competition")
       			if not IsNull(rs.Fields("subcomp")) then output_right = output_right & " " & trim(rs.Fields("subcomp"))
       			output_right = output_right & "</h2>"
	
				''''' Fixture '''''
       
     		   	if rs.Fields("homeaway") = "H" then
					output_right = output_right & "<p class=""score"">Argyle v " & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & "</p>"
		  	  	else 
					output_right = output_right & "<p class=""score"">" & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " v Argyle" & "</p>"   		  
				end if
		
				if rs.Fields("not_played_type") = "P" then
			    	notplayedtype = "POSTPONED"
			      elseif rs.Fields("not_played_type") = "A" then
			    	notplayedtype = "ABANDONED"
			      elseif rs.Fields("not_played_type") = "C" then
			    	notplayedtype = "CANCELLED"
			    end if

				output_right = output_right & "<p class=""font15px bold green"" style=""margin: 15px 0 12px;"">" & notplayedtype & "</p>" 
				output_right = output_right & "<p>" & rs.Fields("details") & "</p>"
				if year(rs.Fields("date_played")) < 9999 then 
					output_right = output_right & "<p>" & "New date: " & rs.Fields("month") & " " & rs.Fields("day") & ".</p>" 
				end if
		
			rs.close
		
			output = output_left & output_right
			
		end if	
	
	end if	
	
end if	
		
end select	
		
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

		sql = "select b.player_id_spell1, b.surname as start_surname, b.forename as start_forename, b1.prime_photo, b1.photo_exists, a.card as start_card, d.surname as sub_surname, d.forename as sub_forename, c.sub_time, c.card as sub_card, c.player_id as sub_playerid, d.player_id_spell1 as sub_playerid_spell1 "
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
			if len(playerid) < 3 then playerid = right("00" & playerid,3)
			
			if rslineup.Fields("photo_exists") = "Y" then
					
				if isnull(rslineup.Fields("prime_photo")) then 
					photoid = playerid
				  else
				  	photoid = playerid & "_" & rslineup.Fields("prime_photo")
				end if
	
				photolist_players = photolist_players & photoid & "," & rslineup.Fields("start_forename") & "," & rslineup.Fields("start_surname") & "^"
			
			end if 
				
			outteam = outteam & "<span style=""white-space: nowrap;"">"
			
			if playerid = 8000 then			'If unknown player, miss out the processing for a link to the player's page, cards, subs etc
			
				outteam = outteam & trim(rslineup.Fields("start_forename")) & " " & trim(rslineup.Fields("start_surname")) & ","
			
			  else
			  	
				outteam = outteam & "<a target=""_blank"" href=""gosdb-players2.asp?pid=" & playerid & """>"
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
			
				outteam = outteam & sub_brackets & ","
				
			end if
			
			outteam = outteam & "</span> "

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
					if laststarters(i,0) = 8000 then		'The id for an unknown player, so don't include a link
					 	outofteam = outofteam & laststarters(i,1) & ", "
					  else
						outofteam = outofteam & "<a target=""_blank"" href=""gosdb-players2.asp?pid=" & laststarters(i,0) & """>"
						outofteam = outofteam & laststarters(i,1) & "</a>, "
					end if
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
					if thisstarters(j,0) = 8000 then		'The id for an unknown player, so don't include a link
					 	intoteam = intoteam & thisstarters(j,1) & ", "
					  else
						intoteam = intoteam & "<a target=""_blank"" href=""gosdb-players2.asp?pid=" & thisstarters(j,0) & """>"
						intoteam = intoteam & thisstarters(j,1) & "</a>, "
					end if 
				end if
			next	
		
			if intoteam > "" then intoteam  = left(intoteam, len(intoteam)-2) & "."		'replace final comma and space with full stop 
			
		end if 
		
End Function  'Getteam
 
Function Getgoals(matchdate)
		 
		  sql = "select b.surname, b.forename, b.initials, b1.prime_photo, b1.photo_exists, time, seqno, pen_ind, a.player_id, b.player_id_spell1 "
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
				
				if rsgoals.Fields("photo_exists") = "Y" and rsgoals.Fields("player_id_spell1") < 9000 then 
				
					playerid = rsgoals.Fields("player_id_spell1")
					if len(playerid) < 3 then playerid = right("00" & playerid,3)
					
					if isnull(rsgoals.Fields("prime_photo")) then 
						photoid = playerid
				  	  else
				  		photoid = playerid & "_" & rsgoals.Fields("prime_photo")
				  	end if	
			
					'only add to scorer list if not already there
					if instr(photolist_scorers,photoid & ",") = 0 then photolist_scorers = photolist_scorers & photoid & "," & rsgoals.Fields("forename") & "," & rsgoals.Fields("surname") & "^"
				
				end if
			
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
			
		  	'The goal list has been prepared, it's format depending on whether there are goal times or not, and penalty and o.g. indicators.
		  	'Whatever the form, each entry is separated buy a comma. Now check if there are consecutive enties that are the same,
		  	'and simplify if possible.
		  	
		  	dim outgoal, simpler_outgoal(12), simpler_goal_entries(12), goalentries, number_word(6)
		  	number_word(2) = "two" : number_word(3) = "three" : number_word(4) = "four" : number_word(5) = "five" : number_word(6) = "six"
		  	
		  	for i = 0 to 12
		  		simpler_goal_entries(i) = 1
		  	next
		  	
		   	outgoal = split(outgoals,", ")
		  	goalentries = Ubound(outgoal)
		  
		  	simpler_outgoal(0) = outgoal(0)		
			i = 0
			j = 1
							  				
			do while j <= goalentries 
				if outgoal(j) = simpler_outgoal(i) then 
					simpler_goal_entries(i) = simpler_goal_entries(i) + 1
					j = j + 1
				  else	
				  	simpler_outgoal(i+1) = outgoal(j)
				  	i = i + 1
				  	j = j + 1
				end if			
			loop
			
			i = 0
			outgoals = ""
			
			do until simpler_outgoal(i) = ""
				outgoals = outgoals & simpler_outgoal(i)
				if simpler_goal_entries(i) > 1 then outgoals = outgoals & " (" & number_word(simpler_goal_entries(i)) & ")"
				outgoals = outgoals & ", "
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
			
			 if matchdate < "1965-08-01" and (rtrim(rsmiles.Fields("type")) = "PS" or rtrim(rsmiles.Fields("type")) = "PSFL" or rtrim(rsmiles.Fields("type")) = "PSFAC") then
			  
			  	'ignore these milestones in the case of 'starts' before 1965 because they will be covered by 'games' milestones
			  
			  else	
			  	
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
				
			  end if
				
			  rsmiles.MoveNext
			  
			Loop
			
			output_milestones = output_milestones & "</ul>"

		end if
				
		rsmiles.close

End Function  'Getmilestones


Function Getmaterial(matchdate)

	 	sql = "select material_type, straight_to_youtube, material_seq, publish_timestamp, publish_by, updateno, material_details1, material_details2, shortname "
		sql = sql & "from event_control a left outer join contributor b on a.publish_by = b.initials "
		sql = sql & "where event_date = '" & matchdate & "' "
		sql = sql & "  and event_type in ('M','V') "
		sql = sql & "  and material_type in ('A','I','P','V','Y') " 
		if phase <> "review" then sql = sql & "  and event_published = 'Y' "  
		sql = sql & "order by material_type, publish_timestamp, material_details1 "

		rs.open sql,conn,1,2
		
		if rs.RecordCount > 0 then 
		
			output_material = "<div id=""material"" class=""" & undofloat_class & """>"
			output_material = output_material & "<p style=""margin: 10px 0 2px;"" class=""style1bold"">Match Material</p>"
			
			audiofound = ""
			argylemediafound = ""
			othervideofound = ""
			updatenolast = ""
			summaryno = 0
			thanks = ""
					
			Do While Not rs.EOF
			
				if rs.Fields("shortname") <> "Steve" then
					if instr(thanks,rs.Fields("shortname")) = 0 then thanks = thanks & rs.Fields("shortname") & ", " 
				end if
									
				if rs.Fields("updateno") <> updatenolast then 
					if updatenolast > "" then output_material = output_material & "</p>"
					updatenolast = rs.Fields("updateno")
					output_material = output_material & "<p style=""margin:0 0 0 -6px; line-height:200%"">"
				end if 
						
				Select Case rs.Fields("material_type")
					Case "A"
						output_material = output_material & "<span style=""white-space: nowrap;""><img class=""audio"" src=""images/mic16.png"">"
						output_material = output_material & "<a href=""soundfiles/" & matchdate & "/" & rs.Fields("material_details1") & """ onclick=""ga('send','event','Download','mp3',this.href);""><span style=""margin-right:12px"">" & rtrim(rs.Fields("material_details2")) & "</span></a></span> "
						audiofound = "y"
					Case "I"	'current system for images
						output_material = output_material & "<span style=""white-space: nowrap;""><img class=""image"" src=""images/camera16.png"">"
						output_material = output_material & "<a href=""photos.asp?parm=" & matchdate & "M" & rs.Fields("material_seq") & """><span style=""margin-right:12px"">" & rs.Fields("material_details1") & "</span></a></span>"
					Case "P"	'old system for images
						output_material = output_material & "<span style=""white-space: nowrap;""><img class=""image"" src=""images/camera16.png"">"
						output_material = output_material & "<a target=""_blank"" href=""" & rs.Fields("material_details1") & """><span style=""margin-right:12px"">" & rs.Fields("material_details2") & "</span></a></span>"
					Case "Y"
						output_material = output_material & "<span style=""white-space: nowrap;""><img class=""video"" src=""images/video16.png"">"
						if rs.Fields("straight_to_youtube") = "Y" then
							output_material = output_material & "<a href=""https://www.youtube.com/watch?v=" & rs.Fields("material_details1") & """>" 
						  else
							output_material = output_material & "<a href=""https://www.youtube.com/embed/" & rs.Fields("material_details1") 
							output_material = output_material & "?rel=0&amp;wmode=transparent"" onclick=""return hs.htmlExpand(this, {objectType: 'iframe'})"" class=""highslide"">"
						end if
						output_material = output_material & "<span style=""margin-right:12px"">" & rtrim(rs.Fields("material_details2")) & "</span></a></span>"
				End Select
				
				rs.MoveNext
			Loop	
	
		end if
		
		rs.close
		
		if updatenolast > "" then
		
			output_material = output_material & "</p>"			'end para for final item of material
			
			if thanks > "" then 
				thanks = "Thanks to today's contributors: " & thanks
				thanks = left(thanks,len(thanks)-2)		'drop final comma and space
      			thanks = StrReverse(thanks)				'three lines to change ...
				thanks = Replace(thanks,",","& ",1,1)	'... the last comma ...
 				thanks = StrReverse(thanks)				'... to ampersand
 				thanks = thanks & ".</p><p class=""style2"" style=""margin:4px 0 0; line-height:120%;"">"
 			end if	
     		output_material = output_material & "<p class=""style2"" style=""margin:4px 0 0; line-height:120%;"">" & thanks 
		   		
   			if audiofound = "y" then
      			output_material = output_material & "Audio material &copy; " 	
      			if matchdate = "2013-09-03" or matchdate = "2013-10-08" then
      				output_material = output_material & "Plymouth Argyle Football Club. Thanks to Plymouth Argyle "
      		  	  else
      				output_material = output_material & "BBC Radio Devon. Thanks to Plymouth Argyle and BBC Radio Devon "
      			end if	
      			output_material = output_material & "for permission to use broadcast extracts. "
      		end if	
 
   			if argylemediafound = "y" then output_material = output_material & "Video courtesy of the argylemedia YouTube channel. " 
      	   	     	   	
      		output_material = output_material & "</p></div>"
      		    				  		
		end if 		 

End Function  'Getmaterial


%>