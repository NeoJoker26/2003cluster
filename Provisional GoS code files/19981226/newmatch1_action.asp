<%@ Language=VBScript %>
<% Option Explicit %>

<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">

<base target="_self">
<link rel="stylesheet" type="text/css" href="../gos2.css">
</head>
<body>

<%
Dim sql, sqlm, sqlmp, sqlmg, sqlme, rowcount, i, j, startnamepart, subnamepart, subsubnamepart, goals, goal, oggoals, oggoal, ogtime, goalnum, goalarray(15,1), goalind, temp0, temp1, temp2
Dim oppogscorers, unusedsubs, opplineup, oppscorers, away, referee
Dim sqlmessage, mailsubject

Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include virtual="/conn_update.inc"--><%

sql = Request.Form("sql") 

if sql = "" then

	oppogscorers = replace(Request.Form("oppogscorers"),"’","'") 	'convert alternative apostrophe to standard apostrophe
	oppogscorers = replace(oppogscorers,"'","''")  					'convert apostrophe to double apostrophe for sql
	unusedsubs = replace(Request.Form("unusedsubs"),"’","'") 		'convert alternative apostrophe to standard apostrophe
	unusedsubs = replace(unusedsubs,"'","''") 						'convert apostrophe to double apostrophe for sql
	opplineup = replace(Request.Form("opplineup"),"’","'") 			'convert alternative apostrophe to standard apostrophe
	opplineup = replace(opplineup,"'","''") 						'convert apostrophe to double apostrophe for sql
	oppscorers = replace(Request.Form("oppscorers"),"’","'") 		'convert alternative apostrophe to standard apostrophe
	oppscorers = replace(oppscorers,"'","''") 						'convert apostrophe to double apostrophe for sql
	oppscorers = replace(oppscorers,";",",") 						'replace a separating semi-colon with a comma
	away = replace(Request.Form("away"),"’","'") 					'convert alternative apostrophe to standard apostrophe
	away = replace(away,"'","''") 									'convert apostrophe to double apostrophe for sql
	referee = replace(Request.Form("referee"),"’","'") 				'convert alternative apostrophe to standard apostrophe
	referee = replace(referee,"'","''") 							'convert apostrophe to double apostrophe for sql
	
	goalind = 0
	
	if oppogscorers > "" then
		oggoals = split(oppogscorers,";")
		oppogscorers = ""
		for each oggoal in oggoals
			temp0 = split(oggoal,"(")
			temp1 = split(temp0(1),")")
			temp2 = split(temp1(0),",")
			for each ogtime in temp2
				goalarray(goalind,0) = trim(ogtime)
				goalarray(goalind,1) = 9000
				goalind = goalind + 1
				oppogscorers = oppogscorers & trim(temp0(0)) & ","		'Rebuild oppogscorers without goal times
			next
		next
		oppogscorers = left(oppogscorers,len(oppogscorers)-1)	'Remove last comma	
	end if
	
	sqlm = "insert into match values("
	sqlm = sqlm & "'" & Request.Form("date") & "',"
	sqlm = sqlm & "'" & Request.Form("ha") & "',"
	sqlm = sqlm & "'" & Request.Form("opposition") & "',"
	sqlm = sqlm & Request.Form("opposition_qual") & ","			'char column but apostrophes sent through in form, because possibly a NULL value
	sqlm = sqlm & "'" & Request.Form("compcode") & "',"
	sqlm = sqlm & Request.Form("subcomp") & ","					'char column but apostrophes sent through in form, because possibly a NULL value 
	sqlm = sqlm & Request.Form("argylegoals") & ","
	sqlm = sqlm & Request.Form("oppositiongoals") & ","
	if lcase(Request.Form("aet")) = "y" then
		sqlm = sqlm & "'" & Request.Form("aet") & "',"
	  else
	 	sqlm = sqlm & "NULL,"
	end if	
	if Request.Form("pensfor") > "" then
		sqlm = sqlm & Request.Form("pensfor") & ","
	  else
	 	sqlm = sqlm & "NULL,"
	end if	
	if Request.Form("pensagainst") > "" then
		sqlm = sqlm & Request.Form("pensagainst") & ","
	  else
	 	sqlm = sqlm & "NULL,"
	end if	
	sqlm = sqlm & replace(Request.Form("attend"),",","") & "," 	'remove possible comma in attendance
	if Request.Form("points") > "" then
		sqlm = sqlm & Request.Form("points") & ","
	  else
	 	sqlm = sqlm & "NULL,"
	end if
	if Request.Form("position") > "" then
		sqlm = sqlm & Request.Form("position") & ","
	  else
	 	sqlm = sqlm & "NULL,"
	end if		
	if Request.Form("oppogscorers") > "" then
		sqlm = sqlm & "'" & oppogscorers & "'," 	
	  else
 		sqlm = sqlm & "NULL,"
	end if	
	sqlm = sqlm & "0,NULL,NULL,NULL,NULL);"
	
	'Check if the match_extra row exists for this set (the match report might already have been lodged): if so, update; if not, insert a the row

	sql = "select count(*) as count "
	sql = sql & "from match_extra " 
	sql = sql & "where date = '" & Request.Form("date") & "' "
	
	rs.open sql,conn,1,2
	rowcount = rs.Fields("count")
	rs.close

	if rowcount > 0 then

		sqlme = "update match_extra set "
		sqlme = sqlme & "non_playing_subs = '" & unusedsubs & "'," 	
		sqlme = sqlme & "opp_team = '" & opplineup & "',"
		if oppscorers > "" then
			sqlme = sqlme & "opp_goals = '" & oppscorers & "'," 	
	  	  else
	 		sqlme = sqlme & "opp_goals = NULL,"
		end if
		sqlme = sqlme & "visitors = '" & away & "'," 			
		sqlme = sqlme & "referee = '" & referee & "' " 				
		sqlme = sqlme & "where date = '" & Request.Form("date") & "' "
 
 	  else
	
		sqlme = "insert into match_extra (date, non_playing_subs, opp_team, opp_goals, visitors, referee) values("
		sqlme = sqlme & "'" & Request.Form("date") & "',"
		sqlme = sqlme & "'" & unusedsubs & "'," 
		sqlme = sqlme & "'" & opplineup & "'," 
		if oppscorers > "" then
			sqlme = sqlme & "'" & oppscorers & "'," 	
		  else
	 		sqlme = sqlme & "NULL,"
		end if
		sqlme = sqlme & "'" & away & "',"			
		sqlme = sqlme & "'" & referee & "'"		
		sqlme = sqlme & ");"	
	
	end if
	
	sqlmp = ""
	
	'Process the players. Note that a starter (start...) might have been substituted (sub...), and
	'that substitute might have been substituted (subsub...). 
	'Further very unlikely levels of substitution are not catered for.
	
	for i = 1 to 11
		'Build a row for each starting player, comprising date | start_id | start_no | start_card | null or sub_id | null | null
		sqlmp = sqlmp & "insert into match_player values("
		sqlmp = sqlmp & "'" & Request.Form("date") & "',"
		startnamepart = split(Request.Form("start" & i),":")
		sqlmp = sqlmp & startnamepart(1) & ","		'the player_id
		sqlmp = sqlmp & i & ","
		if ucase(Request.Form("startcard" & i)) = "R" then
			sqlmp = sqlmp & "'r',"
		  elseif ucase(Request.Form("startcard" & i)) = "Y"	then
			sqlmp = sqlmp & "'y',"
		  else
		  	sqlmp = sqlmp & "NULL,"
		end if
								
		if Request.Form("subtime" & i) = "" then
			'Starting player was not subbed, so finish off his row
			sqlmp = sqlmp & "NULL,NULL,NULL);"
		  else
			'Starting player was subbed, so finish off his row, then create new row for the substitute
			subnamepart = split(Request.Form("sub" & i),":")
			sqlmp = sqlmp & subnamepart(1) & ","		'the player_id of the sub
			sqlmp = sqlmp & "NULL,NULL);"
			'Now build a sub row comprising date | sub_id | start_no | sub_card | null or subsub_id | sub_time | start_id
			sqlmp = sqlmp & "insert into match_player values("	
			sqlmp = sqlmp & "'" & Request.Form("date") & "',"
			sqlmp = sqlmp & subnamepart(1) & ","		'the sub's player_id
			sqlmp = sqlmp & "0,"
			if ucase(Request.Form("subcard" & i)) = "R" then
				sqlmp = sqlmp & "'r',"
		  	  elseif ucase(Request.Form("subcard" & i)) = "Y"	then
				sqlmp = sqlmp & "'y',"
		  	  else
		  		sqlmp = sqlmp & "NULL,"
			end if
			
			if Request.Form("subsubtime" & i) = "" then
				'Substitute was not subbed, so finish off his row
				sqlmp = sqlmp & "NULL,"
				sqlmp = sqlmp & "'" & Request.Form("subtime" & i) & "',"		
				sqlmp = sqlmp & startnamepart(1) & ");"
			  else
			  	'Substitute was subbed, so finish off his row, then create new row for his substitute
				subsubnamepart = split(Request.Form("subsub" & i),":")
				sqlmp = sqlmp & subsubnamepart(1) & ","		'the player_id of the second sub
				sqlmp = sqlmp & "'" & Request.Form("subtime" & i) & "',"		
				sqlmp = sqlmp & startnamepart(1) & ");"		'the player_id of the starter

				'Build a sub-sub row comprising date | sub_id | 0 | subsub_card | null | subsub_time | sub_id
				sqlmp = sqlmp & "insert into match_player values("	
				sqlmp = sqlmp & "'" & Request.Form("date") & "',"
				sqlmp = sqlmp & subsubnamepart(1) & ","		'the second sub's player_id
				sqlmp = sqlmp & "0,"
				if ucase(Request.Form("subsubcard" & i)) = "R" then
					sqlmp = sqlmp & "'r',"
				  elseif ucase(Request.Form("subsubcard" & i)) = "Y"	then
					sqlmp = sqlmp & "'y',"
				  else
					sqlmp = sqlmp & "NULL,"
				end if
				sqlmp = sqlmp & "NULL,"
				sqlmp = sqlmp & "'" & Request.Form("subsubtime" & i) & "',"
				sqlmp = sqlmp & subnamepart(1) & ");"		'the first sub's player_id
			end if
		end if	
		
		'Now collect goals scored by Argyle players
			
		if Request.Form("startgoaltime" & i) > "" then
			goals = split(Request.Form("startgoaltime" & i),",")
			for each goal in goals
				goalarray(goalind,0) = trim(goal)
				goalarray(goalind,1) = startnamepart(1)
				goalind = goalind + 1
			next
		end if
			
		if Request.Form("subgoaltime" & i) > "" then
			goals = split(Request.Form("subgoaltime" & i),",")
			for each goal in goals
				goalarray(goalind,0) = trim(goal)
				goalarray(goalind,1) = subnamepart(1)
				goalind = goalind + 1
			next
		end if
				
		if Request.Form("subsubgoaltime" & i) > "" then
			goals = split(Request.Form("subsubgoaltime" & i),",")
			for each goal in goals
				goalarray(goalind,0) = trim(goal)
				goalarray(goalind,1) = subsubnamepart(1)
				goalind = goalind + 1
			next
		end if
		
	next
		
	'Goals have been collected in goalarray, now sort for right order and insert into match goal
	
	for i = UBound(goalarray) - 1 To 0 Step -1
		if goalarray(i,0) > "" then 
		  for j = 0 to i
    	    if right("00" & replace(goalarray(j,0)," pen",""),3) > right("00" & replace(goalarray(j+1,0)," pen",""),3) then
        	    temp0=goalarray(j+1,0)
        	    temp1=goalarray(j+1,1)
            	goalarray(j+1,0)=goalarray(j,0)
            	goalarray(j+1,1)=goalarray(j,1)
            	goalarray(j,0)=temp0
            	goalarray(j,1)=temp1
        	end if
    	  next
    	end if
   	next

	goalnum = 1
	for i = 0 to UBound(goalarray)
		if goalarray(i,0) > "" then
			sqlmg = sqlmg & "insert into match_goal values("
			sqlmg = sqlmg & "'" & Request.Form("date") & "',"
			sqlmg = sqlmg & goalnum & ","
			sqlmg = sqlmg & goalarray(i,1) & ","
			if right(goalarray(i,0),4) = " pen" then
				sqlmg = sqlmg & left(goalarray(i,0),len(goalarray(i,0))-4) & ","
				sqlmg = sqlmg & "'Y'" 
		  	  else
				sqlmg = sqlmg & goalarray(i,0) & ","
				sqlmg = sqlmg & "NULL"
			end if	
			sqlmg = sqlmg & ");"
			goalnum = goalnum + 1
		end if	
	next
	
	sql = (sqlm & sqlmp & sqlmg & sqlme)
	
	response.write("<div style=""width:800px; margin:72 auto;"">")
	response.write("<p>" & replace(sql,";",";<br>") & "</p>")
	response.write("<form action=""newmatch1_action.asp"" method=""post"" name=""Form1"">")
	response.write("<input type=""hidden"" name=""date"" value=""" & Request.Form("date") & """>")
	response.write("<input type=""hidden"" name=""sql"" value=""" & sql & """>")
	response.write("<input type=""submit"" name=""b1"" value=""Add to GoS-DB"">")
	response.write("</form>")
	response.write("</div>")

  else
  	
	sqlmessage = replace(sql,";",";<br>")
	response.write("<div style=""width:800px; margin:72 auto;"">")
	
	on error resume next
	conn.Execute sql
	if err <> 0 then 
		response.write("<p style=""margin:72px auto"">!!SQL ERROR!!<br><br>" & sqlmessage & "<br><br>Error: " & err.description & "</p>")
		mailsubject = "FAILED"
	  else	
		response.write("<p style=""margin:72px auto 36px"">Match details added successfully</p>")
		response.write("<p style=""margin:36px auto""><a href=""../gosdb-match.asp?date=" & Request.Form("date") & """>View the Match Page</a></p>")
		
		mailsubject = "added"
	end if		
	On Error GoTo 0	
	
	response.write("</div>")

	
	Dim strTo,strFrom,message,subject
	strTo = "match_added@greensonscreen.co.uk"
	strFrom = "match_added@greensonscreen.co.uk"
	subject = "GoS match details " & mailsubject & " for " & Request.Form("date")
	message = sqlmessage

end if	

conn.Close
%>

</body>
</html>