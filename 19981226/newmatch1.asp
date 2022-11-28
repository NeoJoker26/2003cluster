<%@ Language=VBScript %>
<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>Greens on Screen</title>
<base target="_self">
<link rel="stylesheet" type="text/css" href="../gos2.css">

<style>
<!--
input[type=radio] {margin: 1px; padding: 0;}
p {font-size: 11px; text-align: left;}
th {font-size: 11px; text-align: left; font-weight: bold}
td {font-size: 11px; text-align: left;}
input, textarea {font-family: "courier new",serif; font-size: 12px;}

.bold {font-weight: bold;}

.error {margin: 24px 0; font-size: 11px; text-align: center; color: red; font-weight: bold}
-->
</style>

</head>

<body>

<%
Dim scratch, scratcharray, scratchpart, output, errormsg, i, j, x, currentplayer(40,1), totplayers, thisteam(11,6), argyleteam_starter, argyle_scorer(10,1)
Dim match_date, match_opposition, match_opposition_qual, match_ha, match_compcode, match_subcomp, checked1, checked2, checked3
Dim temp1, temp2, temp3, temp4, doing_argyle, argylescore, oppscore, argylegoals, oppgoals, argyleteam, oppteam, argylesubs, attend, away, referee
Dim first_argyle_found, first_opposition_found, argyle_yellow, opposition_yellow, argyle_red, opposition_red

output = ""
errormsg = ""
first_argyle_found = ""
first_opposition_found = ""

scratch = Request.Form("scratch")				'Retrieve the club's match infomation, as pasted into newmatch.asp

scratch = replace(scratch,Chr(10),"^^^")		'End of line character replaces by ^^^ as an easier to detect marker
scratch = replace(scratch,Chr(13),"^^^")
scratch = replace(scratch,";",",")
scratch = replace(scratch,"(4-4-2)","")
scratch = replace(scratch,"(4-5-1)","")
scratch = replace(scratch,"(4-3-3)","")
scratch = replace(scratch,"(4-4-1-1)","")
scratch = replace(scratch," (c)","")
scratch = replace(scratch," (capt)","")
scratch = replace(scratch,"capt, ","")
scratch = replace(scratch," (gk)","")
scratch = replace(scratch," (GK)","")
scratch = replace(scratch,"half-time","HT")
scratch = replace(scratch,"H-T","HT")
scratch = replace(scratch,"’","'")
scratch = replace(scratch,",  ",", ")

for i = 1 to 99		'Remove all squad numbers in the line-ups
	scratch = replace(scratch,", " & i & " ",", ")
	scratch = replace(scratch,": " & i & " ",": ")
	scratch = replace(scratch,"(" & i & " ","(")
next

Dim conn, sql, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include virtual="/conn_read.inc"--><%

	'Get next match
	sql = "select date, opposition, opposition_qual, homeaway, compcode, subcomp " 
	sql = sql & "from season_this a "
	sql = sql & "where not exists (select * from match b where b.date = a.date) "
	sql = sql & "order by date "
	rs.open sql,conn,1,2
	
	if rs.RecordCount > 0 then				
		match_date = rs.Fields("date")
		match_opposition = rs.Fields("opposition")
		if isnull(rs.Fields("opposition_qual")) then 
			match_opposition_qual = "NULL"
		  else 
		  	match_opposition_qual = "'" & rs.Fields("opposition_qual") & "'"
		end if
		match_ha = rs.Fields("homeaway")
		match_compcode = rs.Fields("compcode")
		if isnull(rs.Fields("subcomp")) then 
			match_subcomp = "NULL"
		  else 
		  	match_subcomp = "'" & rs.Fields("subcomp") & "'"
		end if
	  else
		errormsg = "No valid match found"
	end if
	rs.close


scratcharray = split(scratch,"^^^")		'Break up the lines into an array
scratch = ""

for each scratchpart in scratcharray

	scratchpart = scratchpart & " ^^^"		'the space before the ^^^ is important for some of the subsequent splits, and will be removed by trim statements
	
	if left(scratchpart,6) = "Argyle" and first_argyle_found = "" then		'If the line begins 'Argyle', and it hasn't been found before, then this is score line 
		
		first_argyle_found = "Y"
		temp1 = split(scratchpart," ",3)
		argylescore = trim(temp1(1))
		temp2 = split(temp1(2),"^^^")
		argylegoals = trim(temp2(0)) 

	elseif instr(scratchpart,match_opposition) > 0 and first_opposition_found = "" then		'If the line contains the opposition name, and it hasn't been found before, then this is score line 

		first_opposition_found = "Y"
		temp1 = replace(scratchpart,match_opposition,"---")	'to ensure the number of away goals is the second 'word' (avoids a space in the opposition name screwing things) 
		temp2 = split(temp1," ",3)
		oppscore = trim(temp2(1))
		
		if left(trim(temp2(2)),3) <> "^^^" then		'there must be opposition goalscorers listed
			temp3 = split(temp2(2),"^^^")
			temp1 = split(trim(temp3(0)),";")	'break up the opposition goalscorers
			for each temp2 in temp1
				temp2 = trim(temp2)
				temp3 = split(temp2," ",2)			
				oppgoals = oppgoals & temp3(0) & " (" & temp3(1) & "), "
			next
			oppgoals = left(oppgoals,len(oppgoals)-2)	'remove last comma and space
		end if
		
	elseif left(scratchpart,7) = "Argyle " and first_argyle_found = "Y" then		'If the line begins 'Argyle', and it has been found before, then this is a line-up line
		doing_argyle = "Y"
		temp1 = split(scratchpart,":",2)
		if instr(temp1(1),"Substitutes (not used):") > 0 then
			temp2 = split(temp1(1),"Substitutes (not used):")
		   else
		   	temp2 = split(temp1(1),"Substitutes:")
		end if
		argyleteam = trim(temp2(0))		'The first part of the line is the Argyle team
		if right(argyleteam,1) = "." then argyleteam = left(argyleteam,len(argyleteam)-1)	'Drop possible full-stop at the end
		response.write(argyleteam)
		if ubound(temp2) > 0 then
			temp3 = split(temp2(1),"^^^")
			argylesubs = trim(temp3(0))		'Argyle's unused subs
		end if	

	elseif instr(scratchpart,match_opposition) > 0 and first_opposition_found = "Y" then	'Must be the opposition line-up
		doing_argyle = "N"
		temp1 = split(scratchpart,":",2)
		if instr(temp1(1),"Substitutes (not used):") > 0 then
			temp2 = split(temp1(1),"Substitutes (not used):")		'Not interested in the unused sub names for the opposition
		   else
		   	temp2 = split(temp1(1),"Substitutes:")					'Not interested in the unused sub names for the opposition
		end if
		oppteam = trim(temp2(0))		'The opposition line-up
		
	elseif left(scratchpart,8) = "Booked: " then
		for i = 120 to 1 step -1		'Remove all booking times (we don't store them). Doing it in reverses avoids 10 being affected by 1.
			scratchpart = replace(scratchpart," " & i,"")
		next 		
		temp1 = split(mid(scratchpart,9),",")		'Split on a comma
		for each temp2 in temp1
			temp2 = replace(temp2,"^^^","")			'Remove the end of line indicator if it's there 
			temp2 = trim(temp2)
			if doing_argyle = "N" then
					oppteam = replace(oppteam,temp2,temp2 & "[y]")
				  else
				  	argyleteam = replace(argyleteam,temp2,temp2 & "[y]") 
			end if  	
		next

	elseif replace(lcase(left(scratchpart,10)),"-"," ") = "sent off: " and first_opposition_found = "Y" then		'converting all to lower case copes with it being Off or off
		temp1 = split(mid(scratchpart,11) & ",", ",")
		for each temp2 in temp1
			temp2 = trim(temp2)
			if temp2 > "" then 
				temp2 = left(temp2,instr(temp2," ") - 1)
				if doing_argyle = "N" then
					oppteam = replace(oppteam,temp2 & "[y]",temp2)	'remove a '[y]' that might already be there
					oppteam = replace(oppteam,temp2,temp2 & "[r]")
			  	  else
					argyleteam = replace(argyleteam,temp2 & "[y]",temp2)	'remove a '[y]' that might already be there
					argyleteam = replace(argyleteam,temp2,temp2 & "[r]")
			  	end if
			end if
		next
		
	elseif instr(scratchpart,"Attendance:") > 0 then
		temp1 = split(scratchpart,"Attendance:")
		if instr(temp1(1),"(") > 0 then
			temp2 = split(temp1(1),"(")
			attend = trim(temp2(0))
			temp3 = split(temp2(1),"away")
			temp4 = split(temp3(0),")")		' the word 'away' might not have been present, so look for aclosing bracket instead
			away = trim(temp4(0))
		end if

	elseif instr(scratchpart,"Referee:") > 0 then
		temp1 = split(scratchpart,"Referee:")
		temp2 = split(temp1(1),"^^^")
		referee = trim(temp2(0))
		if right(referee,1) = "." then referee = left(referee,len(referee)-1)
	
	end if 
	
	if left(scratchpart,9) <> "Read more" then scratch = scratch & scratchpart

next

scratch = "<p style=""margin: 0 0 6px"">" & scratch
scratch = replace(scratch,"^^^","</p><p style=""margin: 0 0 6px"">")
scratch = scratch & "</p>"

argyleteam_starter = split(argyleteam,",")	'break up argyle starting XI 

temp1 = split(trim(argylegoals),";")	'break up the argyle goalscorers
i = 0
for each temp2 in temp1
	temp2 = trim(temp2)
	temp3 = split(temp2," ",2)			
	argyle_scorer(i,0) = temp3(0) & " (" & temp3(1) & ")"
	i = i + 1
next


if ubound(argyleteam_starter) <> 10 then

	scratch = scratch & "<p class=""error"">Not 10 commas in Argyle's team</p>"
	
  else
    
    i = 0
    
    for each temp1 in argyleteam_starter
    
    	temp2 = split(temp1," (",2) 
    	temp2(0) = trim(temp2(0)) 

    	if right(temp2(0),3) = "[y]" then 
    		thisteam(i,0) = left(temp2(0),len(temp2(0))-3)
    		thisteam(i,1) = "y"
   		  elseif right(temp2(0),3) = "[r]" then 
    		thisteam(i,0) = left(temp2(0),len(temp2(0))-3)
    		thisteam(i,1) = "r"
   		  else  
    		thisteam(i,0) = temp2(0) 
    	end if
    	
    	if ubound(temp2) > 0 then 
    		
    		temp3 = split(temp2(1),")")
    		x = instrrev(temp3(0)," ")			'look for first space from the right (should precede the substitution time time)
    		thisteam(i,3) = mid(temp3(0),x+1)	'substitution time
    		temp3(0) = left(temp3(0),x-1)		'remove the substitution time to leave the substitute's name only
    		    		    	
    		if right(temp3(0),3) = "[y]" then 
    			thisteam(i,4) = left(temp3(0),len(temp3(0))-3)
    			thisteam(i,5) = "y"
   		  	  elseif right(temp3(0),3) = "[r]" then 
    			thisteam(i,4) = left(temp3(0),len(temp3(0))-3)
    			thisteam(i,5) = "r"
   		  	  else  
    			thisteam(i,4) = temp3(0)
    		end if 
    		 
    	end if
    	
		i = i + 1	
    		  
    next
    
    
	' Get all players in squad
	sql = "select rtrim(left(a.surname,8)) + ':' + cast(a.player_id as varchar) as player, rtrim(forename) + ' ' + rtrim(a.surname) as player_name "
	sql = sql & "from player a " 
	sql = sql & "where last_game_year = 9999 "
	sql = sql & " order by 1 "

	rs.open sql,conn,1,2

		i = 0
					
   		Do While Not rs.EOF
   		
			currentplayer(i,0) = rs.Fields("player")
			currentplayer(i,1) = rs.Fields("player_name")
			i = i + 1
			rs.MoveNext
		
		Loop

		rs.close
		
		totplayers = i
	
	
		output = output & "<input type=""hidden"" name=""date"" value=""" & match_date & """>"
		output = output & "<input type=""hidden"" name=""opposition"" value=""" & match_opposition & """>"
		output = output & "<input type=""hidden"" name=""opposition_qual"" value=""" & match_opposition_qual & """>"
		output = output & "<input type=""hidden"" name=""ha"" value=""" & match_ha & """>"
		output = output & "<input type=""hidden"" name=""compcode"" value=""" & match_compcode & """>"
		output = output & "<input type=""hidden"" name=""subcomp"" value=""" & match_subcomp & """>"
		
		output = output & "<p style=""margin:0 0 6px 0""><span class=""bold"">" & match_date & "</span><span style=""margin-left: 20px"">"		
		if match_ha = "H" then
			output = output & "<span style=""margin-left: 20px"" class=""bold"">Argyle </span><input type=""text"" name=""argylegoals"" size=""2"" value=""" & argylescore & """>"
			output = output & "<span style=""margin-left: 10px"" class=""bold"">" & match_opposition & " </span><input type=""text"" name=""oppositiongoals"" size=""2"" value=""" & oppscore & """>"

		  else
			output = output & "<span style=""margin-left: 20px"" class=""bold"">" & match_opposition & " </span><input type=""text"" name=""oppositiongoals"" size=""2"" value=""" & oppscore & """>"
			output = output & "<span style=""margin-left: 10px"" class=""bold"">Argyle </span><input type=""text"" name=""argylegoals"" size=""2"" value=""" & argylescore & """>"
		end if
		output = output & "<span style=""margin-left: 20px"" class=""bold"">AET? </span><select name=""aet"" size=""1""><option>N</option><option>Y</option></select>"
		output = output & "<span style=""margin-left: 20px"" class=""bold"">Attend </span><input type=""text"" name=""attend"" size=""7"" value=""" & attend & """>"
		output = output & "<span style=""margin-left: 20px"" class=""bold"">Away </span><input type=""text"" name=""away"" size=""5"" value=""" & away & """>"
		output = output & "<span style=""margin-left: 20px"" class=""bold"">Referee </span><input type=""text"" name=""referee"" size=""18"" value=""" & referee & """>"
		output = output & "</p>"

		output = output & "<p style=""margin:0 0 6px 0""><span class=""bold"">Total points </span><input type=""text"" name=""points"" size=""2"">"
		output = output & "<span style=""margin-left: 20px"" class=""bold"">Position </span><input type=""text"" name=""position"" size=""2"">"
		output = output & "<span style=""margin-left: 20px"" class=""bold"">Pens for </span><input type=""text"" name=""pensfor"" size=""2"">"
		output = output & "<span style=""margin-left: 20px"" class=""bold"">Pens against </span><input type=""text"" name=""pensagainst"" size=""2"">"	
		output = output & " (fill in Pens when cup-tie settled by penalties)"	
		output = output & "</p>"
		
		output = output & "<p style=""margin:0 0 6px 0""><span class=""bold"">Argyle unused subs </span><input type=""text"" name=""unusedsubs"" size=""80"" value=""" & argylesubs & """></p>"
		
		output = output & "<p style=""margin:0 0 6px 0""><span class=""bold"">Opposition</span><textarea style=""vertical-align:text-top; margin-left:6px;"" name=""opplineup"" rows=""3"" cols=""100"">" & oppteam & "</textarea></p>"
				
		output = output & "<p style=""margin:0 0 6px 0""><span class=""bold"">Opp scorers<sup>*1</sup></span><input style=""margin-left:6px"" type=""text"" name=""oppscorers"" size=""40"" value=""" & oppgoals & """>"
		output = output & "<span style=""margin-left: 20px"" class=""bold"">Opp og scorers<sup>*1</sup></span><input style=""margin-left:6px"" type=""text"" name=""oppogscorers"" size=""30"">"
		output = output & "</p>"
					
		output = output & "<table border=""0"" style=""border-collapse: collapse; margin: 12 auto;"" width=""100%"">"
		output = output & "<tr><th>Starter</th><th>No-Y-R</th><th>Goal times<sup>*2</sup></th><th>Sub</th><th>Time</th><th>No-Y-R</th><th>Goal times</th><th>Sub</th><th>Time</th><th>No-Y-R</th><th>Goal times</th></tr>"

		for i = 0 to 10
			output = output & "<tr><td><select size=""1"" name=""start" & i+1 & """>"
			for j = 0 to totplayers-1
				output = output & "<option"
				if currentplayer(j,1) = thisteam(i,0) then output = output & " selected=""selected"""
				output = output & ">" & currentplayer(j,0) & "</option>"
			next
			output = output & "</select></td>"
			
			output = output & "<td>"
			checked1 = ""
			checked2 = ""
			checked3 = ""
			select case thisteam(i,1)
				case "y"
					checked2 = "checked=""checked"""
				case "r"
					checked3 = "checked=""checked"""
			end select
			output = output & "<input type=""radio"" value=""N"" " & checked1 & " name=""startcard" & i+1 & """>"
			output = output & "<input type=""radio"" value=""Y"" " & checked2 & " name=""startcard" & i+1 & """>"
			output = output & "<input type=""radio"" value=""R"" " & checked3 & " name=""startcard" & i+1 & """>"
			output = output & "</td>"
			
			output = output & "<td>"
  			output = output & "<input type=""text"" name=""startgoaltime" & i+1 & """size=""10"">"
			output = output & "</td>"
			
			output = output & "<td><select size=""1"" name=""sub" & i+1 & """>"
			output = output & "<option>Not subbed</option>"
			for j = 0 to totplayers-1
				output = output & "<option"
				if currentplayer(j,1) = thisteam(i,4) then output = output & " selected"
				output = output & ">" & currentplayer(j,0) & "</option>"
			next
			output = output & "</select></td>"
			
			output = output & "<td>"
			output = output & "<input type=""text"" name=""subtime" & i+1 & """size=""2"" value=""" & thisteam(i,3) & """>" 
			output = output & "</td>"
			
			output = output & "<td>"
			checked1 = ""
			checked2 = ""
			checked3 = ""
			select case thisteam(i,5)
				case "y"
					checked2 = "checked"
				case "r"
					checked3 = "checked"
			end select
			output = output & "<input type=""radio"" value=""N"" " & checked1 & " name=""subcard" & i+1 & """>"
			output = output & "<input type=""radio"" value=""Y"" " & checked2 & " name=""subcard" & i+1 & """>"
			output = output & "<input type=""radio"" value=""R"" " & checked3 & " name=""subcard" & i+1 & """>"
			output = output & "</td>"
			
			output = output & "<td>"
  			output = output & "<input type=""text"" name=""subgoaltime" & i+1 & """size=""5"">"
			output = output & "</td>"		
			
			' This following caters for a sub of a sub

			output = output & "<td><select size=""1"" name=""subsub" & i+1 & """>"
			output = output & "<option>Was sub subbed?</option>"
			for j = 0 to totplayers-1
				output = output & "<option"
				output = output & ">" & currentplayer(j,0) & "</option>"
			next
			output = output & "</select></td>"
			
			output = output & "<td>"
			output = output & "<input type=""text"" name=""subsubtime" & i+1 & """size=""2"">" 
			output = output & "</td>"
			
			output = output & "<td>"
			checked1 = ""
			checked2 = ""
			checked3 = ""
			select case thisteam(i,5)
				case "y"
					checked2 = "checked"
				case "r"
					checked3 = "checked"
			end select
			output = output & "<input type=""radio"" value=""N"" " & checked1 & " name=""subsubcard" & i+1 & """>"
			output = output & "<input type=""radio"" value=""Y"" " & checked2 & " name=""subsubcard" & i+1 & """>"
			output = output & "<input type=""radio"" value=""R"" " & checked3 & " name=""subsubcard" & i+1 & """>"
			output = output & "</td>"
			
			output = output & "<td>"
  			output = output & "<input type=""text"" name=""subsubgoaltime" & i+1 & """size=""4"">"
			output = output & "</td>"
			
			output = output & "</tr>"
		
		next
		output = output & "</table>"
		output = output & "<p style=""margin:0 0 6px 6px;"">*1 Separate players with a semicolon, e.g. Smith (67 pen); Jones (75,85)</p>"
		output = output & "<p style=""margin:0 0 6px 6px;"">*2 Example: show a penalty as 12 pen, 65</p>"
		output = output & "<input style=""margin: 18px 0;"" type=""submit"" name=""b1"" value=""Add to Database"">"
		
	end if		
	
%>

	<div style="width:980px; margin:10px auto; text-align:left;">
	<div style="margin-bottom: 10px; padding: 6px; border: 1px solid #404040;">
	<%response.write(scratch)%>	
	</div>
	<form action="newmatch1_action.asp" method="post" name="Form1">
	<%response.write(output)%>
	</form>
	</div>	

</body>
</html>