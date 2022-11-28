<%@ Language=VBScript %> 
<% Option Explicit %>
<!DOCTYPE html>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GoS Admin</title>
<link rel="stylesheet" type="text/css" href="../gos2.css">
<style>
<!--
#container {
	font-size:11px; 
	text-align:left; 
	width:fit-content;
	margin:24px auto;
	}
-->
</style>
</head>

<body>

<% 
Dim i, output, phase, administrator
Dim goalsfor, goalsagainst, matches(12), starts(12), hmatches(12), hstarts(12), amatches(12), astarts(12), a, b, c, results_count, appears_count
Dim player_array1(1500,10), player_array2(1500,9), prevdate, l_prevdate, time1, time2, time3
Dim r_date, r_lfc, r_player_id_spell1, r_homeaway, r_goalsfor, r_goalsagainst, r_startpos, r_goals

Dim conn, sql, rs, row, rows, rows_eof, r
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_admin.inc"--><%
%>

<div id="container">
<!--#include file="admin_head.inc"-->

<h3 style="margin:6px 0 15px;">UPDATE CONSECUTIVE RESULTS AND APPEARANCES</h3>

<% 

phase = request.form("phase")

select case phase
	case 1
		Call Update
	case else
		Call Display
end select

Response.write(output)
Conn.close


Sub Display

	output = "<form action=""consecutive_refresh.asp"" method=""post"">"
	output = output & "<input type=""hidden"" name=""phase"" value=""1"">"	
	output = output & "<input style=""margin:15px 0;"" type=""submit"" value=""Continue"">"
	output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
	output = output & "</form>"

End Sub


Sub Update

matches(1) = 0
matches(2) = 0
matches(3) = 0
matches(4) = 0
matches(5) = 0
matches(6) = 0
matches(7) = 0
matches(8) = 0
matches(9) = 0
matches(10) = 0
matches(11) = 0
matches(12) = 0

hmatches(1) = 0
hmatches(2) = 0
hmatches(3) = 0
hmatches(4) = 0
hmatches(5) = 0
hmatches(6) = 0
hmatches(7) = 0
hmatches(8) = 0
hmatches(9) = 0
hmatches(10) = 0
hmatches(11) = 0
hmatches(12) = 0

amatches(1) = 0
amatches(2) = 0
amatches(3) = 0
amatches(4) = 0
amatches(5) = 0
amatches(6) = 0
amatches(7) = 0
amatches(8) = 0
amatches(9) = 0
amatches(10) = 0
amatches(11) = 0
amatches(12) = 0

'*** Prepare and load consecutive results

time1 = timer()

sql = "delete from consecutive_results"
conn.execute sql

results_count = 0

r_date = 0
r_homeaway = 1
r_lfc = 2
r_goalsfor = 3
r_goalsagainst = 4
rows_eof = true

sql = "select date, homeaway, lfc, goalsfor, goalsagainst from v_match_all "
sql = sql & " order by date"

set rs = conn.execute(sql)

if not rs.EOF then
	rows = rs.getrows()
	rows_eof = false
end if
rs.close

sql = ""

if not rows_eof then

 for row = 0 to UBound(rows,2)

  	goalsfor = rows(r_goalsfor,row)
  	goalsagainst = rows(r_goalsagainst,row)
 		

	if goalsfor > goalsagainst then 
		a = 1
		b = 2
		c = 3
	  elseif goalsfor = goalsagainst then 
		a = 2
		b = 1
		c = 3
	  else
		a = 3
		b = 1
		c = 2
	end if

	matches(a) = matches(a) + 1
	if matches(a) = 1 then starts(a) = "'" & rows(r_date,row) & "'"
	matches(b) = 0
	starts(b) = "NULL" 
	matches(c) = 0
	starts(c) = "NULL"
	
	if a = 1 then 					'a win so update "can't lose"
		matches(4) = matches(4) + 1
		if matches(4) = 1 then starts(4) = "'" & rows(r_date,row) & "'"
		matches(5) = 0	
		starts(5) = "NULL"
	  
	  elseif a = 2 then 			'a draw so update "can't lose"
	 	matches(4) = matches(4) + 1
		if matches(4) = 1 then starts(4) = "'" & rows(r_date,row) & "'"
		matches(5) = matches(5) + 1
		if matches(5) = 1 then starts(5) = "'" & rows(r_date,row) & "'"
	  
	  else  						'a defeat so update "can't win"
		matches(5) = matches(5) + 1
		if matches(5) = 1 then starts(5) = "'" & rows(r_date,row) & "'"
		matches(4) = 0	
		starts(4) = "NULL"
	end if
	
	if goalsagainst = 0 then 		'clean sheet
		matches(6) = matches(6) + 1
		if matches(6) = 1 then starts(6) = "'" & rows(r_date,row) & "'"
	  else
	  	matches(6) = 0
	  	starts(6) = "NULL"
	end if
			 
	if rows(r_lfc,row) = "F" then
		matches(a+6) = matches(a+6) + 1
		if matches(a+6) = 1 then starts(a+6) = "'" & rows(r_date,row) & "'" 
		matches(b+6) = 0
		starts(b+6) = "NULL" 
		matches(c+6) = 0
		starts(c+6) = "NULL"
				
		if a = 1 then 					'a win so update "can't lose"
			matches(10) = matches(10) + 1
			if matches(10) = 1 then starts(10) = "'" & rows(r_date,row) & "'"
			matches(11) = 0	
			starts(11) = "NULL"
	  
	  	elseif a = 2 then 			'a draw so update "can't lose" and "can't win"
	 		matches(10) = matches(10) + 1
			if matches(10) = 1 then starts(10) = "'" & rows(r_date,row) & "'"
			matches(11) = matches(11) + 1
			if matches(11) = 1 then starts(11) = "'" & rows(r_date,row) & "'"
	  
	  	else  						'a defeat so update "can't win"
			matches(11) = matches(11) + 1
			if matches(11) = 1 then starts(11) = "'" & rows(r_date,row) & "'"
			matches(10) = 0	
			starts(10) = "NULL"
		end if
		
		if goalsagainst = 0 then 		'clean sheet
			matches(12) = matches(12) + 1
			if matches(12) = 1 then starts(12) = "'" & rows(r_date,row) & "'"
	  	  else
	  		matches(12) = 0
	  		starts(12) = "NULL"
		end if
		
		sql = sql & "insert into consecutive_results values("
		sql = sql & "'" & rows(r_date,row) & "',"
		sql = sql & "' ',"
		sql = sql & matches(1) & ","
		sql = sql & matches(2) & ","
		sql = sql & matches(3) & ","
		sql = sql & matches(4) & ","
		sql = sql & matches(5) & ","
		sql = sql & matches(6) & ","
		sql = sql & matches(7) & ","
		sql = sql & matches(8) & ","
		sql = sql & matches(9) & ","
		sql = sql & matches(10) & ","
		sql = sql & matches(11) & ","
		sql = sql & matches(12) & ","
		sql = sql & starts(1) & ","
		sql = sql & starts(2) & ","
		sql = sql & starts(3) & ","
		sql = sql & starts(4) & ","
		sql = sql & starts(5) & ","
		sql = sql & starts(6) & ","
		sql = sql & starts(7) & ","
		sql = sql & starts(8) & ","
		sql = sql & starts(9) & ","
		sql = sql & starts(10) & ","
		sql = sql & starts(11) & ","
		sql = sql & starts(12) & "), "
		
	  else				'Not a Football League game, so not interested in league figures
	  
		sql = sql & "insert into consecutive_results values("
		sql = sql & "'" & rows(r_date,row) & "',"
		sql = sql & "' ',"
		sql = sql & matches(1) & ","
		sql = sql & matches(2) & ","
		sql = sql & matches(3) & ","
		sql = sql & matches(4) & ","
		sql = sql & matches(5) & ","
		sql = sql & matches(6) & ","
		sql = sql & "NULL,"
		sql = sql & "NULL,"
		sql = sql & "NULL,"
		sql = sql & "NULL,"
		sql = sql & "NULL,"
		sql = sql & "NULL,"
		sql = sql & starts(1) & ","
		sql = sql & starts(2) & ","
		sql = sql & starts(3) & ","
		sql = sql & starts(4) & ","
		sql = sql & starts(5) & ","
		sql = sql & starts(6) & ","
		sql = sql & "NULL,"
		sql = sql & "NULL,"
		sql = sql & "NULL,"
		sql = sql & "NULL,"
		sql = sql & "NULL,"
		sql = sql & "NULL), "
	
	end if

	results_count = results_count + 1

	if rows(r_homeaway,row) = "H" then 

		hmatches(a) = hmatches(a) + 1
		if hmatches(a) = 1 then hstarts(a) = "'" & rows(r_date,row) & "'"
		hmatches(b) = 0
		hstarts(b) = "NULL" 
		hmatches(c) = 0
		hstarts(c) = "NULL"
		
		if a = 1 then 					'a win so update "can't lose"
			hmatches(4) = hmatches(4) + 1
			if hmatches(4) = 1 then hstarts(4) = "'" & rows(r_date,row) & "'"
			hmatches(5) = 0	
			hstarts(5) = "NULL"
	  
	  	  elseif a = 2 then 			'a draw so update "can't lose" and "can't win"
	 		hmatches(4) = hmatches(4) + 1
			if hmatches(4) = 1 then hstarts(4) = "'" & rows(r_date,row) & "'"
			hmatches(5) = hmatches(5) + 1
			if hmatches(5) = 1 then hstarts(5) = "'" & rows(r_date,row) & "'"
	  
	  	  else  						'a defeat so update "can't win"
			hmatches(5) = hmatches(5) + 1
			if hmatches(5) = 1 then hstarts(5) = "'" & rows(r_date,row) & "'"
			hmatches(4) = 0	
			hstarts(4) = "NULL"
		end if
	
		if goalsagainst = 0 then 		'clean sheet
			hmatches(6) = hmatches(6) + 1
			if hmatches(6) = 1 then hstarts(6) = "'" & rows(r_date,row) & "'"
	  	  else
	  		hmatches(6) = 0
	  		hstarts(6) = "NULL"
		end if
		
		if rows(r_lfc,row) = "F" then
			hmatches(a+6) = hmatches(a+6) + 1
			if hmatches(a+6) = 1 then hstarts(a+6) = "'" & rows(r_date,row) & "'" 
			hmatches(b+6) = 0
			hstarts(b+6) = "NULL" 
			hmatches(c+6) = 0
			hstarts(c+6) = "NULL"
				
			if a = 1 then 					'a win so update "can't lose"
				hmatches(10) = hmatches(10) + 1
				if hmatches(10) = 1 then hstarts(10) = "'" & rows(r_date,row) & "'"
				hmatches(11) = 0	
				hstarts(11) = "NULL"
	  
	  		  elseif a = 2 then 			'a draw so update "can't lose"
	 			hmatches(10) = hmatches(10) + 1
				if hmatches(10) = 1 then hstarts(10) = "'" & rows(r_date,row) & "'"
				hmatches(11) = hmatches(11) + 1
				if hmatches(11) = 1 then hstarts(11) = "'" & rows(r_date,row) & "'"
	  
	  		  else  						'a defeat so update "can't win"
				hmatches(11) = hmatches(11) + 1
				if hmatches(11) = 1 then hstarts(11) = "'" & rows(r_date,row) & "'"
				hmatches(10) = 0	
				hstarts(10) = "NULL"
			end if
					
			if goalsagainst = 0 then 		'clean sheet
				hmatches(12) = hmatches(12) + 1
				if hmatches(12) = 1 then hstarts(12) = "'" & rows(r_date,row) & "'"
	  	  	  else
	  			hmatches(12) = 0
	  			hstarts(12) = "NULL"
			end if
			
			sql = sql & "("
			sql = sql & "'" & rows(r_date,row) & "',"
			sql = sql & "'H',"
			sql = sql & hmatches(1) & ","
			sql = sql & hmatches(2) & ","
			sql = sql & hmatches(3) & ","
			sql = sql & hmatches(4) & ","
			sql = sql & hmatches(5) & ","
			sql = sql & hmatches(6) & ","
			sql = sql & hmatches(7) & ","
			sql = sql & hmatches(8) & ","
			sql = sql & hmatches(9) & ","
			sql = sql & hmatches(10) & ","
			sql = sql & hmatches(11) & ","
			sql = sql & hmatches(12) & ","
			sql = sql & hstarts(1) & ","
			sql = sql & hstarts(2) & ","
			sql = sql & hstarts(3) & ","
			sql = sql & hstarts(4) & ","
			sql = sql & hstarts(5) & ","
			sql = sql & hstarts(6) & ","
			sql = sql & hstarts(7) & ","
			sql = sql & hstarts(8) & ","
			sql = sql & hstarts(9) & ","
			sql = sql & hstarts(10) & ","
			sql = sql & hstarts(11) & ","
			sql = sql & hstarts(12) & "); "
			
	  	  else				'Not a Football League game, so not interested in league figures
	  	  
			sql = sql & "("
			sql = sql & "'" & rows(r_date,row) & "',"
			sql = sql & "'H',"
			sql = sql & hmatches(1) & ","
			sql = sql & hmatches(2) & ","
			sql = sql & hmatches(3) & ","
			sql = sql & hmatches(4) & ","
			sql = sql & hmatches(5) & ","
			sql = sql & hmatches(6) & ","
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & hstarts(1) & ","
			sql = sql & hstarts(2) & ","
			sql = sql & hstarts(3) & ","
			sql = sql & hstarts(4) & ","
			sql = sql & hstarts(5) & ","
			sql = sql & hstarts(6) & ","			
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL); "
			
		end if
  
	  else
		
		amatches(a) = amatches(a) + 1
		if amatches(a) = 1 then astarts(a) = "'" & rows(r_date,row) & "'"
		amatches(b) = 0
		astarts(b) = "NULL" 
		amatches(c) = 0
		astarts(c) = "NULL"
		
		if a = 1 then 					'a win so update "can't lose"
			amatches(4) = amatches(4) + 1
			if amatches(4) = 1 then astarts(4) = "'" & rows(r_date,row) & "'"
			amatches(5) = 0	
			astarts(5) = "NULL"
	  
	  	  elseif a = 2 then 			'a draw so update "can't lose" and "can't win"
	 		amatches(4) = amatches(4) + 1
			if amatches(4) = 1 then astarts(4) = "'" & rows(r_date,row) & "'"
			amatches(5) = amatches(5) + 1
			if amatches(5) = 1 then astarts(5) = "'" & rows(r_date,row) & "'"
	  
	  	  else  						'a defeat so update "can't win"
			amatches(5) = amatches(5) + 1
			if amatches(5) = 1 then astarts(5) = "'" & rows(r_date,row) & "'"
			amatches(4) = 0	
			astarts(4) = "NULL"
		end if
	
		if goalsagainst = 0 then 		'clean sheet
			amatches(6) = amatches(6) + 1
			if amatches(6) = 1 then astarts(6) = "'" & rows(r_date,row) & "'"
	  	  else
	  		amatches(6) = 0
	  		astarts(6) = "NULL"
		end if
		
		if rows(r_lfc,row) = "F" then
			amatches(a+6) = amatches(a+6) + 1
			if amatches(a+6) = 1 then astarts(a+6) = "'" & rows(r_date,row) & "'" 
			amatches(b+6) = 0
			astarts(b+6) = "NULL" 
			amatches(c+6) = 0
			astarts(c+6) = "NULL"
				
			if a = 1 then 					'a win so update "can't lose"
				amatches(10) = amatches(10) + 1
				if amatches(10) = 1 then astarts(10) = "'" & rows(r_date,row) & "'"
				amatches(11) = 0	
				astarts(11) = "NULL"
	  
	  		  elseif a = 2 then 			'a draw so update "can't lose"
	 			amatches(10) = amatches(10) + 1
				if amatches(10) = 1 then astarts(10) = "'" & rows(r_date,row) & "'"
				amatches(11) = amatches(11) + 1
				if amatches(11) = 1 then astarts(11) = "'" & rows(r_date,row) & "'"
	  
	  		  else  						'a defeat so update "can't win"
				amatches(11) = amatches(11) + 1
				if amatches(11) = 1 then astarts(11) = "'" & rows(r_date,row) & "'"
				amatches(10) = 0	
				astarts(10) = "NULL"
			end if
					
			if goalsagainst = 0 then 		'clean sheet
				amatches(12) = amatches(12) + 1
				if amatches(12) = 1 then astarts(12) = "'" & rows(r_date,row) & "'"
	  	  	  else
	  			amatches(12) = 0
	  			astarts(12) = "NULL"
			end if
			
			sql = sql & "("
			sql = sql & "'" & rows(r_date,row) & "',"
			sql = sql & "'A',"
			sql = sql & amatches(1) & ","
			sql = sql & amatches(2) & ","
			sql = sql & amatches(3) & ","
			sql = sql & amatches(4) & ","
			sql = sql & amatches(5) & ","
			sql = sql & amatches(6) & ","
			sql = sql & amatches(7) & ","
			sql = sql & amatches(8) & ","
			sql = sql & amatches(9) & ","
			sql = sql & amatches(10) & ","
			sql = sql & amatches(11) & ","
			sql = sql & amatches(12) & ","
			sql = sql & astarts(1) & ","
			sql = sql & astarts(2) & ","
			sql = sql & astarts(3) & ","
			sql = sql & astarts(4) & ","
			sql = sql & astarts(5) & ","
			sql = sql & astarts(6) & ","
			sql = sql & astarts(7) & ","
			sql = sql & astarts(8) & ","
			sql = sql & astarts(9) & ","
			sql = sql & astarts(10) & ","
			sql = sql & astarts(11) & ","
			sql = sql & astarts(12) & "); "
			
	  	  else				'Not a Football League game, so not interested in league figures
	  	  
			sql = sql & "("
			sql = sql & "'" & rows(r_date,row) & "',"
			sql = sql & "'A',"
			sql = sql & amatches(1) & ","
			sql = sql & amatches(2) & ","
			sql = sql & amatches(3) & ","
			sql = sql & amatches(4) & ","
			sql = sql & amatches(5) & ","
			sql = sql & amatches(6) & ","
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & astarts(1) & ","
			sql = sql & astarts(2) & ","
			sql = sql & astarts(3) & ","
			sql = sql & astarts(4) & ","
			sql = sql & astarts(5) & ","
			sql = sql & astarts(6) & ","
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL,"
			sql = sql & "NULL); "
			
		end if
		
	end if
	
	results_count = results_count + 1
	
		
	if i mod 100 = 0 and sql > "" then
		conn.execute sql	
		sql = ""
	end if

next
end if

time2 = timer()


'*** Prepare and load consecutive appearances, but in three stages
'*** First, prepare for all three

for i = 1 to 1500
	player_array1(i,0) = 0
	player_array1(i,1) = ""
	player_array1(i,2) = ""
	player_array1(i,3) = 0
	player_array1(i,4) = ""
	player_array1(i,5) = ""
	player_array1(i,6) = 0
	player_array1(i,7) = 0
	player_array1(i,8) = ""
	player_array1(i,9) = ""
	player_array2(i,0) = 0
	player_array2(i,1) = ""
	player_array2(i,2) = ""
	player_array2(i,3) = 0
	player_array2(i,4) = ""
	player_array2(i,5) = ""
	player_array2(i,6) = 0
	player_array2(i,7) = 0
	player_array2(i,8) = ""
	player_array2(i,9) = ""
next

sql = "delete from consecutive_appears"
conn.execute sql

r_date = 0
r_player_id_spell1 = 1
r_startpos = 2
r_goals = 3


	'** Now process for (1) consecutive starts in all competitons
	'** (array slots 0,1,2) 

	prevdate = ""
	rows_eof = true

	sql = "select a.date, player_id_spell1, startpos "
	sql = sql & "from v_match_all a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where b.player_id <> 8000 "
	sql = sql & "  and startpos > 0 "
	sql = sql & "order by a.date, startpos"

	set rs = conn.execute(sql)

	if not rs.EOF then
		rows = rs.getrows()
		rows_eof = false
	end if
	rs.close

	if not rows_eof then

		for row = 0 to UBound(rows,2)
	
			if rows(r_date,row) <> prevdate then Call ProcessPrevMatch_1		'Next match detected	
	   			
	   		player_array1(rows(r_player_id_spell1,row),0) = player_array1(rows(r_player_id_spell1,row),0) + 1								'increment sequence count for starters (all matches)
	   		if player_array1(rows(r_player_id_spell1,row),1) = "" then player_array1(rows(r_player_id_spell1,row),1) = rows(r_date,row)		'a new start date for a sequence (all matches)
	   		player_array1(rows(r_player_id_spell1,row),2) = rows(r_date,row)																'latest date for end of sequence (all matches)

		next

	end if

	'Process the players in the final match
	
	for i = 1 to 1500
		if player_array1(i,0) > 0 then
			if player_array1(i,0) > player_array2(i,0) then
				player_array2(i,0) = player_array1(i,0)
				player_array2(i,1) = player_array1(i,1)
				player_array2(i,2) = player_array1(i,2)
			end if	
		end if
	next
	
	'** Now process for (2) consecutive starts in all league matches
	'** (array slots 3,4,5) 

	prevdate = ""
	rows_eof = true

	sql = "select a.date, player_id_spell1, startpos "
	sql = sql & "from v_match_all a join match_player b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where b.player_id <> 8000 "
	sql = sql & "  and startpos > 0 "
	sql = sql & "  and lfc <> 'c' "
	sql = sql & "order by a.date, startpos"

	set rs = conn.execute(sql)

	if not rs.EOF then
		rows = rs.getrows()
		rows_eof = false
	end if
	rs.close

	if not rows_eof then

		for row = 0 to UBound(rows,2)
	
			if rows(r_date,row) <> prevdate then Call ProcessPrevMatch_2		'Next match detected	
			
	   		player_array1(rows(r_player_id_spell1,row),3) = player_array1(rows(r_player_id_spell1,row),3) + 1								'increment sequence count for starters (league matches)
	   		if player_array1(rows(r_player_id_spell1,row),4) = "" then player_array1(rows(r_player_id_spell1,row),4) = rows(r_date,row)		'a new start date for a sequence (league matches)
	   		player_array1(rows(r_player_id_spell1,row),5) = rows(r_date,row)																'latest date for end of sequence (league matches)

		next

	end if

	'Process the players in the final match
	
	for i = 1 to 1500
		if player_array1(i,3) > 0 then
			if player_array1(i,3) > player_array2(i,3) then
				player_array2(i,3) = player_array1(i,3)
				player_array2(i,4) = player_array1(i,4)
				player_array2(i,5) = player_array1(i,5)
			end if
		end if
	next
	
	'** Now process for (3) consecutive goals scored in all competitons
	'** (array slots 6,7,8,9) 

	prevdate = ""
	rows_eof = true

	sql = "with CTE as ( "
	sql = sql & "select date, player_id, count(*) as goals "
	sql = sql & "from match_goal "
	sql = sql & "group by date, player_id "
	sql = sql & ") "
	sql = sql & "select a.date, player_id_spell1, startpos, isnull(goals,0) "
	sql = sql & "from v_match_all a join match_player b on a.date = b.date left join CTE c on b.date = c.date and b.player_id = c.player_id join player d on b.player_id = d.player_id "
	sql = sql & "where b.player_id <> 8000 "
	sql = sql & "order by a.date, startpos"

	set rs = conn.execute(sql)

	if not rs.EOF then
		rows = rs.getrows()
		rows_eof = false
	end if
	rs.close

	if not rows_eof then

		for row = 0 to UBound(rows,2)
	
			if rows(r_date,row) <> prevdate then Call ProcessPrevMatch_3		'Next match detected	
   		
			if rows(r_goals,row) > 0 then 
				player_array1(rows(r_player_id_spell1,row),6) = player_array1(rows(r_player_id_spell1,row),6) + 1								'increment sequence count for goal-scoring (all matches, even as a sub)
				player_array1(rows(r_player_id_spell1,row),7) = player_array1(rows(r_player_id_spell1,row),7) + rows(r_goals,row)				'increment number of goals in sequence (all matches, even as a sub)
   				if player_array1(rows(r_player_id_spell1,row),8) = "" then player_array1(rows(r_player_id_spell1,row),8) = rows(r_date,row)		'a new start date for a goal scoring sequence (all matches, even as a sub)
   				player_array1(rows(r_player_id_spell1,row),9) = rows(r_date,row)																'latest date for end of goal scoring sequence (all matches, even as a sub)
			end if

		next

	end if

	'Process the players in the final match

	for i = 1 to 1500
		if player_array1(i,6) > 0 then
			if player_array1(i,6) > player_array2(i,6) then
				player_array2(i,6) = player_array1(i,6)
				player_array2(i,7) = player_array1(i,7)
				player_array2(i,8) = player_array1(i,8)
				player_array2(i,9) = player_array1(i,9)
			end if
		end if
	next


'*** Finally, insert new consecutive_appears rows

sql = ""

for i = 1 to 1500
	
	if player_array2(i,0) > 0 then

			sql = sql & "insert into consecutive_appears values("
			sql = sql & i & ","
			sql = sql & player_array2(i,0) & ","
			sql = sql & "'" & player_array2(i,1) & "',"
			sql = sql & "'" & player_array2(i,2) & "', "
			sql = sql & player_array2(i,3) & ", "
			sql = sql & "'" & player_array2(i,4) & "', "
			sql = sql & "'" & player_array2(i,5) & "', "
			sql = sql & player_array2(i,6) & ", "
			sql = sql & player_array2(i,7) & ", "
			sql = sql & "'" & player_array2(i,8) & "', "
			sql = sql & "'" & player_array2(i,9) & "'); "
			
			appears_count = appears_count + 1

	end if
	
	if i mod 50 = 0 and sql > "" then
		conn.execute sql	
		sql = ""
	end if
		
next 

time3 = timer()

output = output & "<p class=""style1bold"" style=""color:green"">consecutive_results has been rebuilt with " & results_count & " rows" & " (" & time2-time1 & " seconds)</p>"
output = output & "<p class=""style1bold"" style=""color:green"">consecutive_appears has been rebuilt with " & appears_count & " rows" & " (" & time3-time2 & " seconds)</p>"

End Sub


Sub Backbutton
	output = output & "<form>"
	output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
	output = output & "</form"
End sub


Function ProcessPrevMatch_1
	
	for i = 1 to 1500
		if player_array1(i,0) > 0 and player_array1(i,2) < prevdate then 	
			if player_array1(i,0) > player_array2(i,0) then
				player_array2(i,0) = player_array1(i,0)
				player_array2(i,1) = player_array1(i,1)
				player_array2(i,2) = player_array1(i,2)
			end if
			player_array1(i,0) = 0
			player_array1(i,1) = ""
			player_array1(i,2) = ""			
		end if 
	next
	prevdate = rows(r_date,row)

End Function

Function ProcessPrevMatch_2

	for i = 1 to 1500
		if player_array1(i,3) > 0 and player_array1(i,5) < prevdate then 	
			if player_array1(i,3) > player_array2(i,3) then
				player_array2(i,3) = player_array1(i,3)
				player_array2(i,4) = player_array1(i,4)
				player_array2(i,5) = player_array1(i,5)
			end if
			player_array1(i,3) = 0
			player_array1(i,4) = ""
			player_array1(i,5) = ""			
		end if 
	next
	prevdate = rows(r_date,row)

End Function

Function ProcessPrevMatch_3
	
	for i = 1 to 1500
		if player_array1(i,6) > 0 and player_array1(i,9) < prevdate then
			if player_array1(i,6) > player_array2(i,6) then
				player_array2(i,6) = player_array1(i,6)
				player_array2(i,7) = player_array1(i,7)
				player_array2(i,8) = player_array1(i,8)
				player_array2(i,9) = player_array1(i,9)
			end if
			player_array1(i,6) = 0
			player_array1(i,7) = 0
			player_array1(i,8) = ""
			player_array1(i,9) = ""							
		end if 
	next	
	prevdate = rows(r_date,row)
	
End Function

%>

</div>
</body>
</html>