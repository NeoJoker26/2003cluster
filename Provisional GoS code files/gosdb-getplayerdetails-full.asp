<%
Dim conn,sql,rs,rsrec,rsrec2,n,work1,work2,scope, playerid,playerlist, starts(2), subs(2), rank(2), goals(2), opposition, displaydate, stillatclub, tagno, manager_years, player_years
Dim calc_age, calc_age_add

response.expires = -1
playerid = Request.QueryString("playerid")
if len(playerid) > 4 then player_id = 1
scope = Request.QueryString("scp")
if instr(scope," or ") > 0 or instr(scope,"union ") > 0 or instr(scope,"drop ") > 0 or instr(scope,"=") > 0 then scope = ""
tagno = 1

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rsrec = Server.CreateObject("ADODB.Recordset")
Set rsrec2 = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

'First do the right hand side
 	
	sql = "with cte1 as ( "
	sql = sql & "select managers, manager_id1, isnull(manager_id2,999) as manager_id2, from_date, isnull(to_date,GETDATE()) as to_date, "
	sql = sql & "surname as player_surname, min(date) as player_min_date, max(date) as player_max_date, "
	sql = sql & "sum(case when startpos > 0 then 1 else 0 end) as player_starts, "
	sql = sql & "sum(case when startpos = 0 then 1 else 0 end) as player_subs "
	sql = sql & "from match_player a join player b on a.player_id = b.player_id "
	sql = sql & "join v_managerspell_horiz c on a.date >= c.from_date and a.date <= isnull(c.to_date,GETDATE()) "
	sql = sql & "where player_id_spell1 = '" & playerid & "' "
	sql = sql & "group by managers, manager_id1, manager_id2, from_date, to_date, surname "
	sql = sql & "), "
	sql = sql & "cte2 as ( "
	sql = sql & "select from_date as manager_from_date, isnull(to_date,GETDATE()) as manager_to_date, "
	sql = sql & "count(*) as manager_games "
	sql = sql & "from match a "
	sql = sql & "join v_managerspell_horiz b on a.date >= b.from_date and a.date <= isnull(b.to_date,GETDATE()) "
	sql = sql & "group by from_date, to_date "
	sql = sql & ") "
	sql = sql & "select managers, manager_id1, manager_id2, "
	sql = sql & "year(manager_from_date) as manager_from_year, year(manager_to_date) as manager_to_year, manager_games, "
	sql = sql & "player_surname, "
	sql = sql & "year(player_min_date) as player_from_year, year(player_max_date) as player_to_year, player_starts, player_subs "
	sql = sql & "from cte1 a join cte2 b on a.from_date = b.manager_from_date "
	sql = sql & "	and a.to_date = b.manager_to_date "
	sql = sql & "order by manager_from_date "
	
	rs.open sql,conn,1,2
			
	if not rs.EOF then outline = "<p style=""margin: 9px 0 0 0;""><b>MANAGERS</b><span class=""style2""> with the number of games in charge and the times " & rtrim(rs.Fields("player_surname")) & " was selected (starts-subs):</span></p>"
		
	Do While Not rs.EOF
		if rs.Fields("manager_from_year") = rs.Fields("manager_to_year") then
			manager_years = rs.Fields("manager_from_year")
		  elseif left(rs.Fields("manager_from_year"),2) = left(rs.Fields("manager_to_year"),2) then
		  	manager_years = rs.Fields("manager_from_year") & "-" & right(rs.Fields("manager_to_year"),2)
		  else manager_years = rs.Fields("manager_from_year") & "-" & rs.Fields("manager_to_year")
		end if  	
		if rs.Fields("player_from_year") = rs.Fields("player_to_year") then
			player_years = rs.Fields("player_from_year")
		  elseif left(rs.Fields("player_from_year"),2) = left(rs.Fields("player_to_year"),2) then
		  	player_years = rs.Fields("player_from_year") & "-" & right(rs.Fields("player_to_year"),2)
		  else player_years = rs.Fields("player_from_year") & "-" & rs.Fields("player_to_year")
		end	if
	  	outline = outline & "<p class=""style1boldgreen"" style=""margin: 4px 0 0;"">"
	  	outline = outline & "<a class=""manager_name"" id=""man-" & rs.Fields("manager_id1") & "-" & rs.Fields("manager_id2") & """ href=""#""><u>" & rs.Fields("managers") & "</u></a></p>"
	  	outline = outline & "<p class=""style1"" style=""margin: 1px 0;"">Man PAFC: " & rs.Fields("manager_games") & " (" & manager_years & ")"   
	  	outline = outline & "<br>Sel " & rtrim(rs.Fields("player_surname")) & ": " & rs.Fields("player_starts") & "-" & rs.Fields("player_subs") & " (" & player_years & ")"	  	 
		rs.Movenext
	loop
	rs.close

	sql = "select e.player_id, e.player_id_spell1, e.spell, e.surname, e.forename, e.initials, count(*) as count "
    sql = sql & "from v_match_all a "
    sql = sql & " join match_player b on a.date = b.date "
    sql = sql & " join match_player c on a.date = c.date "
	sql = sql & " join player d on b.player_id = d.player_id "
	sql = sql & " join player e on c.player_id = e.player_id "
	sql = sql & "where d.player_id in ( "
	sql = sql & "	select player_id "
	sql = sql & "	from player f "	
	sql = sql & "	where f.player_id_spell1 = '" & playerid & "' "
	sql = sql & "	) "
	sql = sql & "and e.player_id <> d.player_id "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "group by e.player_id, e.player_id_spell1, e.spell, e.surname, e.forename, e.initials "
	sql = sql & "order by e.surname, e.forename, e.initials, e.player_id_spell1, e.spell "

	rs.open sql,conn,1,2
	
	outline = outline & "<p style=""margin: 15px 0 4px 0;""><b>TEAM-MATES</b><span class=""style2""> and occasions for all matches in the chosen competitions:</span></p>"
			
	Do While Not rs.EOF
	  	outline = outline & "<p style=""white-space:nowrap; margin: 0 0 3px;""><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & "&scp=" & scope & """><u>" 
	  	if IsNull(rs.Fields("forename")) then
	  	  	outline = outline & rtrim(rs.Fields("surname")) & ", " & rtrim(rs.Fields("initials")) & "</u></a> (" & rs.Fields("count") & ")</p>"
	  	  else
	  	  	outline = outline & rtrim(rs.Fields("surname")) & ", " & rtrim(rs.Fields("forename")) & "</u></a> (" & rs.Fields("count") & ")</p>"
		end if
		rs.Movenext
	loop
	rs.close
		
	outline = outline & "^" 	' *** note the ^ which signals a new column for the calling program (gosdb-players2) 
    
'Now do the middle section
	
	playerlist = "'" & playerid & "'"
	
	sql = "select player_id "
	sql = sql & "from player "
	sql = sql & "where player_id_spell1 = '" & playerid & "'"
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		playerlist = playerlist & ",'" &rs.Fields("player_id") & "'" 
		rs.Movenext
	Loop
	rs.close

	outline = outline & "<table border=""0"" cellpadding=""0"" style=""margin: 4 0 12 -4; border-collapse: collapse"" bordercolor=""#c0c0c0"" >"
  	outline = outline & "<tr>"
    outline = outline & "<td><b>PAFC Summary</b></td>"
    outline = outline & "<td align=""left"" width=""70""><b>Leagues</b></td>"
    outline = outline & "<td align=""left"" width=""70""><b>Cups</b></td>"
    outline = outline & "<td align=""left"" width=""70""><b>All Comps</b></td>"
  	outline = outline & "</tr>"
	
  	sql = "with CTE as "
	sql = sql & "( 	"
  	sql = sql & "select 'A' as queryid, rank() over (order by b.date) as rank, a.player_id "
	sql = sql & "from player a "
	sql = sql & "join match_player b on a.player_id = b.player_id "
	sql = sql & "where b.date = ( "
 	sql = sql & "	select min(c1.date) "
 	sql = sql & "	from player a1 "
	sql = sql & "	join match_player b1 on a1.player_id = b1.player_id "
	sql = sql & "	join v_match_all c1 on b1.date = c1.date "
	sql = sql & "	where a1.player_id = a.player_id "
	sql = sql & "	and compcat in (" & scope & ") "
	sql = sql & "   and LFC <> 'C' "
	sql = sql & "	) "
	sql = sql & "	and a.player_id < 8000 "
	sql = sql & "   and spell = 1 "
	sql = sql & "union all "
  	sql = sql & "select 'B' as queryid, rank() over (order by b.date) as rank, a.player_id "
	sql = sql & "from player a "
	sql = sql & "join match_player b on a.player_id = b.player_id "
	sql = sql & "where b.date = ( "
 	sql = sql & "	select min(c1.date) "
 	sql = sql & "	from player a1 "
	sql = sql & "	join match_player b1 on a1.player_id = b1.player_id "
	sql = sql & "	join v_match_all c1 on b1.date = c1.date "
	sql = sql & "	where a1.player_id = a.player_id "
	sql = sql & "	and compcat in (" & scope & ") "
	sql = sql & "   and LFC = 'C' "
	sql = sql & "	) "
	sql = sql & "	and a.player_id < 8000 "
	sql = sql & "   and spell = 1 "
	sql = sql & "union all "
  	sql = sql & "select 'C' as queryid, rank() over (order by b.date) as rank, a.player_id "
	sql = sql & "from player a "
	sql = sql & "join match_player b on a.player_id = b.player_id "
	sql = sql & "where b.date = ( "
 	sql = sql & "	select min(c1.date) "
 	sql = sql & "	from player a1 "
	sql = sql & "	join match_player b1 on a1.player_id = b1.player_id "
	sql = sql & "	join v_match_all c1 on b1.date = c1.date "
	sql = sql & "	where a1.player_id = a.player_id "
	sql = sql & "	and compcat in (" & scope & ") "
	sql = sql & "	) "
	sql = sql & "	and a.player_id < 8000 "
	sql = sql & "   and spell = 1 "
	sql = sql & ") "
	sql = sql & "select queryid, rank "
	sql = sql & "from CTE "
	sql = sql & "where player_id in (" & playerlist & ") "
	sql = sql & "order by queryid "

	rs.open sql,conn,1,2
  	
  	outline = outline & "<tr>"
    outline = outline & "<td>Ascending player no.</b><br><span style=""font-size:10px"">[on debut]</span></td>"
    
    for n = 0 to 2
    	rank(n) = "-"
    next
    
    Do While Not rs.EOF
    	if rs.Fields("queryid") = "A" then 
    		rank(0) = rs.Fields("rank")
    	end if  
    	if rs.Fields("queryid") = "B" then 
    		rank(1) = rs.Fields("rank")
    	end if
    	if rs.Fields("queryid") = "C" then 
    		rank(2) = rs.Fields("rank")
      	end if  
    	rs.Movenext
	Loop
    
    
    outline = outline & "<td align=""left"">" & rank(0) & "</td>"
    outline = outline & "<td align=""left"">" & rank(1) & "</td>"
    outline = outline & "<td align=""left"">" & rank(2) & "</td>"
     
	rs.close 
  	
  	outline = outline & "<tr>"
    outline = outline & "<td><b>Appearances</b><br><span style=""font-size:10px"">[starts-subs]</span></td>"	
	
	sql = "with detailCTE as "
	sql = sql & "( 	"
	sql = sql & "select 'A' as queryid, player_id_spell1, 1 as starts1, 0 as subs1, 0 as starts2, 0 as subs2, 0 as starts3, 0 as subs3 " 
	sql = sql & "from v_match_all a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 and startpos > 0 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "and LFC <> 'C' "
	sql = sql & "union all "
	sql = sql & "select 'A', player_id_spell1, 0, 1, 0, 0, 0, 0 " 
	sql = sql & "from v_match_all a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 and startpos = 0 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "and LFC <> 'C' "
	sql = sql & "union all "
	sql = sql & "select 'B', player_id_spell1, 0, 0, 1, 0, 0, 0 " 
	sql = sql & "from v_match_all a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 and startpos > 0 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "and LFC = 'C' "
	sql = sql & "union all "
	sql = sql & "select 'B', player_id_spell1, 0, 0, 0, 1, 0, 0 " 
	sql = sql & "from v_match_all a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 and startpos = 0 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "and LFC = 'C' "
 	sql = sql & "union all "
	sql = sql & "select 'C', player_id_spell1, 0, 0, 0, 0, 1, 0 " 
	sql = sql & "from v_match_all a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 and startpos > 0 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "union all "
	sql = sql & "select 'C', player_id_spell1, 0, 0, 0, 0, 0, 1 " 
	sql = sql & "from v_match_all a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 and startpos = 0 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "), "
	sql = sql & "  sumCTE as "
	sql = sql & "( 	"
	sql = sql & "select queryid, player_id_spell1, sum(starts1) as totstarts1, sum(subs1) as totsubs1, sum(starts1+subs1) as tot1, sum(starts2) as totstarts2, sum(subs2) as totsubs2, sum(starts2+subs2) as tot2, sum(starts3) as totstarts3, sum(subs3) as totsubs3, sum(starts3+subs3) as tot3  "
	sql = sql & "from detailCTE "
	sql = sql & "group by queryid, player_id_spell1  "
	sql = sql & "), "
	sql = sql & "  rankCTE as "
	sql = sql & "( 	"
	sql = sql & "select 'A' as queryid, rank() over (order by tot1 desc) as rank, player_id_spell1, totstarts1 as totstarts, totsubs1 as totsubs, tot1 as tot "
	sql = sql & "from sumCTE "
	sql = sql & "where queryid = 'A' "
	sql = sql & "union all "
	sql = sql & "select 'B', rank() over (order by tot2 desc) as rank, player_id_spell1, totstarts2, totsubs2, tot2 "
	sql = sql & "from sumCTE "
	sql = sql & "where queryid = 'B' "
	sql = sql & "union all "
	sql = sql & "select 'C', rank() over (order by tot3 desc) as rank, player_id_spell1, totstarts3, totsubs3, tot3 "
	sql = sql & "from sumCTE "
	sql = sql & "where queryid = 'C' "
	sql = sql & ") "
	sql = sql & "select queryid, rank, totstarts, totsubs, tot "
	sql = sql & "from rankCTE "
	sql = sql & "where player_id_spell1 = '" & playerid & "' "
	sql = sql & "order by queryid "

	rs.open sql,conn,1,2

    for n = 0 to 2
    	starts(n) = 0
    	subs(n) = 0
    	rank(n) = "-"
    next
    
    Do While Not rs.EOF
    	if rs.Fields("queryid") = "A" then 
    		starts(0) = rs.Fields("totstarts")
    		subs(0) = rs.Fields("totsubs") 
    		rank(0) = rs.Fields("rank")
    	end if  
    	if rs.Fields("queryid") = "B" then 
    		starts(1) = rs.Fields("totstarts")
    		subs(1) = rs.Fields("totsubs") 
    		rank(1) = rs.Fields("rank")
    	end if
    	if rs.Fields("queryid") = "C" then 
    		starts(2) = rs.Fields("totstarts")
    		subs(2) = rs.Fields("totsubs") 
    		rank(2) = rs.Fields("rank")
      	end if  
    	rs.Movenext
	Loop
  
    outline = outline & "<td align=""left"">" & starts(0)+subs(0) & "<br><span style=""font-size:10px"">[" & starts(0) & "-" & subs(0) & "]</span></td>"
    outline = outline & "<td align=""left"">" & starts(1)+subs(1) & "<br><span style=""font-size:10px"">[" & starts(1) & "-" & subs(1) & "]</span></td>"
    outline = outline & "<td align=""left""><b>" & starts(2)+subs(2) & "</b><br><span style=""font-size:10px"">[" & starts(2) & "-" & subs(2) & "]</span></td>"

  	outline = outline & "</tr>"
  	
	rs.close

	sql = "select 'A' as queryid, count(distinct d.player_id) as playercount "
	sql = sql & "from v_match_all a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "and spell = 1 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "and LFC <> 'C' "
	sql = sql & "union all "
	sql = sql & "select 'B', count(distinct d.player_id) "
	sql = sql & "from v_match_all a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "and spell = 1 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "and LFC = 'C' "
	sql = sql & "union all "
	sql = sql & "select 'C', count(distinct d.player_id) "
	sql = sql & "from v_match_all a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "and spell = 1 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "order by queryid "
	
	rs.open sql,conn,1,2
 
  	outline = outline & "<tr>"
    outline = outline & "<td>Appearance ranking<br><span style=""font-size:10px"">[1 = most]</span></td>"
    outline = outline & "<td align=""left"">" & rank(0) & "<br><span style=""font-size:10px"">"
    if rs.Fields("playercount") > 0 then outline = outline & "of " & rs.Fields("playercount") 
    outline = outline & "</span></td>"
    rs.Movenext
    outline = outline & "<td align=""left"">" & rank(1) & "<br><span style=""font-size:10px"">"
    if rs.Fields("playercount") > 0 then outline = outline & "of " & rs.Fields("playercount") 
    outline = outline & "</span></td>"
    rs.Movenext
    outline = outline & "<td align=""left"">" & rank(2) & "<br><span style=""font-size:10px"">"
    if rs.Fields("playercount") > 0 then outline = outline & "of " & rs.Fields("playercount") 
    outline = outline & "</span></td>"
    outline = outline & "</tr>"
     
	rs.close 
	       
  	outline = outline & "<tr>"
    outline = outline & "<td><b>Goals Scored</b></td>"

	sql = "with detailCTE as "
	sql = sql & "( 	"
	sql = sql & "select 'A' as queryid, player_id_spell1, surname, forename, initials, 1 as goals1, 0 as goals2, 0 as goals3 " 
	sql = sql & "from v_match_all a join match_goal b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "and LFC <> 'C' "
	sql = sql & "union all "
	sql = sql & "select 'B', player_id_spell1, surname, forename, initials, 0, 1, 0 " 
	sql = sql & "from v_match_all a join match_goal b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "and LFC = 'C' "
	sql = sql & "union all "
	sql = sql & "select 'C', player_id_spell1, surname, forename, initials, 0, 0, 1 " 
	sql = sql & "from v_match_all a join match_goal b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "), "
	sql = sql & "  sumCTE as "
	sql = sql & "( 	"
	sql = sql & "select queryid, player_id_spell1, surname, forename, initials, sum(goals1) as totgoals1, sum(goals2) as totgoals2, sum(goals3) as totgoals3  "
	sql = sql & "from detailCTE "
	sql = sql & "group by queryid, player_id_spell1, surname, forename, initials "
	sql = sql & "), "
	sql = sql & "  rankCTE as "
	sql = sql & "( 	"
	sql = sql & "select 'A' as queryid, rank() over (order by totgoals1 desc) as rank, player_id_spell1, surname, forename, initials, totgoals1 as totgoals "
	sql = sql & "from sumCTE "
	sql = sql & "where queryid = 'A' "
	sql = sql & "union all "
	sql = sql & "select 'B', rank() over (order by totgoals2 desc) as rank, player_id_spell1, surname, forename, initials, totgoals2 "
	sql = sql & "from sumCTE "
	sql = sql & "where queryid = 'B' "
	sql = sql & "union all "
	sql = sql & "select 'C', rank() over (order by totgoals3 desc) as rank, player_id_spell1, surname, forename, initials, totgoals3 "
	sql = sql & "from sumCTE "
	sql = sql & "where queryid = 'C' "
	sql = sql & ") "
	sql = sql & "select queryid, rank, totgoals "
	sql = sql & "from rankCTE "
	sql = sql & "where player_id_spell1 = '" & playerid & "' "
	sql = sql & "order by queryid "

	rs.open sql,conn,1,2

    for n = 0 to 2
    	goals(n) = 0
    	rank(n) = "-"
    next
    
    Do While Not rs.EOF
    	if rs.Fields("queryid") = "A" then 
    		goals(0) = rs.Fields("totgoals")
    		rank(0) = rs.Fields("rank")
    	end if  
    	if rs.Fields("queryid") = "B" then 
    		goals(1) = rs.Fields("totgoals")
    		rank(1) = rs.Fields("rank")
    	end if
    	if rs.Fields("queryid") = "C" then 
    		goals(2) = rs.Fields("totgoals")
    		rank(2) = rs.Fields("rank")
      	end if  
    	rs.Movenext
	Loop
  
    outline = outline & "<td align=""left"">" & goals(0) & "</td>"
    outline = outline & "<td align=""left"">" & goals(1) & "</td>"
    outline = outline & "<td align=""left""><b>" & goals(2) & "</b></td>"

  	outline = outline & "</tr>"
  	
	rs.close

	sql = "select 'A' as queryid, count(distinct d.player_id_spell1) as playercount "
	sql = sql & "from v_match_all a join match_goal b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "and LFC <> 'C' "
	sql = sql & "union all "
	sql = sql & "select 'B', count(distinct d.player_id_spell1) "
	sql = sql & "from v_match_all a join match_goal b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "and LFC = 'C' "
	sql = sql & "union all "
	sql = sql & "select 'C', count(distinct d.player_id_spell1) "
	sql = sql & "from v_match_all a join match_goal b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "and compcat in (" & scope & ") "
	sql = sql & "order by queryid "
	
	rs.open sql,conn,1,2

  	outline = outline & "<tr>"
    outline = outline & "<td>Goals Scored ranking<br><span style=""font-size:10px"">[1 = most]</span></td>"
    outline = outline & "<td align=""left"">" & rank(0) & "<br><span style=""font-size:10px"">"
    if rs.Fields("playercount") > 0 then outline = outline & "of " & rs.Fields("playercount") 
    outline = outline & "</span></td>"
    rs.Movenext
    outline = outline & "<td align=""left"">" & rank(1) & "<br><span style=""font-size:10px"">"
    if rs.Fields("playercount") > 0 then outline = outline & "of " & rs.Fields("playercount")
    outline = outline & "</span></td>"
    rs.Movenext
    outline = outline & "<td align=""left"">" & rank(2) & "<br><span style=""font-size:10px"">"
    if rs.Fields("playercount") > 0 then outline = outline & "of " & rs.Fields("playercount") 
    outline = outline & "</span></td>"    
    outline = outline & "</tr>"
     
	rs.close   	 
    
  	outline = outline & "</tr>"
	outline = outline & "</table>"
	
	sql = "select spell, a.player_id, dob, surname, forename, initials, came_from, went_to, last_game_year, d.date as first_date, e.date as last_date, "
	sql = sql & "d.homeaway as firsthomeaway, d.opposition as firstopposition, d.opposition_qual as firstoppositionqual, d.goalsfor as firstgoalsfor, d.goalsagainst as firstgoalsagainst, "
	sql = sql & "e.homeaway as lasthomeaway, e.opposition as lastopposition, e.opposition_qual as lastoppositionqual, e.goalsfor as lastgoalsfor, e.goalsagainst as lastgoalsagainst "
	sql = sql & "from player a "
	sql = sql & "join match_player b on a.player_id = b.player_id "
	sql = sql & "join match_player c on a.player_id = c.player_id "
	sql = sql & "join v_match_all d on b.date = d.date "
	sql = sql & "join v_match_all e on c.date = e.date "
	sql = sql & "where a.player_id in (" & playerlist & ") "
	sql = sql & "and d.date = ( "
 	sql = sql & "select min(date) "
 	sql = sql & "from player a1 "
	sql = sql & "	join match_player b1 on a1.player_id = b1.player_id "
	sql = sql & "	where a1.player_id = a.player_id "
	sql = sql & "	) "
	sql = sql & "and e.date = ( "
 	sql = sql & "select max(date) "
 	sql = sql & "from player a2 "
	sql = sql & "	join match_player c2 on a2.player_id = c2.player_id "
	sql = sql & "	where a2.player_id = a.player_id "
	sql = sql & "	) "	
	sql = sql & "order by spell "

	rs.open sql,conn,1,2

	n = 0
	
	Do While Not rs.EOF
		n = n + 1 
		outline = outline & "<p style=""margin-top:9;""><b>"
		if rs.RecordCount > 1 then outline = outline & "<u>Spell " & n & "</u> - " 
		outline = outline & "Came from: </b>" & rs.Fields("came_from") & "</p>"
		
		outline = outline & "<p><b>First Match</b> (any comp): " & Day(rs.Fields("first_date")) & " " & MonthName(Month(rs.Fields("first_date")),True) & " " & Year(rs.Fields("first_date"))
		if not IsNull(rs.Fields("dob")) then 
			calc_age = datediff("yyyy",rs.Fields("dob"),rs.Fields("first_date"))
			calc_age_add = dateadd("yyyy", calc_age, rs.Fields("dob"))
			if datediff("y", rs.Fields("first_date"), calc_age_add) > 0 then calc_age = calc_age - 1   
			outline = outline & " (age " & calc_age & ")"		
		end if
		
		if rs.Fields("firsthomeaway") = "H" then
			outline = outline & "<br>Argyle " & rs.Fields("firstgoalsfor") & " " & rs.Fields("firstopposition") & " " & rs.Fields("firstoppositionqual") & " " & rs.Fields("firstgoalsagainst") & "</p>"
		  else
		  	outline = outline & "<br>" & rs.Fields("firstopposition") & " " & rs.Fields("firstoppositionqual") & " " & rs.Fields("firstgoalsagainst") & " " & "Argyle " & rs.Fields("firstgoalsfor") & "</p>"
		 end if
		
		stillatclub = ""
		if rs.Fields("last_game_year") = 9999 then stillatclub = "y"
		
		if stillatclub = "" then 
			outline = outline & "<p><b>Last Match</b> (any comp.): "
		  else
		  	outline = outline & "<p><b>Latest Match</b> (any comp.): "
		end if
		
		outline = outline & Day(rs.Fields("last_date")) & " " & MonthName(Month(rs.Fields("last_date")),True) & " " & Year(rs.Fields("last_date"))
		
		if not IsNull(rs.Fields("dob")) then 
			calc_age = datediff("yyyy",rs.Fields("dob"),rs.Fields("last_date"))
			calc_age_add = dateadd("yyyy", calc_age, rs.Fields("dob"))
			if datediff("y", rs.Fields("last_date"), calc_age_add) > 0 then calc_age = calc_age - 1   
			outline = outline & " (age " & calc_age & ")"		
		end if
		
		if rs.Fields("lasthomeaway") = "H" then
			outline = outline & "<br>Argyle " & rs.Fields("lastgoalsfor") & " " & rs.Fields("lastopposition") & " " & rs.Fields("lastoppositionqual") & " " & rs.Fields("lastgoalsagainst") & "</p>"
		  else
		  	outline = outline & "<br>" & rs.Fields("lastopposition") & " " & rs.Fields("lastoppositionqual") & " " & rs.Fields("lastgoalsagainst") & " " & "Argyle " & rs.Fields("lastgoalsfor") & "</p>"
		 end if
		if stillatclub = "" then outline = outline & "<p><b>Went to: </b>" & rs.Fields("went_to") & "</p>"
		outline = outline & "<p>" 
		outline = outline & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse; margin: 0 0 0 -5;"" width=""380"">"
		
		outline = outline & "<tr><td class=""season"" width=""140"" valign=""top""><b>Playing Record"
		if stillatclub = "y" then outline = outline & "<br>so far"
		outline = outline & ":</b></td><td width=""40"" align=""right"">Lge Apps</td><td width=""45"" align=""right"">Lge Subs</td><td width=""40"" align=""right"">Lge Goals</td>"
		outline = outline & "<td width=""40"" align=""right"">Cup Apps</td><td width=""40"" align=""right"">Cup Subs</td><td width=""40"" align=""right"">Cup Goals</td></tr>"		
	    
		    sql = "select years, sum(LSt) as l_starts, sum (LSu) as l_subs, sum(LG) as l_goals, sum(CSt) as c_starts, sum(CSu) as c_subs, sum(CG) as c_goals "
    		sql = sql & " from ( "
    		sql = sql & "select years, "
    		sql = sql & "case when (startpos > 0 and LFC <> 'C') then 1 else 0 end as LSt, "
    		sql = sql & "case when (startpos = 0 and LFC <> 'C') then 1 else 0 end as LSu, "
    		sql = sql & "case when (startpos > 0 and LFC = 'C') then 1 else 0 end as CSt, "
     		sql = sql & "case when (startpos = 0 and LFC = 'C') then 1 else 0 end as CSu, "
    		sql = sql & "0 as LG, 0 as CG "
    		sql = sql & "from season a join v_match_all b on date between date_start and date_end "
    		sql = sql & " join match_player c on b.date = c.date "
			sql = sql & " join player d on c.player_id = d.player_id "
			sql = sql & "where d.player_id = '" & rs.Fields("player_id") & "' "
			sql = sql & "and compcat in (" & scope & ") "
    		sql = sql & "union all "
			sql = sql & "select years, "
    		sql = sql & "0, 0, 0, 0, "
    		sql = sql & "case when LFC <> 'C' then 1 else 0 end, "
    		sql = sql & "case when LFC = 'C' then 1 else 0 end "
    		sql = sql & "from season a join v_match_all b on date between date_start and date_end "
    		sql = sql & " join match_goal c on b.date = c.date "
			sql = sql & " join player d on c.player_id = d.player_id "
			sql = sql & "where d.player_id = '" & rs.Fields("player_id") & "' "
			sql = sql & "and compcat in (" & scope & ") "
			sql = sql & ") as subsel "
			sql = sql & "group by years "
			sql = sql & "order by years "

			rsrec.open sql,conn,1,2
	
			if rsrec.recordcount = 0 then
				outline = outline & "<tr><td colspan=""7"">No appearances in the selected competitions</td></tr>"
			  else
				Do While Not rsrec.EOF
					outline = outline & "<tr id=""trtag" & tagno & """>"
					outline = outline & "<td nowrap=""nowrap""><a style=""font-family:courier;"" id=""xtag" & tagno & """ href=""javascript:Toggle2('tag" & tagno & "','" & scope & "');"">[+]</a><span id=""d1tag" & tagno & """ style=""display:none;"">" & rs.Fields("player_id") & "</span><span id=""d2tag" & tagno & """ style=""display:none;"">" & rsrec.Fields("years") & "</span>"
					outline = outline & " <span class=""season"" style=""position: relative; border-bottom: 1px solid #d0d0d0;""><a href=""gosdb-season.asp?years=" & rsrec.Fields("years") & """>" & rsrec.Fields("years") & "</a></span></td><td align=""right"">" & rsrec.Fields("l_starts") & "</td><td align=""right"">" & rsrec.Fields("l_subs") & "</td><td align=""right"">" & rsrec.Fields("l_goals") & "</td>"
					outline = outline & "<td align=""right"">" & rsrec.Fields("c_starts") & "</td><td align=""right"">" & rsrec.Fields("c_subs") & "</td><td align=""right"">" & rsrec.Fields("c_goals") & "</td></tr>"
					outline = outline & "<tr><td colspan=""7"" style=""padding: 0 0 0 0; margin: 0 0 0 0""><span id=""tag" & tagno & """ style=""display:none;""><img border=""0"" style=""margin-left: 6;"" src=""images/ajax-loader.gif""></span></td></tr>"			  
					tagno = tagno + 1
					rsrec.Movenext
				loop
			end if
			
			rsrec.close
		
		outline = outline & "</table></p>"
		rs.Movenext
	Loop
	
	rs.close
	
	if scope <> "1,2,3,4,5,6,7" then outline = outline & "<p style=""color: #CC3300; margin: 9 0 12 0;""><b>Remember: the detail here relates to the selected competitions. Select all competitions for a complete view. "

conn.close

response.write(outline)
%>