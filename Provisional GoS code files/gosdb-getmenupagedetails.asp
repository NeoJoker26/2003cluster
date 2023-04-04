<%
Dim conn,sql,rs, n, work1, work2, side, comp, goalspergame1, goalspergame2
Dim tab1a(9,3), tab1b(9,3), tab2a(9,3), tab2b(9,3)

response.expires = -1
side = Request.QueryString("side")
comp = Request.QueryString("comp") 

if comp = "all" then 
table = "v_match_all"
else table = "v_match_FL"
end if

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

outline = side & "^"

if side = "left" or side = "both" then
 
	sql = "with CTE as "
	sql = sql & "( 	"
	sql = sql & "select player_id_spell1, surname, initials, 1 as starts, 0 as subs " 
	sql = sql & "from " & table & " a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 and startpos > 0 "
	sql = sql & "union all "
	sql = sql & "select player_id_spell1, surname, initials, 0 as starts, 1 as subs " 
	sql = sql & "from " & table & " a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 and startpos = 0 "
	sql = sql & ") "
	sql = sql & "select top 10 player_id_spell1, surname, initials, sum(starts) as totstarts, sum(subs) as totsubs, sum(starts+subs) as tot "
	sql = sql & "from CTE "
	sql = sql & "group by player_id_spell1, surname, initials "
	sql = sql & "order by tot desc "
	rs.open sql,conn,1,2

	n = 0 
	Do While Not rs.EOF
	  tab1a(n,0) = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname"))
	  if tab1a(n,0) = "Mathias Kouo-Doumbe" then tab1a(n,0) = "Mathias K-Doumbe"
	  tab1a(n,1) = rs.Fields("player_id_spell1")
	  tab1a(n,2) = rs.Fields("totstarts")
	  tab1a(n,3) = rs.Fields("totsubs")
	  n = n + 1
	  rs.MoveNext
	Loop
	rs.close

	sql = "with CTE as "
	sql = sql & "( "
	sql = sql & "select player_id_spell1, surname, initials, 1 as starts, 0 as subs " 
	sql = sql & "from " & table & " a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 and startpos > 0 "
	sql = sql & "  and player_id_spell1 in (select player_id_spell1 from player where last_game_year = 9999) "
	sql = sql & "union all "
	sql = sql & "select player_id_spell1, surname, initials, 0 as starts, 1 as subs " 
	sql = sql & "from " & table & " a join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 and startpos = 0 "
	sql = sql & "  and player_id_spell1 in (select player_id_spell1 from player where last_game_year = 9999) "
	sql = sql & ") "
	sql = sql & "select top 10 player_id_spell1, surname, initials, sum(starts) as totstarts, sum(subs) as totsubs, sum(starts+subs) as tot "
	sql = sql & "from CTE "
	sql = sql & "group by player_id_spell1, surname, initials "
	sql = sql & "order by tot desc "
	rs.open sql,conn,1,2
	
	n = 0 
	Do While Not rs.EOF
	  tab1b(n,0) = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname"))
	  if tab1b(n,0) = "Mathias Kouo-Doumbe" then tab1b(n,0) = "Mathias K-Doumbe"
	  tab1b(n,1) = rs.Fields("player_id_spell1")
	  tab1b(n,2) = rs.Fields("totstarts")
	  tab1b(n,3) = rs.Fields("totsubs")
	  n = n + 1
	  rs.MoveNext
	Loop
	rs.close
	
	outline = outline & "<table id=""gottable1"">"
    outline = outline & "<tr><td colspan=""6"" style=""border-right-style: none; border-right-width: medium""><p>Top 10 Appearances<img src=""images/dummy.gif"" width=""30""><span class=""tah"">[St=Starts, Sb=Subs]</span><br>"
    if comp = "FL" then
    	outline = outline & "in Leagues"
    	outline = outline & " | "
    	outline = outline & "<a href=""javascript:GetTable('left','all')"" style=""text-decoration:underline;"">All Competitions</a>"
      else
    	outline = outline & "in All Competitions"
    	outline = outline & " | "
    	outline = outline & "<a href=""javascript:GetTable('left','FL')"" style=""text-decoration:underline;"">Leagues</a>"
    end if
    outline = outline & "</p></td></tr>"
    outline = outline & "<tr><td><b>All Time</b></td><td class=""right"">St</td><td class=""right"">Sb</td><td style=""padding-left: 12""><b>Current Squad</b></td><td class=""right"">St</td><td class=""right"">Sb</td></tr>"
	for n = 0 to 9
		work1 = split(tab1a(n,0)," ")
		work2 = split(tab1b(n,0)," ")
		outline = outline & "<tr><td><a href=""gosdb-players2.asp?pid=" & tab1a(n,1) & "&scp=1,2,3,4,5,6,7"">" & left(work1(0),1) & ". " & work1(1) & "</a></td><td class=""right"">" & tab1a(n,2) & "</td><td class=""right"">" & tab1a(n,3) & "</td><td style=""padding-left: 12""><a href=""gosdb-players2.asp?pid=" & tab1b(n,1) & "&scp=1,2,3,4,5,6,7"">" & left(work2(0),1) & ". " & work2(1) & "</a></td><td class=""right"">" & tab1b(n,2) & "</td><td class=""right"">" & tab1b(n,3) & "</td></tr>"
	next
	outline = outline & "</table>"
	
end if

if side = "right" or side = "both" then

    outline = outline & "^"

	sql = "select top 10 player_id_spell1, surname, initials, count(distinct b.date) as appears, count(c.player_id) as goals "
	sql = sql & "from " & table & " a join match_player b on a.date = b.date " 
	sql = sql & "left outer join match_goal c on b.player_id = c.player_id and b.date = c.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "group by player_id_spell1, surname, initials "
	sql = sql & "order by goals desc "
	rs.open sql,conn,1,2
	
	n = 0 
	Do While Not rs.EOF
	  tab2a(n,0) = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname"))
	  if tab2a(n,0) = "Mathias Kouo-Doumbe" then tab2a(n,0) = "Mathias K-Doumbe"
	  tab2a(n,1) = rs.Fields("player_id_spell1")
	  tab2a(n,2) = rs.Fields("goals")
	  tab2a(n,3) = rs.Fields("appears")
	  n = n + 1
  	  rs.MoveNext
	Loop
	rs.close

	sql = "select top 10 player_id_spell1, surname, initials, count(distinct b.date) as appears, count(c.player_id) as goals "
	sql = sql & "from " & table & " a join match_player b on a.date = b.date " 
	sql = sql & "left outer join match_goal c on b.player_id = c.player_id and b.date = c.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where d.player_id < 9000 "
	sql = sql & "  and player_id_spell1 in (select player_id_spell1 from player where last_game_year = 9999) "
	sql = sql & "group by player_id_spell1, surname, initials "
	sql = sql & "order by goals desc "
	rs.open sql,conn,1,2

	n = 0 
	Do While Not rs.EOF
	  tab2b(n,0) = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname"))
	  if tab2b(n,0) = "Mathias Kouo-Doumbe" then tab2b(n,0) = "Mathias K-Doumbe"
	  tab2b(n,1) = rs.Fields("player_id_spell1")
	  tab2b(n,2) = rs.Fields("goals")
	  tab2b(n,3) = rs.Fields("appears")
	  n = n + 1
	  rs.MoveNext
	Loop
	rs.close
	
	outline = outline & "<table id=""gottable2"">"
    outline = outline & "<tr><td colspan=""6"" style=""border-right-style: none; border-right-width: medium""><p>Top 10 Goal Scorers<img src=""images/dummy.gif"" width=""30""><span class=""tah"">[G/G = goals per game]</span><br>"
    if comp = "FL" then
    	outline = outline & "in Leagues"
    	outline = outline & " | "
    	outline = outline & "<a href=""javascript:GetTable('right','all')"" style=""text-decoration:underline;"">All Competitions</a>"
      else
    	outline = outline & "in All Competitions"
    	outline = outline & " | "
    	outline = outline & "<a href=""javascript:GetTable('right','FL')"" style=""text-decoration:underline;"">Leagues</a>"
    end if
    outline = outline & "</p></td></tr>"
    outline = outline & "<tr><td><b>All Time</b></td><td class=""right"">Gls</td><td class=""right"">G/G</td><td style=""padding-left: 12""><b>Current Squad</b><td class=""right"">Gls</td><td class=""right"">G/G</td></tr>"

	for n = 0 to 9
		work1 = split(tab2a(n,0)," ")
		work2 = split(tab2b(n,0)," ")
		goalspergame1 = round(tab2a(n,2)/tab2a(n,3),2) 
	  	if Instr(goalspergame1,".") = 0 then goalspergame1 = goalspergame1 & "."  'no decimal point, must be a whole number, so add a dec. point
	  	goalspergame1 = left(goalspergame1 & "00",4)
		goalspergame2 = round(tab2b(n,2)/tab2b(n,3),2) 
	  	if Instr(goalspergame2,".") = 0 then goalspergame2 = goalspergame2 & "."  'no decimal point, must be a whole number, so add a dec. point
	  	goalspergame2 = left(goalspergame2 & "00",4)
		outline = outline & "<tr><td><a href=""gosdb-players2.asp?pid=" & tab2a(n,1) & "&scp=1,2,3,4,5,6,7"">" & left(work1(0),1) & ". " & work1(1) & "</a></td><td class=""right"">" & tab2a(n,2) & "</td><td class=""right"">" & goalspergame1 & "</td><td style=""padding-left: 12""><a href=""gosdb-players2.asp?pid=" & tab2b(n,1) & "&scp=1,2,3,4,5,6,7"">" & left(work2(0),1) & ". " & work2(1) & "</a></td><td class=""right"">" & tab2b(n,2) & "</td><td class=""right"">" & goalspergame2 & "</td></tr>"
	next
	outline = outline & "</table>"
end if

conn.close
	
response.write(outline)
%>