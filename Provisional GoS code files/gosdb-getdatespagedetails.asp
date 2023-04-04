<%
Dim conn,sql,rs, n, work1, work2, side, comp
Dim tab1a(9,2), tab1b(9,2), tab2a(9,2), tab2b(9,2)

response.expires = -1
side = Request.QueryString("side")

table = "v_match_all"

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

outline = side & "^"

	sql = "select top 10 date, sum(matches) as matches, " 
	sql = sql & "cast(100*sum(wins)/cast(sum(matches) as decimal(4)) as decimal(3,0)) as winrate "
	sql = sql & "from ( "
	sql = sql & "select datename(day,date) + ' ' + datename(month,date) as date, 1 as matches, case when goalsfor > goalsagainst then 1 else 0 end as wins "
	sql = sql & "from match  "
	sql = sql & ") as subsel  " 
	sql = sql & "group by date  " 
	sql = sql & "having sum(matches) > 5  "
	sql = sql & "order by winrate desc  " 
	rs.open sql,conn,1,2
		
	n = 0 
	Do While Not rs.EOF
	  tab1a(n,0) = rs.Fields("date")
	  tab1a(n,1) = rs.Fields("matches")
	  tab1a(n,2) = rs.Fields("winrate")
	  n = n + 1
	  rs.MoveNext
	Loop
	rs.close

	sql = "select  top 10 date, sum(matches) as matches, " 
	sql = sql & "cast((sum(goalsfor) - sum(goalsagainst))/cast(sum(matches) as decimal(4)) as decimal(3,2)) as diffrate "
	sql = sql & "from ( "
	sql = sql & "select datename(day,date) + ' ' + datename(month,date) as date, 1 as matches, goalsfor, goalsagainst "
	sql = sql & "from match  "
	sql = sql & ") as subsel  " 
	sql = sql & "group by date  " 
	sql = sql & "having sum(matches) > 5  "
	sql = sql & "order by diffrate desc  " 
	rs.open sql,conn,1,2
	
	n = 0 
	Do While Not rs.EOF
	  tab1b(n,0) = rs.Fields("date")
	  tab1b(n,1) = rs.Fields("matches")
	  tab1b(n,2) = rs.Fields("diffrate")
	  n = n + 1
	  rs.MoveNext
	Loop
	rs.close
	
	outline = outline & "<table id=""gottable1"" border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#c0c0c0"">"
    outline = outline & "<tr><td colspan=""6""><p class=""style1bold"">The Best of Dates</p>"
    outline = outline & "<p>The most successful dates in the club's history (excludes those with fewer than 5 games)"
   	outline = outline & "<span style=""font-size:10px"">"
   	outline = outline & "<br>P: Games played on this date"
   	outline = outline & "<br>%W: Percentage of games won"
   	outline = outline & "<br>Av GD: Average goal difference"
    outline = outline & "</span></p></td></tr>"
    outline = outline & "<tr><td colspan=""3""><p><b>Top 10  on Wins</b></td><td colspan=""3"" style=""padding-left: 8""><p><b>Top 10  on Goal Diff</b></td></tr>"
    outline = outline & "<tr><td><b>Good Date</b></td><td class=""right"">P</td><td class=""right"">%W</td><td style=""padding-left: 8""><b>Good Date</b></td><td class=""right"">P</td><td class=""right"">Av GD</td></tr>"
	for n = 0 to 9
		outline = outline & "<tr><td><a href=""gosdb-dates.asp?date=" & tab1a(n,0) & """>" & tab1a(n,0) & "</td><td class=""right"">" & tab1a(n,1) & "</td><td class=""right"">" & tab1a(n,2) & "</td><td style=""padding-left: 8""><a href=""gosdb-dates.asp?date=" & tab1b(n,0) & """>" & tab1b(n,0) & "</td><td class=""right"">" & tab1b(n,1) & "</td><td class=""right"">" & tab1b(n,2) & "</td></tr>"
	next
	outline = outline & "</table>"
	
    outline = outline & "^"

	sql = "select top 10 date, sum(matches) as matches, " 
	sql = sql & "cast(100*sum(losses)/cast(sum(matches) as decimal(4)) as decimal(3,0)) as lossrate "
	sql = sql & "from ( "
	sql = sql & "select datename(day,date) + ' ' + datename(month,date) as date, 1 as matches, case when goalsfor < goalsagainst then 1 else 0 end as losses "
	sql = sql & "from match  "
	sql = sql & ") as subsel  " 
	sql = sql & "group by date  " 
	sql = sql & "having sum(matches) > 5  "
	sql = sql & "order by lossrate desc  " 
	rs.open sql,conn,1,2
		
	n = 0 
	Do While Not rs.EOF
	  tab2a(n,0) = rs.Fields("date")
	  tab2a(n,1) = rs.Fields("matches")
	  tab2a(n,2) = rs.Fields("lossrate")
	  n = n + 1
	  rs.MoveNext
	Loop
	rs.close

	sql = "select  top 10 date, sum(matches) as matches, " 
	sql = sql & "cast((sum(goalsfor) - sum(goalsagainst))/cast(sum(matches) as decimal(4)) as decimal(3,2)) as diffrate "
	sql = sql & "from ( "
	sql = sql & "select datename(day,date) + ' ' + datename(month,date) as date, 1 as matches, goalsfor, goalsagainst "
	sql = sql & "from match  "
	sql = sql & ") as subsel  " 
	sql = sql & "group by date  " 
	sql = sql & "having sum(matches) > 5  "
	sql = sql & "order by diffrate  " 
	rs.open sql,conn,1,2
	
	n = 0 
	Do While Not rs.EOF
	  tab2b(n,0) = rs.Fields("date")
	  tab2b(n,1) = rs.Fields("matches")
	  tab2b(n,2) = rs.Fields("diffrate")
	  n = n + 1
	  rs.MoveNext
	Loop
	rs.close
	
	outline = outline & "<table id=""gottable1"" border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#c0c0c0"">"
    outline = outline & "<tr><td colspan=""6""><p class=""style1bold"">The Worst of Dates</p>"
    outline = outline & "<p>The least successful dates in the club's history (excludes those with fewer than 5 games)"
   	outline = outline & "<span style=""font-size:10px"">"
   	outline = outline & "<br>P: Games played on this date"
   	outline = outline & "<br>%L: Percentage of games lost"
   	outline = outline & "<br>Av GD: Average goal difference"
    outline = outline & "</span></p></td></tr>"
    outline = outline & "<tr><td colspan=""3""><p><b>Top 10  on Defeats</b></td><td colspan=""3"" style=""padding-left: 8""><p><b>Top 10  on Goal Diff</b></td></tr>"
    outline = outline & "<tr><td><b>Bad Date</b></td><td class=""right"">P</td><td class=""right"">%L</td><td style=""padding-left: 8""><b>Bad Date</b></td><td class=""right"">P</td><td class=""right"">Av GD</td></tr>"	
    for n = 0 to 9
		outline = outline & "<tr><td><a href=""gosdb-dates.asp?date=" & tab2a(n,0) & """><b>" & tab2a(n,0) & "</b></a></td><td class=""right"">" & tab2a(n,1) & "</td><td class=""right"">" & tab2a(n,2) & "</td><td style=""padding-left: 8""><a href=""gosdb-dates.asp?date=" & tab2b(n,0) & """><b>" & tab2b(n,0) & "</b></td><td class=""right"">" & tab2b(n,1) & "</td><td class=""right"">" & tab2b(n,2) & "</td></tr>"
	next
	outline = outline & "</table>"

conn.close
	
response.write(outline)
%>