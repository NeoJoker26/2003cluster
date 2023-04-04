<%
response.expires = -1
playerlist = Request.QueryString("p")
years = Request.QueryString("y")
scope = Request.QueryString("scp")
if instr(scope," or ") > 0 or instr(scope,"union ") > 0 or instr(scope,"drop ") > 0 or instr(scope,"=") > 0 then scope = ""

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

tagprefix = timer()
tagno = 1
							
	outline = "<table class=""matchlist"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""margin-left: 30; border-collapse: collapse; background-color: #e0f0e0"" >"					    
	outline  = outline & "<tr><td colspan=""8"" padding: 0 0 0 0;"">Match Details | Start or Sub | Goals Scored</td></tr>"	  
	sql = "with CTE as "
	sql = sql & "( "
	sql = sql & "select opposition, opposition_qual, a.date as date, homeaway, "
	sql = sql & "case when goalsfor > goalsagainst then 'W' when goalsfor < goalsagainst then 'L' else 'D' end as result, "
	sql = sql & "case when startpos = 0 then 'Sub' else 'Start' end as startsub, "
	sql = sql & "a.compcode as compcode, subcomp, goalsfor, goalsagainst, "
	sql = sql & "case when c.date is not null then 1 else 0 end as goals "
	sql = sql & "from v_match_all a " 
	sql = sql & " join season on date >= date_start and date <= date_end "
	sql = sql & " join match_player b on a.date = b.date "
	sql = sql & " left outer join match_goal c on a.date = c.date and b.player_id = c.player_id "
	sql = sql & "where b.player_id in (" & playerlist & ") " 
	sql = sql & "  and years = '" & years & "' "
	sql = sql & "  and compcat in (" & scope & ") "
	sql = sql & ") "
	sql = sql & "select opposition, opposition_qual, date, homeaway, result, startsub, compcode, subcomp, goalsfor, goalsagainst, sum(goals) as goals "
	sql = sql & "from CTE "
	sql = sql & "group by opposition, opposition_qual,  date, homeaway, result, startsub, compcode, subcomp, goalsfor, goalsagainst "
	sql = sql & "order by date "
	
	rs.open sql,conn,1,2
	Do While Not rs.EOF

		opposition = rs.Fields("opposition") 
		opposition = replace(opposition,"United","Utd")
		opposition = replace(opposition,"Rovers","Rvrs")
		opposition = replace(opposition,"Wanderers","Wnds")
		opposition = replace(opposition,"Albion","Alb")
		opposition = replace(opposition,"Athletic","Ath")
		opposition = replace(opposition,"County","Co")
		opposition = replace(opposition,"Wednesday","Wed")
		opposition = replace(opposition,"Avenue","Ave")
		opposition = replace(opposition,"Boscombe","Bos")
		opposition = replace(opposition,"Redbridge","Red")
		opposition = replace(opposition," and "," & ")
	
		displaydate = FormatDateTime(rs.Fields("date"),1)
		work1 = split(displaydate," ")
		displaydate = work1(0) & " " & left(work1(1),3) 
		goals = rs.Fields("goals")
		if goals = 0 then goals = " " 
		outline  = outline & "<td nowrap=""nowrap""><a style=""font-family:courier;"" id=""xmattag" & tagprefix & tagno & """ href=""javascript:Toggle3('mattag" & tagprefix & tagno & "');"">[+]</a></td>" 
		outline  = outline & "<td nowrap=""nowrap"">" & displaydate & "<span id=""dmattag" & tagprefix & tagno & """ style=""display:none;"">" & rs.Fields("date") & "</span></td>" 
		outline  = outline & "<td nowrap=""nowrap"">" & rs.Fields("compcode") & " " & rs.Fields("subcomp") & "</td>"
		outline  = outline & "<td nowrap=""nowrap"">" & opposition & " " & rs.Fields("opposition_qual") & " (" & rs.Fields("homeaway") & ")</td>"
		outline  = outline & "<td nowrap=""nowrap"">" & rs.Fields("result") & "</td>" 
		outline  = outline & "<td nowrap=""nowrap"">" & rs.Fields("goalsfor") & "-" & rs.Fields("goalsagainst") & "</td>" 
		outline  = outline & "<td nowrap=""nowrap"">" & rs.Fields("startsub") & "</td>" 
		outline  = outline & "<td nowrap=""nowrap"">" & goals & "</td>"				
		outline  = outline & "</tr>"
		outline  = outline & "<tr><td class=""a"" style=""padding: 0 0 0 0;""></td><td  class=""a"" colspan=""7"" width=""350px"" style=""font-size:10px; font-family: verdana,arial,helvetica,sans-serif; padding: 0 0 0 0;""><span id=""mattag" & tagprefix & tagno & """ style=""display:none; margin: 1 10 4 8; ""><img border=""0"" src=""images/ajax-loader.gif""></span></td></tr>"	  
		tagno = tagno + 1												
		rs.Movenext
	loop
	rs.close
							 
	outline = outline & "</table>"
							
response.write(outline)
%>