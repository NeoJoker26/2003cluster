<%@ Language=VBScript %>
<% Option Explicit %>

<%
Dim conn, sql, rs, id1, id2, output1, output2, penpic, source, manager2clause, debutlist, lastgamelist, fullname, seasonhold
  
id1 = Request.QueryString("id1")
id2 = Request.QueryString("id2")
source = Request.QueryString("source")

output1 = "<img class=""close"" style=""margin: 9px -3px 8px 12px; float: right; border: 0"" src=""images/close.png"">"
 
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

sql = "select rtrim(forename) + ' ' + surname as fullname, photo_name, penpic "   
sql = sql & "from manager " 
sql = sql & "where manager_id = " & id1

rs.open sql,conn,1,2

if not isnull(rs.Fields("photo_name")) then output2 = output2 & "<img style=""margin: 6px 12px 6px 0; float: left; border: 0;"" src=""images/managers/" & rs.Fields("photo_name") & """>"

if source = "player" then 
	output1 = output1 & "<a href=""gosdb-managers.asp"" style=""margin:9px 0""><u>Go to All Managers</u></a>"
	output1 = output1 & "<p style=""font-size:16px; color:#202020; margin:9px 0"">Manager: " & rs.Fields("fullname")
end if
if id2 <> 999 then output2 = output2 & "<p class=""style1bold"" style=""text-align:left; margin: 18px 0 0;"">" & rs.Fields("fullname") & "</p>"

penpic = rs.Fields("penpic")
output2 = output2 & replace(penpic,"|p|","</p><p class=""style1"" style=""margin: 8px 0 6px 0; text-align: justify;"">") & "</p>"
output2 = replace(output2,"|i|","<span style=""font-style:oblique"">")
output2 = replace(output2,"|/i|","</span>")

rs.close

if id2 <> 999 then

	sql = "select rtrim(forename) + ' ' + surname as fullname, photo_name, penpic "   
	sql = sql & "from manager " 
	sql = sql & "where manager_id = " & id2

	rs.open sql,conn,1,2
	
	if not isnull(rs.Fields("photo_name")) then output2 = output2 & "<img style=""margin: 12px 12px 6px 0; clear: both; float: left; border: 0;"" src=""images/managers/" & rs.Fields("photo_name") & """>"

	if source = "player" then output1 = output1 & " and " & rs.Fields("fullname")
	output2 = output2 & "<p class=""style1bold"" style=""text-align:left; margin: 18px 0 0;"">" & rs.Fields("fullname") & "</p>"

	penpic = rs.Fields("penpic")

	output2 = output2 & replace(penpic,"|p|","</p><p class=""style1"" style=""margin: 8px 0 6px 0; text-align: justify;"">") & "</p>"
	output2 = replace(output2,"|i|","<span style=""font-style:oblique"">")
	output2 = replace(output2,"|/i|","</span>")

	rs.close

end if

if id2 = 999 then
	manager2clause = " and manager_id2 is null "
  else
	manager2clause = " and manager_id2 = " & id2
end if

sql = "select 1 as type, years, initials, forename, surname, b.player_id_spell1, manager_id1, manager_id2 "
sql = sql & "from match_player a join player b on a.player_id = b.player_id "
sql = sql & " join season c on a.date >= c.date_start and a.date<= c.date_end "
sql = sql & " join manager_spell d on a.date >= d.from_date and a.date <= isnull(d.to_date,getdate()) "
sql = sql & " where manager_id1 = " & id1 & manager2clause		
sql = sql & " and a.date = ( "
sql = sql & "	select min(date) "
sql = sql & "	from match_player x	join player y on x.player_id = y.player_id "
sql = sql & "	where y.player_id_spell1 = b.player_id_spell1 "
sql = sql & "	 ) "
sql = sql & "union all "
sql = sql & "select 2, years, initials, forename, surname, b.player_id_spell1, manager_id1, manager_id2 "
sql = sql & "from match_player a join player b on a.player_id = b.player_id "
sql = sql & " join season c on a.date >= c.date_start and a.date <= c.date_end "
sql = sql & " join manager_spell d on a.date >= d.from_date and a.date <= isnull(d.to_date,getdate()) "
sql = sql & " where manager_id1 = " & id1 & manager2clause
sql = sql & " and not exists (select * from player z where z.player_id_spell1 = b.player_id_spell1 "	'ignore if currently at the club
sql = sql & " 		          and last_game_year = 9999) "	
sql = sql & " and a.date = ( "
sql = sql & "	select max(date) "
sql = sql & "	from match_player x join player y on x.player_id = y.player_id "
sql = sql & "	where y.player_id_spell1 = b.player_id_spell1 "
sql = sql & "	  and y.last_game_year < 9999 "
sql = sql & "	 ) "
sql = sql & "order by type,years, surname, initials "

rs.open sql,conn,1,2

debutlist = ""
lastgamelist = ""

Do While Not rs.EOF

	if not isnull(rs.Fields("forename")) then
		fullname = trim(rs.Fields("forename")) & " " & trim(rs.Fields("surname"))
	  else
		fullname = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname"))	
	end	if
	
	if rs.Fields("type") = 1 then
		
		if rs.Fields("years") <> seasonhold then
			if seasonhold > "" then debutlist = left(debutlist,len(debutlist)-2) & " "	'drop last comma and space, then add space
			debutlist = debutlist & "<span style=""color:#606060; font-weight:bold"">" & rs.Fields("years") & "</span>: "
			seasonhold = rs.Fields("years") 
		end if	
		debutlist = debutlist & "<a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """ target=""_blank"">" & fullname & "</a>, " 
	  else
	  	if lastgamelist = "" then seasonhold = ""
	  	if rs.Fields("years") <> seasonhold then
	  		if seasonhold > "" then lastgamelist = left(lastgamelist,len(lastgamelist)-2) & " "	'drop last comma and space, then add space
			lastgamelist = lastgamelist & "<span style=""color:#606060; font-weight:bold"">" & rs.Fields("years") & "</span>: "
			seasonhold = rs.Fields("years")
		end if
		lastgamelist = lastgamelist & "<a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """ target=""_blank"">" & fullname & "</a>, " 				
	end if
		
	rs.MoveNext
	
Loop
rs.close

if debutlist > "" then 
	debutlist = left(debutlist,len(debutlist)-2) & "."				'drop last comma and space, and add full stop
  else
  	debutlist = "None"
end if

if lastgamelist > "" then 
	lastgamelist = left(lastgamelist,len(lastgamelist)-2) & "."		'drop last comma and space, and add full stop
  else
  	lastgamelist = "None"
end if

output2 = output2 & "<p class=""playerlisthead"">Players given their first-team debut: </p>"
output2 = output2 & "<p class=""playerlist"">" & debutlist & "</p>"
output2 = output2 & "<p class=""playerlisthead"">Final games played: </p>"
output2 = output2 & "<p class=""playerlist"">" & lastgamelist & "</p>"

conn.close	

response.write(output1 & "</p>" & output2)

%>