<%
Dim conn,sql,rs,rsrec,n,work1,work2,playerid,fs,photo,photoname,years

response.expires = -1
playerid = Request.QueryString("playerid")
if len(playerid) > 4 then player_id = 1
years = Request.QueryString("years")
if instr(years," or ") > 0 or instr(years,"union ") > 0 or instr(years,"drop ") > 0 or instr(years,"=") > 0 then years = ""
 
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rsrec = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

outline = "<table border=""0"" cellpadding=""5"" cellspacing=""0"" style=""border-collapse: collapse"" width=""380"">"
outline = outline & "<tr><td style=""padding-left: 18px;"" valign=""top"" width=""50%"">"

	sql = "select a.player_id, a.player_id_spell1, a.surname, a.forename, a.initials, a.dob, b.prime_photo "
	sql = sql & "from player a right outer join player b on a.player_id_spell1 = b.player_id "
	sql = sql & "where a.player_id = '" & playerid & "'"
	rs.open sql,conn,1,2

	if len(rs.Fields("player_id_spell1")) < 4 then 
		photoname = right("00" & rs.Fields("player_id_spell1"),3)
	  else
	  	photoname = rs.Fields("player_id_spell1")
	end if
	
	if not IsNull(rs.Fields("prime_photo")) then photoname = photoname & "_" & rs.Fields("prime_photo")

	photoname = photoname & ".jpg"

	outline = outline & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" " 
	outline = outline & "style=""border-collapse: collapse; margin: 12 0 8 0; color: #ffffff; background-color: #404040;"" width=""100%"">"
	outline = outline & "<tr><td width=""100%""><p style=""margin: 1 0 1 0; font-size:13px; font-weight:bold;"">"
	if IsNull(rs.Fields("forename")) then
		outline = outline & rs.Fields("initials") & " " & UCase(trim(rs.Fields("surname")))
	  else
	  	outline = outline & UCase(rs.Fields("forename")) & " " & UCase(trim(rs.Fields("surname")))
	end if
	
	outline = outline & "</p></td><td>"
	outline = outline &  "</p></td></tr></table>"	
	
	
		    sql = "select years, date_start, sum(St) as starts, sum (Su) as subs, sum(G) as goals "
    		sql = sql & " from ( "
    		sql = sql & "select years, date_start, "
    		sql = sql & "case when startpos > 0 then 1 else 0 end as St, "
    		sql = sql & "case when startpos = 0 then 1 else 0 end as Su, "
    		sql = sql & "0 as G "
    		sql = sql & "from season a join match_player b on date between date_start and date_end "
			sql = sql & " join player c on b.player_id = c.player_id "
			sql = sql & "where c.player_id = '" & playerid & "' and years = '" & years & "'"
    		sql = sql & "union all "
			sql = sql & "select years, date_start, "
    		sql = sql & "0, 0, 1 "
    		sql = sql & "from season a join match_goal b on date between date_start and date_end "
			sql = sql & " join player c on b.player_id = c.player_id "
			sql = sql & "where c.player_id = '" & playerid & "' and years = '" & years & "'"
			sql = sql & ") as subsel "
			sql = sql & "group by years, date_start "

			rsrec.open sql,conn,1,2
			
			if not IsNull(rs.Fields("dob")) then 
				outline = outline & "<p style=""margin-bottom:6;""><b>Age at start of season:</b> " & int(datediff("d",rs.Fields("dob"),rsrec.Fields("date_start"))/365.25) & "</p>"
			end if

			outline = outline & "<p style=""margin-bottom:3;""><b>In " & years & ":</b></p>"
			outline = outline & "<p style=""margin-bottom:3;""><b>Starts:</b> " & rsrec.Fields("starts") & "</p>" 
			if years > "1964-1965" then outline = outline & "<p style=""margin-bottom:2;""><b>Subs:</b> " & rsrec.Fields("subs") & "</p>" 
			outline = outline & "<p style=""margin-bottom:3;""><b>Goals:</b> " & rsrec.Fields("goals") & "</p>" 
			
			outline = outline & "<p style=""margin-top:12;margin-bottom:6;""><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & "&scp=1,2,3,4,5,6,7"">Go to his full record on the Players' Pages</a></p>"	
				
			rsrec.close

	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set photofolder = fs.GetFolder(Server.MapPath("/gosdb/photos/players/small"))

	if (fs.FileExists(photofolder & "/" & photoname)) <> true then photoname = "nophoto.jpg"
	outline = outline & "</td><td  style=""text-align: center"" valign=""top"" width=""50%""><img border=""0"" src=""gosdb/photos/players/small/" & photoname & """ align=""right""></td>"

rs.close
conn.close

outline = outline & "</tr></table>"
	
response.write(outline)
%>