<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<%
Dim conn,sql,rs, n, work1, work2, scope, initial, last_player_id_spell1, namehold, surnamehold, yearshold, last_game_year, scotind, linkto, height, leftmargin, camefrom, fromtables, username

response.expires = -1

camefrom = Request.Querystring("camefrom")
username = Request.Querystring("username")
if camefrom = "gosdb-playerupdate.asp" then
	height = "20"
	leftmargin = "0"
	playertables = ", b.name from player a left outer join contributor b on a.penpic_pending_author = b.name "
  else
  	height = "300"
  	leftmargin = "44px"
  	playertables = "from player a "
  	camefrom = "gosdb-players2.asp"
end if  	

scope = Request.Querystring("scp")
if instr(scope," or ") > 0 or instr(scope,"union ") > 0 or instr(scope,"drop ") > 0 or instr(scope,"=") > 0 then scope = ""
initial = Request.QueryString("initial")
initial = replace(initial,"'","''")
if instr(initial," or ") > 0 or instr(initial," union ") > 0 or instr(initial,"drop ") > 0 or instr(initial,"=") > 0 then initial = ""

scotind = ""
if ucase(left(initial,2)) = "MC" then 
	initial2 = "Mac" & right(initial, len(initial)-2)
	scotind = "y"
end if
if ucase(left(initial,3)) = "MAC" then 
	initial2 = "Mc" & right(initial, len(initial)-3)
	scotind = "y"
end if

if initial = "" then
	outline = "<img border=""0"" src=""images/dummbar_0.gif"" height=""" & height & """ width=""1"">"
  else
	Set conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	%><!--#include file="conn_read.inc"--><%

 	if scotind = "y" then
 		sql = "select player_id, player_id_spell1, surname, aka_surname, a.initials, first_game_year, last_game_year "
		sql = sql & playertables
		sql = sql & "where (surname like '" & initial & "%' or surname like '" & initial2 & "%' "
		sql = sql & "  or aka_surname like '" & initial & "%' or aka_surname like '" & initial2 & "%') "
		sql = sql & "order by surname, initials, player_id_spell1, spell "
	  else
	   	sql = "select player_id, player_id_spell1, surname, aka_surname, a.initials, first_game_year, last_game_year "
		sql = sql & playertables
		sql = sql & "where (surname like '" & initial & "%' "
		sql = sql & "  or aka_surname like '" & initial & "%') "
		sql = sql & "order by surname, initials, player_id_spell1, spell "
	end if
	  
	rs.open sql,conn,1,2
	
	if rs.RecordCount = 0 then
		outline = "<p style=""margin: 9px 0 4px " & leftmargin & "; color:red"">No players begin with these letters</p><img border=""0"" src=""images/dummbar_0.gif"" height=""" & height & """ width=""1"">"
	  else
		outline = "<p style=""margin: 0 0 4px " & leftmargin & "; margin-bottom: 4px"">Name & first-last game per spell at club"
		last_player_id_spell1 = 0
		
		Do While Not rs.EOF
			if rs.Fields("player_id_spell1") <> last_player_id_spell1 then		'a new name
				if last_player_id_spell1 <> 0 then outline = outline & namehold & yearshold & "</p>" 	'finish off last name
				if isnull(rs.Fields("aka_surname")) then
					surnamehold = trim(rs.Fields("surname"))
				  else
				  	surnamehold = trim(rs.Fields("surname")) & " (aka " & trim(rs.Fields("aka_surname")) & ")"
				end if
				if camefrom = "gosdb-playerupdate.asp" then
					if trim(rs.Fields("name")) = username or isnull(rs.Fields("name")) then
						namehold = "<p style=""margin: 0""><a href=""" & camefrom & "?pid=" & rs.Fields("player_id_spell1") & "&scp=" & scope & """>" & surnamehold & ", " & trim(rs.Fields("initials")) & "</a>&nbsp;&nbsp;"
		  		  	  else
		  				namehold = "<p style=""margin: 0"">" & surnamehold & ", " & trim(rs.Fields("initials")) & "&nbsp;&nbsp;"
					end if
				  else	
				  	namehold = "<p style=""margin-left:44""><a href=""" & camefrom & "?pid=" & rs.Fields("player_id_spell1") & "&scp=" & scope & """>" & surnamehold & ", " & trim(rs.Fields("initials")) & "</a>&nbsp;&nbsp;" 
  				end if

				yearshold = ""
				last_player_id_spell1 = rs.Fields("player_id_spell1")
			end if
			last_game_year =  rs.Fields("last_game_year") 
			if last_game_year = 9999 then last_game_year = "present"
			if yearshold > "" then yearshold = yearshold & ", "
			if rs.Fields("first_game_year") = last_game_year then 
			 	yearshold = yearshold & rs.Fields("first_game_year") 
			  else
			 	yearshold = yearshold & rs.Fields("first_game_year") & "-" & last_game_year 
			end if
		rs.MoveNext
		Loop
		
		outline = outline & namehold & yearshold & "</p>"		'last entry

	end if
	rs.close
	conn.close
	
end if
	
response.write(outline)
%>