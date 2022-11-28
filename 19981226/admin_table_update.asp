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
	border-collapse: collapse;
	}
td,th {
	border:1px solid #ddd;
	}
td {
	padding: 3px;
	}
th {
	padding: 5px;
	}
::placeholder {
	color:green;
	font-size:10px;
	}
-->
</style>
</head>

<body>

<% 
Dim output, phase, table_name, column, columns, column_list, filtercol, filter, readonly, old_value, new_value, rowID, administrator, logvalues, n
Dim conn, sql, sql1, sql2, rs 

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include virtual="/conn_read.inc"-->

<div id="container">
<!--#include file="admin_head.inc"-->

<%
phase = request.form("phase")
table_name = request.form("table")

' Get the table's column names for phases 3, 4 & 5
if phase > "2" then
	sql = "select column_name "
	sql = sql & "from information_schema.columns " 
	sql = sql & "where table_name = '" & table_name & "' "	
	sql = sql & "order by ordinal_position "

	rs.open sql,conn,1,2
	Do While Not rs.EOF
		if rs.Fields("column_name") <> "ID" then column_list = column_list & rs.Fields("column_name") & ","		'ignore the identity column
		rs.Movenext
	Loop
	rs.close
	column_list = left(column_list,len(column_list)-1)	'remove final comma	
end if

select case phase
	case "1"
		select case table_name
			case "FL_results"
					filtercol = "date "
					sql1 = "select date, home_team, away_team, home_goals, away_goals, ID "
					sql1 = sql1 & "from FL_results " 	
					sql2 = "order by date, home_team "
			case "match"
					filtercol = "date "
					sql1 = "select date, opposition, homeaway, ID "
					sql1 = sql1 & "from match " 	
					sql2 = "order by date "
			case "match_extra"
					filtercol = "date "
					sql1 = "select date, ID "
					sql1 = sql1 & "from match_extra " 	
					sql2 = "order by date "
			case "onthisday"
					filtercol = "month"
					sql1 = "select month, day, seqno, ID "
					sql1 = sql1 & "from onthisday " 	
					sql2 = "order by month, day, seqno "
			case "opposition"
					readonly = "name_then"
					filtercol = "name_then"
					sql1 = "select name_then, ID "
					sql1 = sql1 & "from opposition " 	
					sql2 = "order by name_then "
			case "player"
					readonly = "player_id"
					filtercol = "surname"
					sql1 = "select trim(surname) + ', ' + isNull(trim(forename),'?') + ' ' + cast(player_id as varchar(4)) as player, ID "
					sql1 = sql1 & "from player " 
					sql2 = "and player_id < 8000 "	
					sql2 = sql2 & "order by player "
			case "player_squad"
					readonly = "squad_no"
					filtercol = "surname"
					sql1 = "select season_no, ID "
					sql1 = sql1 & "from player_squad " 	
					sql2 = "order by season_no, squad_no "
			case "season"
					readonly = "years"
					filtercol = "years"
					sql1 = "select years, ID "
					sql1 = sql1 & "from season " 	
					sql2 = "order by years "
			case "season_this"
					filtercol = "opposition"
					sql1 = "select date, opposition, homeaway, ID "
					sql1 = sql1 & "from season_this " 	
					sql2 = "order by date "
			case "team_photo"
					filtercol = "years"
					sql1 = "select years, seq_no, ID "
					sql1 = sql1 & "from team_photo " 	
					sql2 = "order by years, seq_no "
			case "venue"
					filtercol = "club_name_then "
					sql1 = "select club_name_then, ground_name, first_game, last_game, ID "
					sql1 = sql1 & "from venue " 	
					sql2 = "order by club_name_then, first_game "
		end select		
		Call Phase1
	case "2"	
		Call Phase2
	case "3"	
		Call Phase3
	case "4"	
		Call Phase4
	case "5"	
		Call Phase5
	case else
		Call Phase0
end select

Response.write(output)
Conn.close

Sub Phase0

	output = output & "<form action=""admin_table_update.asp"" method=""post"">"
	output = output & "<table>"
	output = output & "<tr><th>Table</th><th>Column for Filter</th></tr>"
		
	output = output & "<tr><td><input type=""radio"" name=""table"" id=""season_this"" value=""season_this"">"
		output = output & "<label for=""season_this""> Fixtures</label></td>"
		output = output & "</td><td>Opposition</td></tr>"
	
	output = output & "<tr><td><input type=""radio"" name=""table"" id=""FL_results"" value=""FL_results"">"
		output = output & "<label for=""FL_results""> League Results</label></td>"
		output = output & "</td><td>Date</td></tr>"

	output = output & "<tr><td><input type=""radio"" name=""table"" id=""match"" value=""match"">"
		output = output & "<label for=""match""> Match</label></td>"
		output = output & "</td><td>Date</td></tr>"
			
	output = output & "<tr><td><input type=""radio"" name=""table"" id=""match_extra"" value=""match_extra"">"
		output = output & "<label for=""match_extra""> Match Extra</label></td>"
		output = output & "</td><td>Month & Day</td></tr>"
		
	output = output & "<tr><td><input type=""radio"" name=""table"" id=""onthisday"" value=""onthisday"">"
		output = output & "<label for=""onthisday""> On This Day</label></td>"
		output = output & "</td><td>Month</td></tr>"
		
	output = output & "<tr><td><input type=""radio"" name=""table"" id=""opposition"" value=""opposition"">"
		output = output & "<label for=""opposition""> Opposition</label></td>"
		output = output & "</td><td>Opposition</td></tr>"
				
	output = output & "<tr><td><input type=""radio"" name=""table"" id=""player"" value=""player"">"
		output = output & "<label for=""player""> Player</label></td>"
		output = output & "</td><td>Surname</td></tr>"
		
	output = output & "<tr><td><input type=""radio"" name=""table"" id=""player_squad"" value=""player_squad"">"
		output = output & "<label for=""player_squad""> Player_Squad</label></td>"
		output = output & "</td><td>Surname</td></tr>"

	output = output & "<tr><td><input type=""radio"" name=""table"" id=""season"" value=""season"">"
		output = output & "<label for=""season""> Season</label></td>"
		output = output & "</td><td>Years</td></tr>"
		
	output = output & "<tr><td><input type=""radio"" name=""table"" id=""team_photo"" value=""team_photo"">"
		output = output & "<label for=""team_photo""> Team Photo</label></td>"
		output = output & "</td><td>Years</td></tr>"
		
	output = output & "<tr><td><input type=""radio"" name=""table"" id=""venue"" value=""venue"">"
		output = output & "<label for=""venue""> Venue</label></td>"
		output = output & "</td><td>Opposition</td></tr>"
		
	output = output & "</table><br><br>"	
	
	output = output & "<input type=""radio"" name=""new-exist"" id=""new"" value=""New"">"
		output = output & "<label for=""new-exist""> Create new row</label><br><br>"
	output = output & "<input type=""radio"" name=""new-exist"" id=""exist"" value=""Exist"">"
		output = output & "<label for=""new-exist""> Amend existing row</label>"
		output = output & "<input style=""margin-left:10px;"" type=""text"" name=""filter"" id=""filter"" size=10 placeholder=""Optional filter""><br>"
			
	output = output & "<input type=""hidden"" name=""phase"" value=""1"">"			
	output = output & "<input style=""margin:30px 0;"" type=""submit"" value=""Next"">"
	output = output & "</form>"

End sub

Sub Phase1

	if request.form("table") = "" then 
		output = output & "<p class=""style1boldred"">Choose a table</p>"
		Call Backbutton
	  elseif request.form("new-exist") = "" then
		output = output & "<p class=""style1boldred"">Choose New or Update</p>"
	  else	
		output = output & "<form action=""admin_table_update.asp"" method=""post"">"
	
		if request.form("new-exist") = "New" then
			output = output & "<input type=""hidden"" name=""rowID"" value=""New"">"
			output = output & "<input type=""hidden"" name=""phase"" value=""3"">"
		  else
		  	filter = request.form("filter")
		  	if filter > "" then sql1 = sql1 & "where " & filtercol & " like '%" & filter & "%' "
		  	sql = sql1 & sql2
		  	output = output & "<input type=""hidden"" name=""phase"" value=""2"">"
		  	output = output & "<input type=""hidden"" name=""sql"" value=""" & sql & """>"	
		end if
	
		output = output & "<input type=""hidden"" name=""admin"" value=""" & administrator & """>"
		output = output & "<input type=""hidden"" name=""table"" value=""" & table_name & """>"	
		output = output & "<input type=""hidden"" name=""readonly"" value=""" & readonly & """>"	
		output = output & "<input style=""margin:30px 5px 0 0;"" type=""submit"" value=""Next"">"
		output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
		output = output & "</form>"
	end if
	
End sub

Sub Phase2

	administrator = request.form("admin") 
	table_name = request.form("table")
	readonly = request.form("readonly")
	sql = request.form("sql")
	
	output = output & "<form action=""admin_table_update.asp"" method=""post"">"	

 	rs.open sql,conn,1,2
	Do While Not rs.EOF
		output = output & "<input type=""radio"" name=""rowID"" id=""rowID" & rs.Fields("ID") & """ value=""" & rs.Fields("ID") & """>"
		output = output & "<label for=""rowID" & rs.Fields("ID") & """> "
		select case table_name
			case "FL_results"
					output = output & rs.Fields("date")	& ": " & rs.Fields("home_team") & " " & rs.Fields("home_goals")	& " - " & rs.Fields("away_goals") & " " & rs.Fields("away_team")
			case "match"
					output = output & rs.Fields("date")	& " " & rs.Fields("opposition") & " " & rs.Fields("homeaway")	
			case "match_extra"
					output = output & rs.Fields("date")
			case "onthisday"
					output = output & rs.Fields("month") & "-" & rs.Fields("day")  & " " & rs.Fields("seqno")
			case "opposition"
					output = output & rs.Fields("name_then")				
			case "player"
					output = output & rs.Fields("player")
			case "player_squad"
					output = output & rs.Fields("season_no")
			case "season"
					output = output & rs.Fields("years")
			case "season_this"
					output = output & rs.Fields("date")	& " " & rs.Fields("opposition") & " " & rs.Fields("homeaway")
			case "team_photo"
					output = output & rs.Fields("years") & " / " & rs.Fields("seq_no")
			case "venue"
					output = output & rs.Fields("club_name_then") & " / " & rs.Fields("ground_name") & " / " & rs.Fields("first_game") & " - " & rs.Fields("last_game")			
		end select
		output = output & "</label><br>"
		rs.Movenext
	Loop
	rs.close
		
	output = output & "<input type=""hidden"" name=""phase"" value=""3"">"
	output = output & "<input type=""hidden"" name=""admin"" value=""" & administrator & """>"
	output = output & "<input type=""hidden"" name=""table"" value=""" & table_name & """>"	
	output = output & "<input type=""hidden"" name=""readonly"" value=""" & readonly & """>"	
	output = output & "<input style=""margin:30px 5px 0 0;"" type=""submit"" value=""Next"">"
	output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
	output = output & "</form>"		
	
End sub

Sub Phase3

	administrator = request.form("admin")	
	table_name = request.form("table")
	readonly = request.form("readonly")
	rowID = request.form("rowID")
	
	output = output & "<table>"
	output = output & "<form action=""admin_table_update.asp"" method=""post"">"
	n = 1
	columns = split(column_list,",")
	
	if rowId = "New" then
	
		for each column in columns
			output = output & "<tr>"
			output = output & "<td>" & column & "</td>" 
			output = output & "<td><input type=""text"" name=""N" & n & """"
			if table_name = "player" and column = "player_id" then
				sql = "select max(player_id) + 1 as next_player_id from player where player_id < 8000 "
				rs.open sql,conn,1,2
				output = output & " value = """ & rs.Fields("next_player_id") & """ readonly" 
				rs.close
			end if
			output = output & "></td></tr>"
			n = n + 1
		next
		output = output & "</table>"

	  else   

		sql = "select " & column_list & " "
		sql = sql & "from " & table_name & " "
		sql = sql & "where ID = '" & rowID & "' "
		rs.open sql,conn,1,2
			for each column in columns
				output = output & "<tr>"
				output = output & "<td>" & column & "</td>" 
				output = output & "<td><input type=""text"" name=""N" & n & """ value="""
				if isNull(Eval("rs.Fields(""" & column & """)")) then 
					old_value = ""
				  else 
				  	old_value = Eval("rs.Fields(""" & column & """)")
				end if
				output = output & old_value & """"
				if column = readonly then output = output & " readonly"
				output = output & "></td></tr>"
				output = output & "<input type=""hidden"" name=""O" & n & """ value=""" & old_value & """>"
				n = n + 1
			next
		rs.close
		output = output & "</table>"
	
	end if
	
	output = output & "<input type=""hidden"" name=""phase"" value=""4"">"
	output = output & "<input type=""hidden"" name=""admin"" value=""" & administrator & """>"
	output = output & "<input type=""hidden"" name=""table"" value=""" & table_name & """>"
	output = output & "<input type=""hidden"" name=""rowID"" value=""" & rowID & """>"
	output = output & "<input style=""margin:30px 5px 0 0;"" type=""submit"" value=""Next"">"
	output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
	output = output & "</form>"

End sub

Sub Phase4

	administrator = request.form("admin")
	table_name = request.form("table")
	rowID = request.form("rowID")	
	
	output = output & "<table>"
	output = output & "<tr><th>Column</th><th>Old Value</th><th>New Value</th></tr>"
	output = output & "<form action=""admin_table_update.asp"" method=""post"">"
	
	n = 1
	columns = split(column_list,",")
	for each column in columns
		old_value = Eval("request.form(""O" & n & """)")
		new_value = Eval("request.form(""N" & n & """)")
		if old_value <> new_value then
			output = output & "<tr>"
			output = output & "<td>" & column & "</td>" 
			output = output & "<td>" & old_value & "</td>"
			output = output & "<td>" & new_value & "</td>"
			output = output & "</tr>"
			output = output & "<input type=""hidden"" name=""C" & n & """ value=""" & column & """>"
			output = output & "<input type=""hidden"" name=""O" & n & """ value=""" & old_value & """>"
			output = output & "<input type=""hidden"" name=""N" & n & """ value=""" & new_value & """>"
		end if
		n = n + 1
	next

	output = output & "</table>"
	output = output & "<input type=""hidden"" name=""phase"" value=""5"">"
	output = output & "<input type=""hidden"" name=""admin"" value=""" & administrator & """>"
	output = output & "<input type=""hidden"" name=""table"" value=""" & table_name & """>"
	output = output & "<input type=""hidden"" name=""rowID"" value=""" & rowID & """>"
	output = output & "<input style=""margin:30px 5px 0 0;"" type=""submit"" value=""Update"">"
	output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
	output = output & "</form>"

End sub

Sub Phase5

	administrator = request.form("admin")
	table_name = request.form("table")
	rowID = request.form("rowID")
	administrator = request.form("admin")
	
	conn.close
	%><!--#include file="conn_admin.inc"--><%
	
	if rowID = "New" then	'row to be inserted
	
		sql = "insert into " & table_name & " ("
		columns = split(column_list,",")
		for each column in columns
			sql = sql & column & ","
		next	
		sql = left(sql,len(sql)-1) & ") "	'replace last comma with close bracket and space
		
		sql = sql & "values("
		n = 1
		for each column in columns
			if len(trim(Eval("request.form(""N" & n & """)"))) = 0 then
				new_value = new_value & "NULL,"
			  else
				new_value = new_value & "'" & trim(replace(Eval("request.form(""N" & n & """)"),"'","''")) & "',"
			end if
			n = n + 1
		next
		new_value = left(new_value,len(new_value)-1)	'remove last comma
		sql = sql & new_value & ") "
		output = output & "<p class=""style1"">" & sql & "</p>"
		on error resume next
		conn.Execute sql
		if err = 0 then
			output = output & "<p class=""style1bold"" style=""color:green"">Row successfully inserted: """ & new_value & """ in " & table_name & "</p>"
			sql = "select ident_current('" & table_name & "') "
			set rs = conn.Execute(sql)  'retrieve the identity column value just inserted (wanted for display of the new row)
			rowID = rs(0)
			rs.close
			new_value = replace(new_value,"','","|")			'convert the insert values ...
			new_value = replace(new_value,"',NULL","|NULL")		'... to a single string ...
			new_value = replace(new_value,"NULL,'","NULL|")		'... for inserting into admin_log
			if right(new_value,4) = "NULL" then new_value = new_value & "'"	'terminate the string correctly if the last characters are "NULL" 
		 	logvalues = "'" & administrator & "','I','" & table_name & "',NULL,NULL," & new_value
		 	Call AddToLog
		  else
			Response.Write("<p class=""style1bold"" style=""color:red"">SQL ERROR!!<br>Statement: " & sql & "<br>Error: " & err.description & "</p>")
		end if		
		
	  else		'row to be updated (one update statement for each changed column)
	
		n = 1
		columns = split(column_list,",")
		for each column in columns
			if column = Eval("request.form(""C" & n & """)") then
				old_value = Eval("request.form(""O" & n & """)")
				new_value = Eval("request.form(""N" & n & """)")
				if old_value <> new_value then
					sql = "update " & table_name & " "
					sql = sql & "set " & column & " = "
					if len(trim(new_value)) = 0 then
						sql = sql & "NULL "
					  else
						sql = sql & "'" & trim(replace(new_value,"'","''")) & "' "
					end if
					sql = sql & "where ID = " & rowID
					if len(trim(old_value)) = 0 then
						sql = sql & " and " & column & " is null "
					  else
						sql = sql & " and " & column & " = '" & trim(replace(old_value,"'","''")) & "' "
					end if
					output = output & "<p class=""style1"">" & sql & "</p>"
					on error resume next
					conn.Execute sql
					if err = 0 then
						output = output & "<p class=""style1bold"" style=""color:green"">Column successfully updated. Value: """ & new_value & """ Column: " & column & " Table: " & table_name & "</p>"
					 	logvalues = "'" & administrator & "','U','" & table_name & "','" & column & "','" & trim(replace(old_value,"'","''")) & "','" & trim(replace(new_value,"'","''")) & "'"
					 	logvalues = replace(logvalues,"'NULL'","NULL")
				 		Call AddToLog
				  	  else
						Response.Write("<p class=""style1bold"" style=""color:red"">SQL ERROR!!<br>Statement: " & sql & "<br>Error: " & err.description & "</p>")
					end if
				end if	
			end if
			n = n + 1
		next
				
	end if
	
	On error GoTo 0
	
	' Insert/Update complete, now display this row
	
	output = output & "<table style=""max-width:800px;"">"
	sql = "select " & column_list & " "
	sql = sql & "from " & table_name & " "
	sql = sql & "where ID = '" & rowID & "' "
	rs.open sql,conn,1,2
		columns = split(column_list,",")
		for each column in columns
			output = output & "<tr>"
			output = output & "<td style=""font-size:11px;"">" & column & "</td>" 
			output = output & "<td style=""font-size:12px;"">"
			if isNull(Eval("rs.Fields(""" & column & """)")) then 
				new_value = ""
			  else 
			  	new_value = Eval("rs.Fields(""" & column & """)")
			end if
			output = output & new_value & "</td></tr>"
		next
	rs.close
	output = output & "</table>"
	
	output = output & "<form style=""margin-top:20px"" action=""admin_table_update.asp"">"
    output = output & "<button type=""submit"">Back to Start</button>"
    output = output & "</form>"

End sub

Sub AddToLog

	sql = "insert into admin_log (administrator, IRD, table_name, column_name, old_value, new_value) "
	sql = sql & "values(" & logvalues & ") "
	on error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<p class=""style1bold"" style=""color:blue"">SQL ERROR!<br>Statement: " & sql & "<br>Error: " & err.description)  & "</p>" 

End sub 

Sub Backbutton
	output = output & "<form>"
	output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
	output = output & "</form"
End sub
 
%>
</div>
</body>
</html>