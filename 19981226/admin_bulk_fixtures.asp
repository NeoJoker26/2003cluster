<%@ Language=VBScript %> 
<% Option Explicit %>
<!DOCTYPE html>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>GoS Admin</title>

<head>
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
#container td {
  border: 1px solid #ddd;
  padding: 3px;
}
#container table {
  margin: 15px auto;
}
-->
</style>
</head>

<body>

<div id="container">
<h2 style="margin:15px 0 6px; text-align:center;">GREENS ON SCREEN ADMINISTRATION</h2>
<h3 style="margin:6px 0 15px; text-align:center;">LOAD NEXT SEASON'S FIXTURES</h3>

<% 
Dim output, phase, insertstring, administrator, fixturelist, fixtures, fixture, fixtureparts, fixturepart, insertstmts, insertstmt, insertstmtpart, insertstmtpart2, logvalues, n
Dim conn, sql, rs 

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

phase = request.form("phase")

select case phase
	case "1"		
		Call Phase1
	case "2"	
		Call Phase2
	case "3"	
		Call Phase3
	case else
		Call Phase0
end select

Response.write(output)

Sub Phase0

	output = "<p class=""style1"" style=""width:325px;""><b>IMPORTANT:</b> This page is for the bulk loading the League fixtures, normally published in June, and must not be used to add further games (e.g. Cup matches) or amend existing details. Use the geneble admin page for those purposes.</p>"

	output = output & "<form action=""admin_bulk_fixtures.asp"" method=""post"">"
	output = output & "Name: <input type=""text"" name=""admin"" id=""admin"">"
	output = output & "<input type=""hidden"" name=""phase"" value=""1"">"			
	output = output & "<input style=""margin:18px auto;"" type=""submit"" value=""Continue"">"
	output = output & "</form>"

End sub

Sub Phase1
 
 	administrator = Request.Form("admin")

	output = output & "<p class=""style4"" class=""width:400px, text-align:left"">Paste fixtures here: </p>"
	
	output = output & "<form action=""admin_bulk_fixtures.asp"" method=""post"">"
	output = output & "<textarea rows=""50"" cols=""60"" name=""fixturelist""></textarea>"
	output = output & "<p class=""style4"">Press Continue to load the season's fixtures</p>"
	
	output = output & "<input type=""hidden"" name=""phase"" value=""2"">"
	output = output & "<input type=""hidden"" name=""admin"" value=""" & administrator & """>"
	
	output = output & "<input style=""margin:18px auto;"" type=""submit"" value=""Continue"">"
	output = output & "</form>"		
	
End sub

Sub Phase2

	administrator = Request.Form("admin")
	fixturelist = Request.Form("fixturelist")
	fixturelist = replace(fixturelist,Chr(10),"^")		'replace end of line character with ^
	fixturelist = replace(fixturelist,Chr(13),"^")		'replace end of line character with ^
	fixturelist = replace(fixturelist,"^^","^")			'ensure only one ^ at end of line
	if right(fixturelist,1) = "^" then fixturelist = left(fixturelist,len(fixturelist)-1) 	'remove the last ^ as it messes with the following split
	fixtures = split(fixturelist,"^")					'fixtures will have been created with a ^ in between each one
	
	output = output & "<table>"
 	for each fixture in fixtures
		output = output & "<tr>"
		insertstring = insertstring & "insert into season_this (date,opposition,homeaway,compcode) values("
		fixtureparts = split(fixture,",")
		for each fixturepart in fixtureparts
			output = output & "<td>" & fixturepart & "</td>"
			insertstring = insertstring & "'" & fixturepart & "',"		
		next
		output = output & "</tr>"
		insertstring = left(insertstring,len(insertstring)-1) & ");"
	next		 
	output = output & "</table>"
		
	output = output & "<form action=""admin_bulk_fixtures.asp"" method=""post"">"
	output = output & "<input type=""hidden"" name=""phase"" value=""3"">"
	output = output & "<input type=""hidden"" name=""admin"" value=""" & administrator & """>"
	output = output & "<input type=""hidden"" name=""insertstring"" value=""" & insertstring & """>"
	output = output & "<input style=""margin:18px auto;"" type=""submit"" value=""Delete all old ones and add the new season's fixtures"">"
	output = output & "</form>"

End sub

Sub Phase3

	administrator = request.form("admin")
	insertstring = request.form("insertstring")
	
	%><!--#include file="conn_admin.inc"--><%
	
	sql = "delete from season_this "
	on error resume next
	conn.Execute sql
	if err = 0 then
		response.write("<p class=""style1bold"" style=""color:green"">Old fixtures deleted</p>")
	 	logvalues = "'" & administrator & "','D','season_this',NULL,'All rows in table',NULL"
	 	Call AddToLog
	  else
		Response.Write("<p class=""style1bold"" style=""color:red"">SQL ERROR!!<br>Statement: " & sql & "<br>Error: " & err.description & "</p>")
	end if		

	insertstring = left(insertstring,len(insertstring)-1) 	'remove the last ';' as it messes with the following split
	insertstmts = split(insertstring,";")
	for each insertstmt in insertstmts
		sql = insertstmt
		on error resume next
		conn.Execute sql
		if err = 0 then
			insertstmtpart = split(insertstmt,"values(")
			insertstmtpart2 = insertstmtpart(1)
			insertstmtpart2 = left(insertstmtpart2,len(insertstmtpart2)-1)
			insertstmtpart2 = replace(insertstmtpart2,"','","|")
			output = output & "<p class=""style1bold"" style=""color:green"">Fixture successfully loaded: " & insertstmtpart2 & " in 'season_this'</p>" 
	 		logvalues = "'" & administrator & "','I','season_this',NULL,NULL," & insertstmtpart2 
		 	Call AddToLog
		  else
			Response.Write("<p class=""style1bold"" style=""color:red"">SQL ERROR!!<br>Statement: " & sql & "<br>Error: " & err.description & "</p>")
		end if
	next
	
	On error GoTo 0
	
	' Inserts complete, now display all rows in table
	output = output & "<table>"
	n = 0
	sql = "select date, opposition, homeaway, compcode "
	sql = sql & "from season_this "
	sql = sql & "order by date "

	rs.open sql,conn,1,2
		
		Do While Not rs.EOF
			output = output & "<tr>"
			output = output & "<td>" & rs.Fields("date") & "</td>" 
			output = output & "<td>" & rs.Fields("opposition") & "</td>"
			output = output & "<td>" & rs.Fields("homeaway") & "</td>"
			output = output & "<td>" & rs.Fields("compcode") & "</td>"	
			output = output & "</tr>"
			n = n + 1
			rs.MoveNext					
		Loop
		
	rs.close
	Conn.close
	
	output = output & "</table>"
	output = output & "<p class=""style4"" class=""width:400px, text-align:center,"">" & n & " fixtures have been successfully loaded</p>"

End sub

Sub AddToLog

	sql = "insert into admin_log (administrator, IRD, table_name, column_name, old_value, new_value) "
	sql = sql & "values(" & logvalues & ") "
	on error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<p class=""style1bold"" style=""color:blue"">SQL ERROR!<br>Statement: " & sql & "<br>Error: " & err.description)  & "</p>" 

End sub 

%>
</div>
</body>
</html>