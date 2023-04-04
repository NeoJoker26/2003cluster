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

<% Dim teamlist(23,1), i, j, output, phase, administrator, results, resultparta, resultpartb, insertvalues, allinsertvalues, insertstmt
Dim resultdate, attend, hometeam, awayteam, homescore, awayscore, logvalues

Dim conn,sql,rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_admin.inc"--><%
%>

<div id="container">
<!--#include file="admin_head.inc"-->

<h3 style="margin:6px 0 15px;">UPDATE LEAGUE TABLE</h3>

<% 

phase = request.form("phase")

select case phase
	case 1
		Call Format
	case 2
		Call Insert
	case else
		Call Input
end select

Response.write(output)
Conn.close

Sub Input

	output = "<form action=""league_table_update.asp"" method=""post"">"

	output = output & "<p style=""margin-top: 12; margin-bottom: 6;"">Important! Ensure correct date for results:</p>"
	resultdate = year(Date) & "-" & right("0" & month(Date),2) & "-" & right("0" & day(Date),2)
	output = output & "<p style=""margin-top: 0; margin-bottom: 3;""><input type=""text"" name=""date"" size=""10"" value=""" & resultdate & """></p>"
	output = output & "<p style=""margin-top: 12; margin-bottom: 6;""><a target=""_blank"" href=""http://www.espn.co.uk/football/fixtures"">Copy results from here</a>, then paste here:</p>"
	output = output & "<p style=""margin-top: 0; margin-bottom: 3;""><textarea rows=""38"" name=""results"" cols=""100"" wrap=""wrap""></textarea></p>" 
  
	output = output & "<input type=""hidden"" name=""phase"" value=""1"">"			
	output = output & "<input style=""margin:30px 0;"" type=""submit"" value=""Format Results"">"
	output = output & "</form>"

End Sub

Sub Format

	results = Request.Form("results")
	resultdate = Request.Form("date")

	Do While InStr(1,results,"  ")
		results = Replace(results,"  ", " ")	'Replace 2 consecutive spaces with a single space
	Loop
	
	sql = "select name_now, name_espn "
	sql = sql & "from opposition "  
	sql = sql & "where this_season = 'Y' "
	sql = sql & "union all "
	sql = sql & "select name_now, name_espn "
	sql = sql & "from opposition_dummy_for_PAFC "  
	sql = sql & "order by name_now "
	rs.open sql,conn,1,2

	i = 0

	Do While Not rs.EOF
		teamlist(i,0) = rs.Fields("name_espn")
		teamlist(i,1) = rs.Fields("name_now")
		i = i + 1			
		rs.MoveNext
	Loop
	
	rs.close
	
	output = "<form action=""league_table_update.asp"" method=""post"">" 
	
	resultparta = split(results,vbCrLf)
		
	for i = 0 to ubound(resultparta) step 3 
	
		' the result is in three lines:
		' 1. (i) the home team
		' 2. (i+1) the score
		' 3. (i+2) the away team
		
		hometeam = ""
		awayteam = ""
		homescore = ""
		awayscore = ""
		attend = "0"	
	
		for j = 0 to 23
			'check if the home team name matches an espn name (if one exists) or a full team name 
			if instr(resultparta(i),teamlist(j,0)) > 0 or instr(resultparta(i),teamlist(j,1)) > 0 then 
		  		hometeam = teamlist(j,1)
				exit for
			end if
		next
		
		resultparta(i+1) = trim(resultparta(i+1))
		
		resultpartb = split(resultparta(i+1),"-") 	'split the score
		homescore = trim(resultpartb(0))
		awayscore = trim(resultpartb(1))
		
		if left(resultparta(i+2),1) = chr(9) then resultparta(i+2) = mid(resultparta(i+2),1)	'remove first tab on the away team line
		
		resultpartb = split(resultparta(i+2),chr(9)) 	'split on tab
		
		for j = 0 to 23
			'check if the away team name matches an espn name (if one exists) or a full team name 
			if instr(resultpartb(1),teamlist(j,0)) > 0 or instr(resultpartb(1),teamlist(j,1)) > 0 then 
		  		awayteam = teamlist(j,1)
				exit for
			end if
		next
		
		attend = resultpartb(Ubound(resultpartb))		'last value in the line
		attend = replace(attend,",","")					'remove comma in attendance
		if not IsNumeric(attend) then attend = "NULL"
				
		insertvalues = "'" & resultdate & "','" & hometeam & "','" & awayteam & "'," & homescore & "," & awayscore & "," & attend
		output = output & "<p style=""margin-top: 0; margin-bottom: 3;"">" & insertvalues & "</p>"
		allinsertvalues = allinsertvalues & insertvalues & "|"
						
	next
	
	output = output & "<input type=""hidden"" name=""allinsertvalues"" value=""" & allinsertvalues & """>"	
	output = output & "<input type=""hidden"" name=""phase"" value=""2"">"		
	output = output & "<input style=""margin:30px 5px 0 0;"" type=""submit"" value=""Store Results"">"
	output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
	output = output & "</form>"

End Sub 

Sub Insert

	allinsertvalues = Request.Form("allinsertvalues")
	allinsertvalues = left(allinsertvalues,len(allinsertvalues)-1) 	'remove the last '|' as it messes with the following split
	insertvalues = split(allinsertvalues,"|")
	
	for each insertstmt in insertvalues
		sql = "insert into FL_results values(" & insertstmt & ");"
		on error resume next
		conn.Execute sql
		on error goto 0
		if err = 0 then
			insertstmt = replace(insertstmt,"','","|")
			insertstmt = replace(insertstmt,",'","|")
			insertstmt = replace(insertstmt,"'","|")
			output = output & "<p class=""style1bold"" style=""color:green"">Result successfully loaded: " & insertstmt & " in 'FL_results'</p>" 
	 		logvalues = "'" & administrator & "','I','FL_results',NULL,NULL,'" & insertstmt & "'"
		 	Call AddToLog
		  else
			Response.Write("<p class=""style1bold"" style=""color:red"">SQL ERROR!!<br>Statement: " & sql & "<br>Error: " & err.description & "</p>")
		end if
	next
	
	output = output & "<p class=""style4"" class=""width:400px, text-align:center,"">League Table Plus has been updated. <a target=""_blank"" href=""http://www.greensonscreen.co.uk/progresstables.asp?maint=y"">Check Table</a>.</p>"

End Sub

Sub AddToLog

	sql = "insert into admin_log (administrator, IRD, table_name, column_name, old_value, new_value) "
	sql = sql & "values(" & logvalues & "); "

	on error resume next
	conn.Execute sql
	on error goto 0

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