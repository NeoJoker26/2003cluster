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

<% Dim i, output, phase, administrator, logvalues, attend, lastline

Dim conn,sql,rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_admin.inc"--><%
%>

<div id="container">
<!--#include file="admin_head.inc"-->

<h3 style="margin:6px 0 15px;">UPDATE MISSING LEAGUE ATTENDANCES</h3>

<% 

phase = request.form("phase")

select case phase
	case 1
		Call Update
	case else
		Call Display
end select

Response.write(output)
Conn.close

Sub Display

	output = "<form action=""update_missing_attendances.asp"" method=""post"">"
	
	sql = "select date, home_team, away_team "
	sql = sql & "from FL_results "  
	sql = sql & "where attendance is null and date > '2005-08-01' "
	sql = sql & "order by date, home_team "
	rs.open sql,conn,1,2
	
	i = 0
	
	Do While Not rs.EOF
		output = output & "<p style=""margin-top: 0; margin-bottom: 3;"">" & rs.Fields("date") & ": " & rs.Fields("home_team") & " v " & rs.Fields("away_team")
		output = output & " <input type=""hidden"" name=""date" & i & """ value=""" & rs.Fields("date") & """><input type=""hidden"" name=""home" & i & """ value=""" & rs.Fields("home_team") & """>"
		output = output & "<input type=""text"" name=""attend" & i & """ size=""5"" style=""margin-left:10px;""></p>"
		i = i + 1		
		rs.MoveNext
	Loop
	
	rs.close

	output = output & "<input type=""hidden"" name=""phase"" value=""1"">"	
	output = output & "<input type=""hidden"" name=""lastline"" value=""" & i-1 & """>"			
	output = output & "<input style=""margin:15px 0;"" type=""submit"" value=""Update Attendance(s)"">"
	output = output & "</form>"

End Sub

Sub Update

	lastline = request.form("lastline")
		
	for i = 0 to lastline
	
		attend = request.form("attend" & i)
		attend = replace(attend,",","")					'remove any comma in attendance		
				
		if IsNumeric(attend) then
		 
			sql = "update FL_results set attendance = " & attend & " where date = '" & request.form("date" & i) & "' and home_team = '" & request.form("home" & i)  & "';"
			on error resume next
			conn.Execute sql
			on error goto 0
			if err = 0 then
				output = output & "<p class=""style1bold"" style=""color:green"">Attendance successfully loaded in 'FL_results': " & request.form("date" & i) & " - " & request.form("home" & i) & " - " & attend & "</p>" 
		 		logvalues = "'" & administrator & "','R','FL_results','attendance',NULL,'" & attend & "'"
			 	Call AddToLog
			  else
				Response.Write("<p class=""style1bold"" style=""color:red"">SQL ERROR!!<br>Statement: " & sql & "<br>Error: " & err.description & "</p>")
			end if
		end if
	next
	
	output = output & "<p class=""style4"" class=""width:400px, text-align:center,"">Attendance(s) in FL_results have been updated.</p>"

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