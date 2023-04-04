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

<% 
Dim output, phase, administrator
Dim conn,sql,rs

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_admin.inc"--><%
%>

<div id="container">
<!--#include file="admin_head.inc"-->

<% 

phase = request.form("phase")

select case phase
	case 1
		Call Refresh
	case else
		Call Input
end select

Response.write(output)

Sub Input

	output = "<form action=""milestone_refresh.asp"" method=""post"">"
	output = output & "<input type=""hidden"" name=""phase"" value=""1"">"			
	output = output & "<input style=""margin:10px 0 0;"" type=""submit"" value=""Confirm Milestone Refresh"">"
	output = output & "</form>"

End Sub

Sub Refresh

	Set conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	%><!--#include file="conn_admin.inc"--><%
	
	sql = "select count(*) as rows "
	sql = sql & "from match_milestone "
	
	rs.open sql,conn,1,2
	output = "<p class=""style1bold"">Milestones updated</p>" 
	output = "<p class=""style1"">" & rs.Fields("rows") & " before</p>"
	rs.close
	
	 
	' --- club matches divisible by 100 
	
	sql = "delete from match_milestone where type ='CM' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(order by date) as quantity, date "
	sql = sql & "from match "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity) "
	sql = sql & "select date, 'CM', 10, quantity "
	sql = sql & "from cte "
	sql = sql & "where quantity % 100 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club matches at home divisible by 100 
	
	sql = "delete from match_milestone where type ='CMH' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by date) as quantity, date, homeaway "
	sql = sql & "from match "
	sql = sql & "where homeaway = 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity) "
	sql = sql & "select date, 'CMH', 20, quantity "
	sql = sql & "from cte "
	sql = sql & "where quantity % 100 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club matches away divisible by 100 
	
	sql = "delete from match_milestone where type ='CMA' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by date) as quantity, date, homeaway "
	sql = sql & "from match "
	sql = sql & "where homeaway <> 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity) "
	sql = sql & "select date, 'CMA', 20, quantity "
	sql = sql & "from cte "
	sql = sql & "where quantity % 100 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club matches in the FL divisible by 100 
	
	sql = "delete from match_milestone where type ='CMFL' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(order by date) as quantity, date, LFC, a.compcode "
	sql = sql & "from match a join competition b on a.compcode = b.compcode "
	sql = sql & "where lfc = 'F' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity) "
	sql = sql & "select date, 'CMFL', 30, quantity "
	sql = sql & "from cte  "
	sql = sql & "where quantity % 100 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club matches at home in the FL divisible by 100 
	
	sql = "delete from match_milestone where type ='CMFLH' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by date) as quantity, date, LFC, a.compcode, homeaway "
	sql = sql & "from match a join competition b on a.compcode = b.compcode "
	sql = sql & "where lfc = 'F' "
	sql = sql & "and homeaway = 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity) "
	sql = sql & "select date, 'CMFLH', 40, quantity "
	sql = sql & "from cte  "
	sql = sql & "where quantity % 100 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club matches away in the FL divisible by 100 
	
	sql = "delete from match_milestone where type ='CMFLA' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by date) as quantity, date, LFC, a.compcode, homeaway "
	sql = sql & "from match a join competition b on a.compcode = b.compcode "
	sql = sql & "where lfc = 'F' "
	sql = sql & "and homeaway <> 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity) "
	sql = sql & "select date, 'CMFLA', 40, quantity "
	sql = sql & "from cte  "
	sql = sql & "where quantity % 100 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club goals divisible by 100 
	
	sql = "delete from match_milestone where type ='CG';  "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(order by b.date) as quantity, b.date, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join match_goal b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'CG', 50, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte "
	sql = sql & "where quantity % 100 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club goals at home divisible by 100 
	
	sql = "delete from match_milestone where type ='CGH' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by b.date) as quantity, b.date, homeaway, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join match_goal b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where homeaway = 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'CGH', 60, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte  "
	sql = sql & "where quantity % 100 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club goals away divisible by 100 
	
	sql = "delete from match_milestone where type ='CGA' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by b.date) as quantity, b.date, homeaway, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join match_goal b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & "where homeaway <> 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'CGA', 60, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte  "
	sql = sql & "where quantity % 100 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club goals in the FL divisible by 100 
	
	sql = "delete from match_milestone where type ='CGFL' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(order by c.date) as quantity, c.date, LFC, a.compcode, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join match_goal c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where lfc = 'F' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'CGFL', 70, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte "
	sql = sql & "where quantity % 100 = 0  "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club goals at home in the FL divisible by 100 
	
	sql = "delete from match_milestone where type ='CGFLH' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by c.date) as quantity, c.date, LFC, a.compcode, homeaway, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join match_goal c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where lfc = 'F' "
	sql = sql & "and homeaway = 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'CGFLH', 80, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte "
	sql = sql & "where quantity % 100 = 0  "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club goals away in the FL divisible by 100  
	
	sql = "delete from match_milestone where type ='CGFLA' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by c.date) as quantity, c.date, LFC, a.compcode, homeaway, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join match_goal c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where lfc = 'F' "
	sql = sql & "and homeaway <> 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'CGFLA', 80, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte "
	sql = sql & "where quantity % 100 = 0  "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club goals against opposition divisible by 50 
	
	sql = "delete from match_milestone where type ='CGO';  "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by opposition order by b.date) as quantity, b.date, player_id_spell1, surname, forename, initials, opposition "
	sql = sql & "from match a join match_goal b on a.date = b.date join player c on b.player_id = c.player_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials, opposition) "
	sql = sql & "select date, 'CGO', 85, quantity, player_id_spell1, surname, forename, initials, opposition "
	sql = sql & "from cte "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club matches in the FA Cup divisible by 50  
	
	sql = "delete from match_milestone where type ='CMFAC' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(order by date) as quantity, date, LFC, a.compcode "
	sql = sql & "from match a join competition b on a.compcode = b.compcode "
	sql = sql & "where a.compcode = 'FAC' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity) "
	sql = sql & "select date, 'CMFAC', 70, quantity "
	sql = sql & "from cte  "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club matches at home in the FA Cup divisible by 50  
	
	sql = "delete from match_milestone where type ='CMFACH' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by date) as quantity, date, LFC, a.compcode, homeaway "
	sql = sql & "from match a join competition b on a.compcode = b.compcode  "
	sql = sql & "where a.compcode = 'FAC' "
	sql = sql & "and homeaway = 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity) "
	sql = sql & "select date, 'CMFACH', 80, quantity "
	sql = sql & "from cte  "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club matches away in the FA Cup divisible by 50  
	
	sql = "delete from match_milestone where type ='CMFACA' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by date) as quantity, date, LFC, a.compcode, homeaway "
	sql = sql & "from match a join competition b on a.compcode = b.compcode "
	sql = sql & "where a.compcode = 'FAC' "
	sql = sql & "and homeaway <> 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity) "
	sql = sql & "select date, 'CMFACA', 80, quantity "
	sql = sql & "from cte  "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club goals in the FA Cup divisible by 50  
	
	sql = "delete from match_milestone where type ='CGFAC';  "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(order by c.date) as quantity, c.date, LFC, a.compcode, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join match_goal c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where a.compcode = 'FAC' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'CGFAC', 90, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club goals at home in the FA Cup divisible by 50 
	
	sql = "delete from match_milestone where type ='CGFACH';  "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by c.date) as quantity, c.date, LFC, a.compcode, homeaway, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join match_goal c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where a.compcode = 'FAC' "
	sql = sql & "and homeaway = 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'CGFACH', 100, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- club goals away in the FA Cup divisible by 50 
	
	sql = "delete from match_milestone where type ='CGFACA';  "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by case homeaway when 'H' then 'H' else 'A' end order by c.date) as quantity, c.date, LFC, a.compcode, homeaway, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join match_goal c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where a.compcode = 'FAC' "
	sql = sql & "and homeaway <> 'H' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'CGFACA', 100, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- manager first match 
	
	sql = "delete from match_milestone where type ='M1M' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by manager_id order by date) as quantity, date, manager_id, surname, forename, manager_id2 "
	sql = sql & "from match a join manager_spell b on a.date >= b.from_date and a.date <= isnull(b.to_date,'9999-12-31') join manager c on b.manager_id1 = c.manager_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, manager_id, surname, forename) "
	sql = sql & "select date, 'M1M', 110, quantity, manager_id, surname, forename   "
	sql = sql & "from cte  "
	sql = sql & "where quantity = 1 "
	sql = sql & "and manager_id2 is null "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- manager last match 
	
	sql = "delete from match_milestone where type ='MLM' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by manager_id order by date) as quantity, date, manager_id, surname, forename, manager_id2 "
	sql = sql & "from match a join manager_spell b on a.date >= b.from_date and a.date <= b.to_date join manager c on b.manager_id1 = c.manager_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, manager_id, surname, forename) "
	sql = sql & "select date, 'MLM', 130, quantity, manager_id, surname, forename  "
	sql = sql & "from cte a  "
	sql = sql & "where not exists (select * from cte b where b.manager_id = a.manager_id and b.quantity > a.quantity) "
	sql = sql & "and manager_id2 is null "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- manager matches divisible by 50 
	
	sql = "delete from match_milestone where type ='MM' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by manager_id order by date) as quantity, date, manager_id, surname, forename, manager_id2 "
	sql = sql & "from match a join manager_spell b on a.date >= b.from_date and a.date <= isnull(b.to_date,'9999-12-31') join manager c on b.manager_id1 = c.manager_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, manager_id, surname, forename) "
	sql = sql & "select date, 'MM', 120, quantity, manager_id, surname, forename "
	sql = sql & "from cte "
	sql = sql & "where quantity % 50 = 0 "
	sql = sql & "and manager_id2 is null "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- opposition first match 
	
	sql = "delete from match_milestone where type ='O1M' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by name_now order by date) as quantity, date, name_now "
	sql = sql & "from match a join opposition b on a.opposition = b.name_then "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, opposition) "
	sql = sql & "select date, 'O1M', 140, quantity, name_now "
	sql = sql & "from cte  "
	sql = sql & "where quantity = 1 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- opposition first match in the FL  
	
	sql = "delete from match_milestone where type ='O1MFL' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by name_now order by date) as quantity, date, name_now "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join opposition c on a.opposition = c.name_then "
	sql = sql & "where lfc = 'F' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, opposition) "
	sql = sql & "select date, 'O1MFL', 142, quantity, name_now "
	sql = sql & "from cte  "
	sql = sql & "where quantity = 1 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- opposition matches divisible by 10 
	
	sql = "delete from match_milestone where type ='OM' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by name_now order by date) as quantity, date, name_now "
	sql = sql & "from match a join opposition b on a.opposition = b.name_then "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, opposition) "
	sql = sql & "select date, 'OM', 150, quantity, name_now "
	sql = sql & "from cte "
	sql = sql & "where quantity % 10 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- opposition FL matches divisible by 10 
	
	sql = "delete from match_milestone where type ='OMFL' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by name_now order by date) as quantity, date, name_now "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join opposition c on a.opposition = c.name_then "
	sql = sql & "where lfc = 'F' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, opposition) "
	sql = sql & "select date, 'OMFL', 152, quantity, name_now "
	sql = sql & "from cte "
	sql = sql & "where quantity % 10 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player first match (but not only match, or a current player so we can't yet say it's an only match) 
	
	sql = "delete from match_milestone where type ='P1M' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by date) as quantity, date, player_id_spell1, surname, forename, initials, last_game_year "
	sql = sql & "from match_player a join player b on a.player_id = b.player_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'P1M', 200, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a "
	sql = sql & "where quantity = 1 "
	sql = sql & "and last_game_year < 9999 "
	sql = sql & "and exists (select * from cte b where b.player_id_spell1 = a.player_id_spell1 and quantity > 1) "
	sql = sql & "union "
	sql = sql & "select date, 'P1M', 200, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity = 1 "
	sql = sql & "and last_game_year = 9999 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player first start (but not first match, i.e. was a sub before) 
	
	sql = "delete from match_milestone where type ='P1S' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by date) as quantity, date, startpos, player_id_spell1, surname, forename, initials, last_game_year "
	sql = sql & "from match_player a join player b on a.player_id = b.player_id "
	sql = sql & "where startpos > 0 "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'P1S', 205, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a " 
	sql = sql & "where quantity = 1 " 
	sql = sql & "and last_game_year < 9999 " 
	sql = sql & "and startpos > 0 " 
	sql = sql & "and exists (select * from cte b where b.player_id_spell1 = a.player_id_spell1 and quantity > 1) " 
	sql = sql & "and exists (select * from match_player b join player c on b.player_id = c.player_id where c.player_id_spell1 = a.player_id_spell1 and b.date < a.date and b.startpos = 0) " 
	sql = sql & "union " 
	sql = sql & "select date, 'P1S', 205, quantity, player_id_spell1, surname, forename, initials " 
	sql = sql & "from cte a "  
	sql = sql & "where quantity = 1 " 
	sql = sql & "and last_game_year = 9999 " 
	sql = sql & "and startpos > 0 "
	sql = sql & "and exists (select * from match_player b join player c on b.player_id = c.player_id where c.player_id_spell1 = a.player_id_spell1 and b.date < a.date and b.startpos = 0) "
	 
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player only match 
	
	sql = "delete from match_milestone where type ='POM' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by date) as quantity, date, player_id_spell1, surname, forename, initials, last_game_year "
	sql = sql & "from match_player a join player b on a.player_id = b.player_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'POM', 200, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity = 1 "
	sql = sql & "and last_game_year < 9999 "
	sql = sql & "and not exists (select * from cte b where b.player_id_spell1 = a.player_id_spell1 and quantity = 2) "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player last match 
	
	sql = "delete from match_milestone where type ='PLM' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by date) as quantity, date, player_id_spell1, surname, forename, initials, last_game_year "
	sql = sql & "from match_player a join player b on a.player_id = b.player_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PLM', 220, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where last_game_year < 9999 "
	sql = sql & "and not exists (select * from cte b where b.player_id_spell1 = a.player_id_spell1 and b.quantity > a.quantity) "
	sql = sql & "and exists (select * from cte c where c.player_id_spell1 = a.player_id_spell1 and c.quantity < a.quantity) "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player matches divisible by 50 
	
	sql = "delete from match_milestone where type ='PM' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by date) as quantity, date, player_id_spell1, surname, forename, initials, last_game_year "
	sql = sql & "from match_player a join player b on a.player_id = b.player_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PM', 210, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player matches in the FL divisible by 50 
	
	sql = "delete from match_milestone where type ='PMFL' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by a.date) as quantity, a.date, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join match_player c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where lfc = 'F' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PMFL', 230, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player matches in the FAC divisible by 10 
	
	sql = "delete from match_milestone where type ='PMFAC' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by a.date) as quantity, a.date, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join match_player c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where compcode = 'FAC' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PMFAC', 230, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity % 10 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player starts divisible by 50 
	
	sql = "delete from match_milestone where type ='PS' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by date) as quantity, date, player_id_spell1, surname, forename, initials, last_game_year "
	sql = sql & "from match_player a join player b on a.player_id = b.player_id "
	sql = sql & "where startpos > 0 "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PS', 210, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player starts in the FL divisible by 50 
	
	sql = "delete from match_milestone where type ='PSFL' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by a.date) as quantity, a.date, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join match_player c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where lfc = 'F' "
	sql = sql & "and startpos > 0 "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PSFL', 230, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player starts in the FAC divisible by 10 
	
	sql = "delete from match_milestone where type ='PSFAC' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by a.date) as quantity, a.date, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join match_player c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where compcode = 'FAC' "
	sql = sql & "and startpos > 0 "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PSFAC', 230, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity % 10 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	
	' --- player first goal (but not only goal, or a current player so we can't yet say it's an only goal) 
	
	sql = "delete from match_milestone where type ='P1G' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by date) as quantity, date, player_id_spell1, surname, forename, initials, last_game_year "
	sql = sql & "from match_goal a join player b on a.player_id = b.player_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'P1G', 240, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity = 1 "
	sql = sql & "and last_game_year < 9999 "
	sql = sql & "and exists (select * from cte b where b.player_id_spell1 = a.player_id_spell1 and quantity > 1) "
	sql = sql & "union all "
	sql = sql & "select date, 'P1G', 240, quantity, player_id_spell1, surname, forename, initials   "
	sql = sql & "from cte a  "
	sql = sql & "where quantity = 1 "
	sql = sql & "and last_game_year = 9999 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player only goal 
	
	sql = "delete from match_milestone where type ='POG' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by date) as quantity, date, player_id_spell1, surname, forename, initials, last_game_year "
	sql = sql & "from match_goal a join player b on a.player_id = b.player_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'POG', 240, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity = 1 "
	sql = sql & "and last_game_year < 9999 "
	sql = sql & "and not exists (select * from cte b where b.player_id_spell1 = a.player_id_spell1 and quantity = 2) "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player last goal 
	
	sql = "delete from match_milestone where type ='PLG' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by date) as quantity, date, player_id_spell1, surname, forename, initials, last_game_year "
	sql = sql & "from match_goal a join player b on a.player_id = b.player_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PLG', 260, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where last_game_year < 9999 "
	sql = sql & "and not exists (select * from cte b where b.player_id_spell1 = a.player_id_spell1 and b.quantity > a.quantity) "
	sql = sql & "and exists (select * from cte c where c.player_id_spell1 = a.player_id_spell1 and c.quantity < a.quantity) "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player goals divisible by 50 
	
	sql = "delete from match_milestone where type ='PG' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by date) as quantity, date, player_id_spell1, surname, forename, initials, last_game_year "
	sql = sql & "from match_goal a join player b on a.player_id = b.player_id "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PG', 250, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player goals in the FL divisible by 50 
	
	sql = "delete from match_milestone where type ='PGFL' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by a.date) as quantity, a.date, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join competition b on a.compcode = b.compcode join match_goal c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where lfc = 'F' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PGFL', 270, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity % 50 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	' --- player goals in the FAC divisible by 10 
	
	sql = "delete from match_milestone where type ='PGFAC' "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "with cte as "
	sql = sql & "(select row_number() over(partition by player_id_spell1 order by a.date) as quantity, a.date, player_id_spell1, surname, forename, initials "
	sql = sql & "from match a join match_goal c on a.date = c.date join player d on c.player_id = d.player_id "
	sql = sql & "where compcode = 'FAC' "
	sql = sql & ") "
	sql = sql & "insert into match_milestone (date, type, seq, quantity, player_id_spell1, surname, forename, initials) "
	sql = sql & "select date, 'PGFAC', 270, quantity, player_id_spell1, surname, forename, initials "
	sql = sql & "from cte a  "
	sql = sql & "where quantity % 10 = 0 "
	
	On error resume next
	conn.Execute sql
	if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
	On Error GoTo 0
	
	sql = "select count(*) as rows "
	sql = sql & "from match_milestone "
	
	rs.open sql,conn,1,2
	output = output & "<p class=""style1"" >" & rs.Fields("rows") & " after</p>" 
	rs.close
	
	conn.close
	
	End sub
%>
	
</div>
</body>
</html>