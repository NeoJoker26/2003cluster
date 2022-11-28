<%@ Language=VBScript %>
<% Option Explicit %>

<% 
'Set Session Locale to English(UK) to ensure correct display of date for FormatDateTime(xxx,1)
Session.LCID=2057 

Dim conn, sql, rs, input, inputs, goalsfor, goalsagainst, output, output1, season_no1, season_no2, LFCvalue, HAvalue, opposition, crowd, displaydate, work1, venue
  
input = Request.QueryString("input")
inputs = split(input,"-")
goalsfor = inputs(0)
goalsagainst = inputs(1)
LFCvalue = inputs(2)
HAvalue = inputs(3)
season_no1 = inputs(4)
season_no2 = inputs(5)

select case LFCvalue
	case "F"
		LFCvalue = "'F'"
	case "C"
		LFCvalue = "'C'"
	case else
		LFCvalue = "'L','F','C'"
 end select
 
 select case HAvalue
	case "H"
		HAvalue = "'H'"
	case "A"
		HAvalue = "'A'"
	case else
		HAvalue = "'H','A'"
 end select

output = "<p class=""style1bold"" style=""margin: 36px auto 0"">Click on a date for the full match page (in a new tab)</p>"
output = output & "<table id=""table2"" style=""border-collapse: collapse; cell-spacing: 0; margin-top: 12px"">"

output1 = ""

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

sql = "select top 50 opposition, opposition_qual, date, "
sql = sql & "case homeaway when 'H' then 'H' else 'A' end as homeawayHA"
sql = sql & ", shortcomp, subcomp, goalsfor, goalsagainst, attendance, notes, ground_name "
sql = sql & "from v_match_all " 
sql = sql & " join season on date between date_start and date_end "
sql = sql & " join opposition on opposition = name_then "
sql = sql & " left outer join venue on name_then = club_name_then and date between first_game and last_game "
sql = sql & "where goalsfor = " & goalsfor & " "
sql = sql & "  and goalsagainst = " & goalsagainst & " "
sql = sql & "  and season_no between '" & season_no1 & "' and '" & season_no2 & "' "
sql = sql & "  and LFC in (" & LFCvalue & ") "
sql = sql & "  and homeaway in (" & HAvalue & ") "
sql = sql & "order by date desc"

rs.open sql,conn,1,2

Do While Not rs.EOF

	opposition = rs.Fields("opposition") 
	opposition = replace(opposition,"United","Utd")
	opposition = replace(opposition,"Rovers","Rvrs")
	opposition = replace(opposition,"Wanderers","Wnds")
	opposition = replace(opposition,"Albion","Alb")
	opposition = replace(opposition,"Athletic","Ath")
	opposition = replace(opposition,"County","Co")
	opposition = replace(opposition,"Sheffield Wednesday","Sheffield Wed")
	opposition = replace(opposition,"Avenue","Ave")
	opposition = replace(opposition," and "," & ")
	
	if IsNumeric(rs.Fields("attendance")) then 
		crowd = FormatNumber(rs.Fields("attendance"),0,0,0,-1)
		else
		crowd = "Unknown"
	end if
	
	displaydate = FormatDateTime(rs.Fields("date"),1)
	work1 = split(displaydate," ")
	displaydate = work1(0) & " " & left(work1(1),3) & " " & work1(2) 

	venue = rs.Fields("ground_name")
	
	output1  = output1 & "<tr>"
	output1  = output1 & "<td class=""matchdate nowrap""><a href=""gosdb-match.asp?date=" & rs.Fields("date") & """ target=""_blank"">" & displaydate & "</a></td>"	
	if rs.Fields("homeawayHA") = "H" then
		venue = "Home Park"
		if displaydate = "18 Mar 1961" then venue = "Plainmoor"
		output1  = output1 & "<td class=""right"" align=""right"">" & "Argyle" & "</td>"   
		output1  = output1 & "<td class=""nowrap noverticalborder"">" & rs.Fields("goalsfor") & " - " & rs.Fields("goalsagainst") & "</td>" 
		output1  = output1 & "<td>" & opposition & " " & rs.Fields("opposition_qual") & "</td>"
		output1  = output1 & "<td class=""nowrap"">" & rs.Fields("shortcomp") & " " & rs.Fields("subcomp") & "</td>"
		output1  = output1 & "<td class=""right"">" & crowd & "</td>" 
		output1  = output1 & "<td>" & venue & "</td>"   
		output1  = output1 & "</tr>"
	  else  
		output1  = output1 & "<td class=""right"">" & opposition & " " & rs.Fields("opposition_qual") & "</td>"  
		output1  = output1 & "<td class=""nowrap"">" & rs.Fields("goalsagainst") & " - " & rs.Fields("goalsfor") & "</td>"
		output1  = output1 & "<td>" & "Argyle" & "</td>"	
		output1  = output1 & "<td class=""nowrap"">" & rs.Fields("shortcomp") & " " & rs.Fields("subcomp") & "</td>"
		output1  = output1 & "<td class=""right"">" & crowd & "</td>" 
		output1  = output1 & "<td>" & venue & "</td>" 
		output1  = output1 & "</tr>"		
	end if
	
	rs.MoveNext 
Loop 

rs.close

conn.close

if output1 > "" then response.write(output & output1 & "</table>")	

%>