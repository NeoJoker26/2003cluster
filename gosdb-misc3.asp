<%@ Language=VBScript %> 
<% Option Explicit %>
<% dim scope
scope = Request.Form("scope")
if scope = "" then scope = Request.Querystring("scope")	'try for a url parameter
scope = replace(scope," ","")
%>
<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Author" content="Trevor Scallan">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<title>GoS-DB Miscellaneous Report</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
#table1 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1}
#table2 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1}
#table3 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; white-space:nowrap; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1}
.med{font-size:medium;font-weight:normal;padding:0;margin:0}#res{padding-right:1em;margin:0 16px}ol li{list-style:none}.g{margin:1em 0}li.g{font-size:small;font-family:arial,sans-serif}.s{max-width:42em}-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline, heading, season_no1, selected_s1, season_no2, selected_s2, season1opts, season2opts, selyears1, selyears2, lastyears
Dim selected_comp(3), competition, LFCvalue, result, restrictions


Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%
%>
  <center>
  <table border="0" cellspacing="0" style="border-collapse: collapse" 
  cellpadding="0" width="980">
    <tr>
    <td width="260" valign="top" style="text-align:center;">

	<p style="text-align: center; margin-top:0; margin-bottom:3">
	<a href="gosdb.asp"><font color="#404040"><img border="0" src="images/gosdb-small.jpg" align="left"></font></a><font color="#404040"> 
	<b><font style="font-size: 15px">Search by<br>
	</font></b><span style="font-size: 15px"><b>Player</b></span></font><p style="text-align: center; margin-top:0; margin-bottom:0">
	<b>
	<a href="gosdb.asp">Back to<br>GoS-DB Hub</a></b></p>

	</td>
    
  	<td width="460" align="center" style="text-align: center" valign="top">	
	<p style="margin-top:12; margin-bottom:0; text-align:center; font-size:18px; color:#006E32">
    MISCELLANEOUS REPORTS</p>  
    
	<p style="margin-top:6; margin-bottom:0; text-align:center; font-size:13px">
    <b>Report 3:  Attendance Highs and Lows</b></p>  
    </td>
        
	<td width="260" valign="top"  align="justify">
	'<span style="font-size: 10px">Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
     
    </td>
    </tr>   
	</table>
	<center>
		
	<form style="font-size: 10px; padding: 0; margin: 0;" action="gosdb-misc3.asp" method="post" name="form1">
	
      
<%
competition = Request.Form("competition")
season_no1 = Request.Form("season1")
season_no2 = Request.Form("season2")

select case competition
	case "FLG"
		selected_comp(1) = "selected"
		LFCvalue = "'F'"
		heading = "Football League"
	case "CUP"
		selected_comp(2) = "selected"
		LFCvalue = "'C'"
		heading = "Cup Competitions"
	case else
		selected_comp(0) = "selected"
		LFCvalue = "'L','F','C'"
		heading = "All Competitions"
 end select
 
 outline = "<select name=""competition"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ALL"" " & selected_comp(0) & ">All Competitions</option>"
 outline = outline & "<option value=""FLG"" " & selected_comp(1) & ">Football League</option>" 
  outline = outline & "<option value=""CUP"" " & selected_comp(2) & ">All Cups</option>" 
 outline = outline & "</select>"

 sql = "select season_no, years "
 sql = sql & "from season "
 rs.open sql,conn,1,2
 
 if season_no1 = "" then season_no1 = 1
 if season_no2 = "" then season_no2 = CStr(rs.RecordCount)
 
 season1opts = ""
 season2opts = ""
   
 Do While Not rs.EOF
  if CStr(rs.Fields("season_no")) = season_no1 then 
    selected_s1 = "selected"
    if season_no1 > 1 then
    	selyears1 = rs.Fields("years") 
    end if
   else selected_s1 = ""
  end if
  season1opts = season1opts & "<option value=""" & rs.Fields("season_no") & """ " & selected_s1 & ">From " & rs.Fields("years") & "</option>"
  if CStr(rs.Fields("season_no")) = season_no2 then
    selected_s2 = "selected"
    if season_no2 < CStr(rs.RecordCount) then 
    selyears2 = rs.Fields("years") 
    end if
   else selected_s2 = ""
  end if
  season2opts = season2opts & "<option value=""" & rs.Fields("season_no") & """ " & selected_s2 & ">To " & rs.Fields("years") & "</option>"
  lastyears = rs.Fields("years")
 rs.MoveNext
 Loop
 
 if season_no1 > 1 and season_no2 < rs.RecordCount then heading = heading & " from " & selyears1 & " to " & selyears2
 if season_no1 = 1 and season_no2 < rs.RecordCount then heading = heading & " from 1903-1904 to " & selyears2
 if season_no1 > 1 and season_no2 = rs.RecordCount then heading = heading & " from " & selyears1 & " to " & lastyears
 
 rs.close
 
 outline = outline & " <select name=""season1"" style=""font-size: 10px"">" & season1opts & "</select>"  
 outline = outline & " <select name=""season2"" style=""font-size: 10px"">" & season2opts & "</select>"
  outline = outline & "<br><input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 9 0 0 0"" value=""Select options and click to redisplay"" name=""B1""></p>"  
 outline = outline & "</form>"
 
 outline = outline & "<p style=""margin-top:6; margin-bottom:0; text-align:center; font-size:13px; color:#006E32""><b>" & heading & "</b></p>"
 
 	sql = "select count(attendance) as attcount, count(*) as matchcount "
	sql = sql & "from v_match_all join season on date between date_start and date_end "
	sql = sql & "where season_no between " & season_no1 & " and " & season_no2
	sql = sql & "  and LFC in (" & LFCvalue & ") "
	rs.open sql,conn,1,2
	
	outline = outline & "<p style=""margin-top:3; margin-bottom:9; text-align:center;"">Note: attendance figures for " & rs.Fields("matchcount") - rs.Fields("attcount") & " of the " & rs.Fields("matchcount") & " matches in<br>the selected category and date range are unrecorded</p>"
 
  	rs.close
 	
 	response.write(outline)
 %>
 
 	<table border="0" cellpadding="10" cellspacing="10" style="border-collapse: collapse; margin:0 0 12 0" bordercolor="#111111">
  	<tr>
    <td valign="top">
      
	<%
	outline = ""
	sql = "select top 50 rank() over (order by attendance desc) as rank, date, attendance, opposition, shortcomp, subcomp, " 
	sql = sql & " goalsfor, goalsagainst "
	sql = sql & "from v_match_all join season on date between date_start and date_end "
	sql = sql & "where season_no between " & season_no1 & " and " & season_no2
	sql = sql & "  and LFC in (" & LFCvalue & ") "
	sql = sql & "  and homeaway = 'H' "
	sql = sql & "  and attendance is not null "
	rs.open sql,conn,1,2
	
	outline = outline & "<table id=""table1"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"">"
 	outline  = outline & "<tr>"
      outline  = outline & "<td colspan=""4""><p style=""margin: 3 0 3 0""><b>Highest Attendances at Home Park*</b></p></td>"
	  outline  = outline & "</tr>"

	Do While Not rs.EOF
		
		if rs.Fields("goalsfor") > rs.Fields("goalsagainst") then 
			result = "W"
	  	  elseif rs.Fields("goalsfor") = rs.Fields("goalsagainst") then
	  		result = "D"
	  	  else
	    	result = "L"
		end if
		
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td>" & rs.Fields("attendance") & "</td>" 
		outline  = outline & "<td>" & FormatDateTime(rs.Fields("date"),1) & "</td>"
		outline  = outline & "<td>" & rs.Fields("shortcomp") & " " & rs.Fields("subcomp") & "</td>"
		outline  = outline & "</tr><tr>"
		outline  = outline & "<td colspan=""2""></td><td colspan=""2"">" 
		outline  = outline & result & " " & rs.Fields("goalsfor") & "-" & rs.Fields("goalsagainst") & " v " & rs.Fields("opposition") & "</td>"
		outline  = outline & "</tr>"   
  		rs.MoveNext
	Loop
		
	rs.close
	outline  = outline & "</table>"
	response.write(outline)
	%>
    <p style="margin-top: 6; margin-bottom: 6">*Note: GoS-DB confines itself to 
    competitive<br>games, but let's not forget the 37639 who saw<br>the 3-2 win against 
    Santos on March 14th, 1973.</p></td>
	<td valign="top">    
	<%
	outline = ""
	sql = "select top 50 rank() over (order by attendance desc) as rank, date, attendance, opposition, shortcomp, subcomp, " 
	sql = sql & " goalsfor, goalsagainst "
	sql = sql & "from v_match_all join season on date between date_start and date_end "
	sql = sql & "where season_no between " & season_no1 & " and " & season_no2
	sql = sql & "  and LFC in (" & LFCvalue & ") "
	sql = sql & "  and homeaway <> 'H' "
	sql = sql & "  and attendance is not null "
	rs.open sql,conn,1,2  
	
	outline = outline & "<table id=""table2"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"">"
 	outline  = outline & "<tr>"
      outline  = outline & "<td colspan=""4""><p style=""margin: 3 0 3 0""><b>Highest Attendances On The Road</b></p></td>"
	  outline  = outline & "</tr>"

	Do While Not rs.EOF
		
		if rs.Fields("goalsfor") > rs.Fields("goalsagainst") then 
			result = "W"
	  	  elseif rs.Fields("goalsfor") = rs.Fields("goalsagainst") then
	  		result = "D"
	  	  else
	    	result = "L"
		end if
		
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td>" & rs.Fields("attendance") & "</td>" 
		outline  = outline & "<td>" & FormatDateTime(rs.Fields("date"),1) & "</td>"
		outline  = outline & "<td>" & rs.Fields("shortcomp") & " " & rs.Fields("subcomp") & "</td>"
		outline  = outline & "</tr><tr>"
		outline  = outline & "<td colspan=""2""></td><td colspan=""2"">" 
		outline  = outline & result & " " & rs.Fields("goalsfor") & "-" & rs.Fields("goalsagainst") & " v " & rs.Fields("opposition") & "</td>"
		outline  = outline & "</tr>"
  		rs.MoveNext
	Loop
		
	rs.close
	outline  = outline & "</table>"
	response.write(outline)
	%>
	</td>
	<td valign="top">
	<%
	outline = ""
	sql = "select top 50 rank() over (order by attendance) as rank, date, attendance, opposition, shortcomp, subcomp, "
	sql = sql & " goalsfor, goalsagainst " 
	sql = sql & "from v_match_all join season on date between date_start and date_end "
	sql = sql & "where season_no between " & season_no1 & " and " & season_no2
	sql = sql & "  and LFC in (" & LFCvalue & ") "
	sql = sql & "  and homeaway = 'H' "
	sql = sql & "  and attendance is not null "
	rs.open sql,conn,1,2
	
	outline = outline & "<table id=""table3"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"">"
 	outline  = outline & "<tr>"
      outline  = outline & "<td colspan=""4""><p style=""margin: 3 0 3 0""><b>Lowest Attendances at Home Park</b></p></td>"
	  outline  = outline & "</tr>"

	Do While Not rs.EOF
			
		if rs.Fields("goalsfor") > rs.Fields("goalsagainst") then 
			result = "W"
	  	  elseif rs.Fields("goalsfor") = rs.Fields("goalsagainst") then
	  		result = "D"
	  	  else
	    	result = "L"
		end if
		
		outline  = outline & "<tr>"
		outline  = outline & "<td>" & rs.Fields("rank") & "</td>"
		outline  = outline & "<td>" & rs.Fields("attendance") & "</td>" 
		outline  = outline & "<td>" & FormatDateTime(rs.Fields("date"),1) & "</td>"
		outline  = outline & "<td>" & rs.Fields("shortcomp") & " " & rs.Fields("subcomp") & "</td>"
		outline  = outline & "</tr><tr>"
		outline  = outline & "<td colspan=""2""></td><td colspan=""2"">" 
		outline  = outline & result & " " & rs.Fields("goalsfor") & "-" & rs.Fields("goalsagainst") & " v " & rs.Fields("opposition") & "</td>"
		outline  = outline & "</tr>" 
  		rs.MoveNext
	Loop
		
	rs.close
	outline  = outline & "</table>"
	response.write(outline)
	%>
	</td></tr>

<%
conn.close
%>	
	
</table>
</center><br>

<!--#include file="base_code.htm"-->
</body>

</html>