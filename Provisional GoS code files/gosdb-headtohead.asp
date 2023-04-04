<%@ Language=VBScript %>
<% Option Explicit %>

<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Author" content="Trevor Scallan">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<title>Greens on Screen</title><link rel="stylesheet" type="text/css" href="gos2.css">
<style>
<!--
form { margin: 0; }
-->
</style>

<script language="javascript">
function Toggle(item) {
   obj=document.getElementById(item);
   visible=(obj.style.display!="none")
   key=document.getElementById("x" + item);
   if (visible) {
     obj.style.display="none";
     key.innerHTML="[+]";
   } else {
      obj.style.display="block";
      key.innerHTML="[-]";
   }
}
</script>
</head>
<body>
<!--#include file="top_code.htm"-->

<%
Dim conn,sql,rs, n, tableview, outline, totalsave, team, heading1, heading2, headtext, competition, selected_comp(4), homeaway, selected_ha(2)
Dim season_no1, selected_s1, season_no2, selected_s2, season1opts, season2opts, selyears1, selyears2, orderby, selected_or(11), orderby_text, homeaway_text, restrictions, ordered, ordered_warn
Dim i, j, teamline(150)

competition = Request.Form("competition")
homeaway = Request.Form("homeaway")
season_no1 = Request.Form("season1")
season_no2 = Request.Form("season2")
orderby = Request.Form("order") 		

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%
%>

<div id="sv1">
<center>
<table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" cellpadding="0">
 <tr>
 <td width="270" valign="top">
<div style="width:260;">
<p style="text-align: center; margin-top:0; margin-bottom:3">
<a href="gosdb.asp"><font color="#404040"><img border="0" src="images/gosdb-small.jpg" align="left"></font></a><font 
color="#404040"> 
<b><font style="font-size: 15px">Search by<br>
</font></b><span style="font-size: 15px"><b>Opposition</b></span></font><p style="text-align: center; margin-top:0; margin-bottom:0">
<b>
<a href="gosdb.asp">Back to<br>GoS-DB Hub</a></b>
</div>	
</td>
 <td align="center">
 <p style="margin-top: 6; margin-bottom: 9"><font style="font-size: 18px;" color="#006e32">HEAD TO HEAD RESULTS</font></p>
   
 <form style="font-size: 10px; padding: 0; margin: 0;" action="gosdb-headtohead.asp" method="post" name="form1">

<%
 
outline = "<p style=""margin-top: 0px; margin-bottom: 6px;""><font style=""font-size: 12px;"" color=""#202020""><b>" & heading1 & "</b></font></p>"
response.write(outline)
  
restrictions = ""
  
select case competition
	case "FLG"
		selected_comp(1) = "selected"
		tableview = "v_match_FL"
		heading1 = "Football League"
		headtext = "all matches in tier 2 [Div 2 to 1991, Div 1 to 2003 and the Championship]; tier 3 [Div 3 South to 1958, Div 3 to 1991, Div 2 to 2003]; and tier 4 [Div 3, 1992-2003], all from 1920 to 1939 and 1946 to the present day." 
		restrictions = "Y"
	case "LGS"
		selected_comp(2) = "selected"
		tableview = "v_match_all_league"
		heading1 = "All Leagues"
		headtext = "all matches in all league competitions, including the Southern league [1903-20]; the Western League, a mid-week league including south-east clubs [1903-08]; the Football League [1920-39, 1946-present]; the South West Regional League [1939-40]; and the Football League South [1945-46]."
		restrictions = "Y"
	case "FAC"
		selected_comp(3) = "selected"
		tableview = "v_match_FA"
		heading1 = "FA Cup"
		headtext = "all matches in the FA Cup [1903-present]." 		
		restrictions = "Y"
	case "CUP"
		selected_comp(4) = "selected"
		tableview = "v_match_cups"
		heading1 = "All Cups"
		headtext = "all matches in all knock-out competitions, including the FA Cup [1903-present]; the Football League War Cup [1940]; the Football League Cup (various sponsors) [1960-present]; the Full Members Cup [1986] (a competition for tiers 1 and 2, also known as the Simod Cup [1987-88] and Zenith Data Systems Cup [1989-91]); the Football League Trophy (a generic name for a competition for tiers 3 and 4, including the Associate Members Cup [1984], the Freight Rovers Trophy [1985-86], the Autoglass Trophy[1993], the Auto Windshields Shield [1994-2000], the LDV Vans Trophy [2000-03] and the Johnstone's Paint Trophy [from 2010]); and official pre-season competitions (the Watney Cup [1973], the Anglo Scottish Cup [1977-79] and the Football League Group Cup [1981])." 		
		restrictions = "Y"
	case else
		selected_comp(0) = "selected"
		tableview = "match"
		heading1 = "All Competitions"
		headtext = "all competitive first-team games since the club turned professional, including the Southern League [1903-1920]; the Western League, a mid-week league including south-east clubs [1903-08]; the Football League [1920-39, 1946-present]; the South West Regional League [1939-40]; the Football League South [1945-46]; and all Cup competitions (see the Cup option for details)." 
 end select
 heading2 = heading1
 heading1 = heading1 & "<a style=""font-family:courier"" id=""xheadtext"" href=""javascript:Toggle('headtext');"">[+]</a>"
 
 outline = "<select name=""competition"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ALL"" " & selected_comp(0) & ">All Competitions</option>"
 outline = outline & "<option value=""FLG"" " & selected_comp(1) & ">Football League</option>"  
 outline = outline & "<option value=""LGS"" " & selected_comp(2) & ">All Leagues</option>"  
 outline = outline & "<option value=""FAC"" " & selected_comp(3) & ">FA Cup</option>"
 outline = outline & "<option value=""CUP"" " & selected_comp(4) & ">All Cups</option>"    
 outline = outline & "</select>"
 

 select case homeaway
	case "ho" 
		selected_ha(1) = "selected"
		homeaway_text = " and homeaway = 'H'" 
		heading1 = heading1 & ", home only"
		restrictions = "Y"
	case "ao" 
		selected_ha(2) = "selected" 
		homeaway_text = " and homeaway = 'A'"
		heading1 = heading1 & ", away only"
		restrictions = "Y"
 	case else
		selected_ha(0) = "selected"
		homeaway_text = ""
 end select	

 outline = outline & " <select name=""homeaway"" style=""font-size: 10px"">"
 outline = outline & "<option value=""ha"" " & selected_ha(0) & ">Home & Away</option>" 
 outline = outline & "<option value=""ho"" " & selected_ha(1) & ">Home only</option>"
 outline = outline & "<option value=""ao"" " & selected_ha(2) & ">Away only</option>"
 outline = outline & "</select>"
 
 sql = "select season_no, years "
 sql = sql & "from season "
 rs.open sql,conn,1,2
 
 if season_no1 = "" then season_no1 = 1
 if season_no2 = "" then season_no2 = CStr(rs.RecordCount)
 
 if season_no1 > 1 or season_no2 < CStr(rs.RecordCount) then restrictions = "Y"
 
 season1opts = ""
 season2opts = ""
   
 Do While Not rs.EOF
  if CStr(rs.Fields("season_no")) = season_no1 then 
    selected_s1 = "selected"
    if season_no1 > 1 then
    	selyears1 = rs.Fields("years") 
    	heading1 = heading1 & " from " & selyears1
    end if
   else selected_s1 = ""
  end if
  season1opts = season1opts & "<option value=""" & rs.Fields("season_no") & """ " & selected_s1 & ">From " & rs.Fields("years") & "</option>"
  if CStr(rs.Fields("season_no")) = season_no2 then
    selected_s2 = "selected"
    if season_no2 < CStr(rs.RecordCount) then 
    selyears2 = rs.Fields("years") 
    heading1 = heading1 & " to " & selyears2
    end if
   else selected_s2 = ""
  end if
  season2opts = season2opts & "<option value=""" & rs.Fields("season_no") & """ " & selected_s2 & ">To " & rs.Fields("years") & "</option>"
 rs.MoveNext
 Loop
 
 rs.close
 
 outline = outline & " <select name=""season1"" style=""font-size: 10px"">" & season1opts & "</select>"  
 outline = outline & " <select name=""season2"" style=""font-size: 10px"">" & season2opts & "</select>"
 
 select case orderby
	case "P" 
		selected_or(1) = "selected"
		orderby_text = " P DESC, name_now "
		ordered = "Y"
	case "W" 
		selected_or(2) = "selected"
		orderby_text = " W DESC, name_now "
		heading1 = heading1 & ", ordered by wins"
		ordered = "Y"  
	case "D"
		selected_or(3) = "selected"
		orderby_text = " D DESC, name_now "
		heading1 = heading1 & ", ordered by draws"
		ordered = "Y"
	case "L"
		selected_or(4) = "selected"
		orderby_text = " L DESC, name_now "
		heading1 = heading1 & ", ordered by defeats"
		ordered = "Y"
	case "F"
		selected_or(5) = "selected"
		orderby_text = " F DESC, name_now "
		heading1 = heading1 & ", ordered by goals-for"
		ordered = "Y"
	case "A"
		selected_or(6) = "selected"
		orderby_text = " A DESC, name_now "
		heading1 = heading1 & ", ordered by goals-against"
		ordered = "Y"
	case "WP" 
		selected_or(7) = "selected"
		orderby_text = " WP DESC, name_now "
		heading1 = heading1 & ", ordered by wins per game"
		ordered = "Y"  
	case "DP"
		selected_or(8) = "selected"
		orderby_text = " DP DESC, name_now "
		heading1 = heading1 & ", ordered by draws per game"
		ordered = "Y"
	case "LP"
		selected_or(9) = "selected"
		orderby_text = " LP DESC, name_now "
		heading1 = heading1 & ", ordered by defeats per game"
		ordered = "Y"
	case "FP"
		selected_or(10) = "selected"
		orderby_text = " FP DESC, name_now "
		heading1 = heading1 & ", ordered by goals-for per game"
		ordered = "Y"
	case "AP"
		selected_or(11) = "selected"
		orderby_text = " AP DESC, name_now "
		heading1 = heading1 & ", ordered by goals-against per game"
		ordered = "Y"
	case else
		selected_or(0) = "selected"
		orderby_text = " name_now "
		ordered = ""				
 end select	
 
 outline = outline & "<p style=""margin:6 0 0 0""><select name=""order"" style=""font-size: 10px"">"
 outline = outline & "<option value=""O"" " & selected_or(0) & ">Order by Opposition</option>"  
 outline = outline & "<option value=""P"" " & selected_or(1) & ">Order by Played</option>"  
 outline = outline & "<option value=""W"" " & selected_or(2) & ">Order by Wins</option>"
 outline = outline & "<option value=""D"" " & selected_or(3) & ">Order by Draws</option>" 
 outline = outline & "<option value=""L"" " & selected_or(4) & ">Order by Defeats</option>" 
 outline = outline & "<option value=""F"" " & selected_or(5) & ">Order by Goals For</option>" 
 outline = outline & "<option value=""A"" " & selected_or(6) & ">Order by Goals Against</option>"
 outline = outline & "<option value=""WP"" " & selected_or(7) & ">Order by Wins/Game</option>"
 outline = outline & "<option value=""DP"" " & selected_or(8) & ">Order by Draws/Game</option>" 
 outline = outline & "<option value=""LP"" " & selected_or(9) & ">Order by Losses/Game</option>" 
 outline = outline & "<option value=""FP"" " & selected_or(10) & ">Order by For/Game</option>" 
 outline = outline & "<option value=""AP"" " & selected_or(11) & ">Order by Against/Game</option>"       
 outline = outline & "</select></p>"
 outline = outline & "<input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 9 0 0 0"" value=""Select options above and click here to redisplay"" name=""B1""></p>" 
 response.write(outline)
 %>

 </form>

  
 </td>
 <td width="270" valign="top">
 <p style="margin-bottom:4pt; margin-right:3; margin-left:15" align="justify">Our record against 
 every team ever played. Select options to change the scope of the search. 
 Select a team for full match details.</td>
 </tr>
 <tr><td colspan="3" align="center">
 
 <%
  outline = ""
  if ordered = "Y" then
  	ordered_warn = ", and the display has been re-ordered"
  	else ordered_warn = ""
  end if			
  if restrictions = "Y" then 
  	outline = outline & "<p style=""margin-top: 0px; margin-bottom: 6px;""><font style=""font-size: 11px;"" color=""#900033""><b>Reminder: options have limited the results " & ordered_warn & "</b></font><a href=""http://www.greensonscreen.co.uk/gosdb-headtohead.asp""> [reset to initial state]</a></p>"
  end if
  if restrictions <> "Y" and ordered = "Y" then 
  	outline = outline & "<p style=""margin-top: 0px; margin-bottom: 6px;""><font style=""font-size: 11px;"" color=""#900033""><b>Reminder: the display has been re-ordered</b></font><a href=""http://www.greensonscreen.co.uk/gosdb-headtohead.asp""> [reset to initial state]</a></p>"
  end if

  outline = outline & "<div id=""headtext"" style=""width:900px; display:none""><p style=""margin:0 0 9 0; text-align: justify"">"
  outline = outline & "<b>" & heading2 & ":</b> " & headtext
  outline = outline & "</p></div>"
  response.write(outline)

sql = "select name_now, sum(p) as P, sum(w) as W, sum(d) as D, sum(l) as L, sum(f) as F, sum(a) as A, "
sql = sql & " 1000*sum(w)/sum(p) as WP, 1000*sum(d)/sum(p) as DP, 1000*sum(l)/sum(p) as LP, 1000*sum(f)/sum(p) as FP, 1000*sum(a)/sum(p) as AP "
sql = sql & "from ( "
sql = sql & "select name_now, 1 as p, "
sql = sql & "case when goalsfor > goalsagainst then 1 else 0 end as w, "
sql = sql & "case when goalsfor = goalsagainst then 1 else 0 end as d, "
sql = sql & "case when goalsfor < goalsagainst then 1 else 0 end as l, "
sql = sql & "goalsfor as f, goalsagainst as a  "
sql = sql & "from " & tableview & " join opposition on opposition = name_then join season on date between date_start and date_end "
sql = sql & "where season_no between " & season_no1 & " and " & season_no2 & homeaway_text
sql = sql & ") as subsel "
sql = sql & "group by name_now with rollup "
sql = sql & "order by" & orderby_text

rs.open sql,conn,1,2

i = 1
j = i

Do While Not rs.EOF

	if IsNull(rs.Fields("name_now")) then
		' the total line - position it at the top
		j = 0		' force these counts to the top
		i = i - 1   ' so we don't miss out an array position
		teamline(j)  = "<tr><td><b>Total</b></td>"
	  else
		team = Replace(rs.Fields("name_now")," ","%20")
		team = Replace(rs.Fields("name_now"),"&","%26")  
		teamline(j)  = teamline(j) & "<tr><td nowrap=""nowrap""><a href=""gosdb-results.asp?team=" & team & "&comp=" & competition & "&s1=" & selyears1 & "&s2=" & selyears2 &""">"
		teamline(j)  = teamline(j) & "<u>" & rs.Fields("name_now") & "</u></a></td>"
	end if
	teamline(j)  = teamline(j) & "<td>" & rs.Fields("p") & "</td>"
	teamline(j)  = teamline(j) & "<td>" & rs.Fields("w") & "</td>"  
	teamline(j)  = teamline(j) & "<td>" & rs.Fields("d") & "</td>"  
	teamline(j)  = teamline(j) & "<td>" & rs.Fields("l") & "</td>"  
	teamline(j)  = teamline(j) & "<td>" & rs.Fields("f") & "</td>"  
	teamline(j)  = teamline(j) & "<td>" & rs.Fields("a") & "</td></tr>" 
	
	i = i + 1
	j = i   

	rs.MoveNext
	
Loop

%> 
 
 </td></tr>
 </table>	

<table style="border-collapse: collapse;" border="0" bordercolor="#111111" cellpadding="0" cellspacing="0" 

<% 	
if rs.RecordCount > 10 then
	response.write(" width=""900""")
   else response.write(" width=""440""")
end if
%>

><tbody>
<tr>

<%
i = 0
outline = ""

n = 0

Do Until teamline(i) = ""

	if n = 0 then
		outline  = outline & "<td align=""center"" valign=""top"">"
		outline  = outline & "<table style=""border-collapse: collapse;"" border=""1"" bordercolor=""#c0c0c0"" cellpadding=""0"" cellspacing=""0"" cols=""11"">"
		outline  = outline & "<tbody>"
		outline  = outline & "<tr>"
		outline  = outline & "<td><b>Opposition</b></td>"
		outline  = outline & "<td width=""30""><b>P</b></td>"
		outline  = outline & "<td width=""30""><b>W</b></td>"
		outline  = outline & "<td width=""30""><b>D</b></td>"
		outline  = outline & "<td width=""30""><b>L</b></td>"
		outline  = outline & "<td width=""30""><b>F</b></td>"
		outline  = outline & "<td width=""30""><b>A</b></td>"
		outline  = outline & "</tr>"
	end if

	outline  = outline & teamline(i)
	i = i + 1 
	
	n = n + 1
	if rs.RecordCount > 10 and n = Int((rs.RecordCount/2)+1) then
		outline  = outline & "</tbody>"
		outline  = outline & "</table>"
		outline  = outline & "</td>" 
		n = 0
	end if 

Loop 

'Finish with total line repeated at the bottom
outline  = outline & teamline(0) & "</tbody></table></td>"
response.write(outline)
	
rs.close
conn.close		

%>

</tr>
<tr>
<td colspan="2" align="center" valign="top">
<p style="margin: 12px 12px 0pt;">
Where a team has changed its name, all results are shown under its latest name, 
but contemporary names are used on the results page for each club.
</p>
</td>
</tr>

</tbody>
</table>
</center>
<br>
</div>

<!--#include file="base_code.htm"-->
</body></html>