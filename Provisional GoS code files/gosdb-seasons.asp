<%@ Language=VBScript %> <% Option Explicit %>
<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Author" content="Trevor Scallan">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<title>GoS-DB All Seasons</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
.rowhlt { background-color: #f8f8f8; }
.hover { background-color: #c8e0c7; }

.seasons {
	border-collapse: collapse; 
	margin: 12px 0;
}
.seasons th, .seasons td {
	border: 1px solid #c0c0c0;
	padding: 2px 3px 2px 3px; 
}
-->
</style>

<script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script>
<script>
$(document).ready(function(){

	$(".season").hover(function () {
    	$(this).toggleClass("hover");
	});
	
    $('.season').on('click',function() {
        $(this).append('<img style="position:absolute; left:3px; border:0;" src="images/ajax-loader.gif">');
    });
    
});
</script>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, n, outline, totalsave, team, heading1, division, att, homeatt, awayatt 
Dim season_no1, selected_s1, season_no2, selected_s2, season1opts, season2opts, selyears1, selyears2, lastyears, orderby, selected_or(40), orderby_text, homeaway_text, restrictions, ordered
Dim i, j, teamline(150)

season_no1 = Request.Form("season1")
season_no2 = Request.Form("season2")
orderby = Request.Form("order") 		

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%
%>
<div style="margin:0px auto; width:980px">
  <table>
    <tr>
      <td width="260" valign="top" style="text-align: left">
		<div style="width:260;">
		<p style="text-align: center; margin-top:0; margin-bottom:3">
		<a href="gosdb.asp"><font color="#404040"><img border="0" src="images/gosdb-small.jpg" align="left"></font></a><font 
		color="#404040"> 
		<b><font style="font-size: 15px">Search by<br>
		</font></b><span style="font-size: 15px"><b>Season</b></span></font><p style="text-align: center; margin-top:0; margin-bottom:0">
		<b>
		<a href="gosdb.asp">Back to<br>GoS-DB Hub</a> </b>
		</div>
      </td>
      <td align="center" valign="top" style="text-align: center" width="460">
      <p style="margin-top: 9; margin-bottom: 18">
      <span style="font-size: 18px"><font color="#006E32">
      SEASONS' RESULTS</font></p>
      <form style="font-size: 10px; padding: 0; margin: 0;" 
      action="gosdb-seasons.asp" method="post" name="form1">

<%
 
 restrictions = ""
 
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
 
 if season_no1 > 1 and season_no2 < CStr(rs.RecordCount) then heading1 = heading1 & "From " & selyears1 & " to " & selyears2
 if season_no1 = 1 and season_no2 < CStr(rs.RecordCount) then heading1 = heading1 & "From 1903-1904 to " & selyears2
 if season_no1 > 1 and season_no2 = CStr(rs.RecordCount) then heading1 = heading1 & "From " & selyears1 & " to " & lastyears
 
 rs.close
 
 outline = outline & " <select name=""season1"" style=""font-size: 10px"">" & season1opts & "</select>"  
 outline = outline & " <select name=""season2"" style=""font-size: 10px"">" & season2opts & "</select>"
 
 select case orderby
	case "S" 
		selected_or(0) = "selected"
		orderby_text = " years "
		ordered = ""
	case "S-" 
		selected_or(1) = "selected"
		orderby_text = " years DESC "
		heading1 = heading1 & ", ordered by seasons, descending"
		ordered = "Y"
	case "P" 
		selected_or(2) = "selected"
		orderby_text = " P DESC, years "
		heading1 = heading1 & ", ordered by games played"
		ordered = "Y" 
	case "W" 
		selected_or(3) = "selected"
		orderby_text = " W DESC, years "
		heading1 = heading1 & ", ordered by wins"
		ordered = "Y"  
	case "D"
		selected_or(4) = "selected"
		orderby_text = " D DESC, years "
		heading1 = heading1 & ", ordered by draws"
		ordered = "Y"
	case "L"
		selected_or(5) = "selected"
		orderby_text = " L DESC, years "
		heading1 = heading1 & ", ordered by defeats"
		ordered = "Y"
	case "F"
		selected_or(7) = "selected"
		orderby_text = " F DESC, years "
		heading1 = heading1 & ", ordered by goals-for"
		ordered = "Y"
	case "A"
		selected_or(7) = "selected"
		orderby_text = " A DESC, years "
		heading1 = heading1 & ", ordered by goals-against"
		ordered = "Y"
	case "PO" 
		selected_or(8) = "selected"
		orderby_text = " PO DESC, years "
		heading1 = heading1 & ", ordered by points (3 for a win for all seasons)"
		ordered = "Y" 
		 
	case "HP" 
		selected_or(9) = "selected"
		orderby_text = " HP DESC, years "
		heading1 = heading1 & ", ordered by home games played"
		ordered = "Y" 
	case "HW" 
		selected_or(10) = "selected"
		orderby_text = " HW DESC, years "
		heading1 = heading1 & ", ordered by home wins"
		ordered = "Y"  
	case "HD"
		selected_or(11) = "selected"
		orderby_text = " HD DESC, years "
		heading1 = heading1 & ", ordered by home draws"
		ordered = "Y"
	case "HL"
		selected_or(12) = "selected"
		orderby_text = " HL DESC, years "
		heading1 = heading1 & ", ordered by home defeats"
		ordered = "Y"
	case "HF"
		selected_or(13) = "selected"
		orderby_text = " HF DESC, years "
		heading1 = heading1 & ", ordered by home goals-for"
		ordered = "Y"
	case "HA"
		selected_or(14) = "selected"
		orderby_text = " HA DESC, years "
		heading1 = heading1 & ", ordered by home goals-against"
		ordered = "Y"
	case "HPO" 
		selected_or(15) = "selected"
		orderby_text = " HPO DESC, years "
		heading1 = heading1 & ", ordered by home points (3 for a win for all seasons)"
		ordered = "Y"  

	case "AP" 
		selected_or(16) = "selected"
		orderby_text = " AP DESC, years "
		heading1 = heading1 & ", ordered by away games played"
		ordered = "Y" 
	case "AW" 
		selected_or(17) = "selected"
		orderby_text = " AW DESC, years "
		heading1 = heading1 & ", ordered by away wins"
		ordered = "Y"  
	case "AD"
		selected_or(18) = "selected"
		orderby_text = " AD DESC, years "
		heading1 = heading1 & ", ordered by away draws"
		ordered = "Y"
	case "AL"
		selected_or(19) = "selected"
		orderby_text = " AL DESC, years "
		heading1 = heading1 & ", ordered by away defeats"
		ordered = "Y"
	case "AF"
		selected_or(20) = "selected"
		orderby_text = " AF DESC, years "
		heading1 = heading1 & ", ordered by away goals-for"
		ordered = "Y"
	case "AA"
		selected_or(21) = "selected"
		orderby_text = " AA DESC, years "
		heading1 = heading1 & ", ordered by away goals-against"
		ordered = "Y"
	case "APO" 
		selected_or(22) = "selected"
		orderby_text = " APO DESC, years "
		heading1 = heading1 & ", ordered by away points (3 for a win for all seasons)"
		ordered = "Y"  

	case "AT" 
		selected_or(23) = "selected"
		orderby_text = " AT desc, years "
		heading1 = heading1 & ", ordered by attendance"
		ordered = "Y" 
	case "HAT" 
		selected_or(24) = "selected"
		orderby_text = " HAT desc, years "
		heading1 = heading1 & ", ordered by home attendance"
		ordered = "Y" 
	case "AAT" 
		selected_or(25) = "selected"
		orderby_text = " AAT desc, years "
		heading1 = heading1 & ", ordered by away attendance"
		ordered = "Y" 

	case "CP" 
		selected_or(26) = "selected"
		orderby_text = " CP DESC, years "
		heading1 = heading1 & ", ordered by cup games played"
		ordered = "Y" 
	case "CW" 
		selected_or(27) = "selected"
		orderby_text = " CW DESC, years "
		heading1 = heading1 & ", ordered by cup wins"
		ordered = "Y"  
	case "CD"
		selected_or(28) = "selected"
		orderby_text = " CD DESC, years "
		heading1 = heading1 & ", ordered by cup draws"
		ordered = "Y"
	case "CL"
		selected_or(29) = "selected"
		orderby_text = " CL DESC, years "
		heading1 = heading1 & ", ordered by cup defeats"
		ordered = "Y"
	case "CF"
		selected_or(30) = "selected"
		orderby_text = " CF DESC, years "
		heading1 = heading1 & ", ordered by cup goals-for"
		ordered = "Y"
	case "CA"
		selected_or(31) = "selected"
		orderby_text = " CA DESC, years "
		heading1 = heading1 & ", ordered by cup goals-against"
		ordered = "Y"
	case "U"
		selected_or(32) = "selected"
		orderby_text = " player_count DESC, years "
		heading1 = heading1 & ", ordered by players used"
		ordered = "Y"
	case "E"
		selected_or(33) = "selected"
		orderby_text = " flendpos, years "
		heading1 = heading1 & ", ordered by final position in the Football League"
		ordered = "Y"
	case "CS"
		selected_or(34) = "selected"
		orderby_text = " clean_sheets DESC, years "
		heading1 = heading1 & ", ordered by clean sheets"
		ordered = "Y"

	case else
		selected_or(0) = "selected"
		orderby_text = " years "
		ordered = ""
						
 end select	
 
 outline = outline & " <select name=""order"" style=""font-size: 10px"">"
 outline = outline & "<option value=""S"" " & selected_or(0) & ">Order by Season</option>" 
 outline = outline & "<option value=""S-"" " & selected_or(1) & ">Order by Season, latest first</option>"
 outline = outline & "<option value=""E"" " & selected_or(33) & ">Order by Position in Football League</option>"
 outline = outline & "<option value=""U"" " & selected_or(32) & ">Order by Players Used</option>"
 outline = outline & "<option value=""CS"" " & selected_or(34) & ">Order by Clean Sheets</option>"    
 outline = outline & "<option value=""P"" " & selected_or(2) & ">Order by Played</option>"  
 outline = outline & "<option value=""W"" " & selected_or(3) & ">Order by Wins</option>"
 outline = outline & "<option value=""D"" " & selected_or(4) & ">Order by Draws</option>" 
 outline = outline & "<option value=""L"" " & selected_or(5) & ">Order by Defeats</option>" 
 outline = outline & "<option value=""F"" " & selected_or(6) & ">Order by Goals For</option>" 
 outline = outline & "<option value=""A"" " & selected_or(7) & ">Order by Goals Against</option>"
 outline = outline & "<option value=""PO"" " & selected_or(8) & ">Order by Points (3-1-0 for all)</option>"
 outline = outline & "<option value=""HP"" " & selected_or(9) & ">Order by Home Played</option>"  
 outline = outline & "<option value=""HW"" " & selected_or(10) & ">Order by Home Wins</option>"
 outline = outline & "<option value=""HD"" " & selected_or(11) & ">Order by Home Draws</option>" 
 outline = outline & "<option value=""HL"" " & selected_or(12) & ">Order by Home Defeats</option>" 
 outline = outline & "<option value=""HF"" " & selected_or(13) & ">Order by Home Goals For</option>" 
 outline = outline & "<option value=""HA"" " & selected_or(14) & ">Order by Home Goals Against</option>"
 outline = outline & "<option value=""HPO"" " & selected_or(15) & ">Order by Home Points (3-1-0 for all)</option>"
 outline = outline & "<option value=""AP"" " & selected_or(16) & ">Order by Away Played</option>"  
 outline = outline & "<option value=""AW"" " & selected_or(17) & ">Order by Away Wins</option>"
 outline = outline & "<option value=""AD"" " & selected_or(18) & ">Order by Away Draws</option>" 
 outline = outline & "<option value=""AL"" " & selected_or(19) & ">Order by Away Defeats</option>" 
 outline = outline & "<option value=""AF"" " & selected_or(20) & ">Order by Away Goals For</option>" 
 outline = outline & "<option value=""AA"" " & selected_or(21) & ">Order by Away Goals Against</option>"
 outline = outline & "<option value=""APO"" " & selected_or(22) & ">Order by Away Points (3-1-0 for all)</option>"
 outline = outline & "<option value=""AT"" " & selected_or(23) & ">Order by Avg Attend</option>"
 outline = outline & "<option value=""HAT"" " & selected_or(24) & ">Order by Avg Home Attend</option>" 
 outline = outline & "<option value=""AAT"" " & selected_or(25) & ">Order by Avg Away Attend</option>" 
 outline = outline & "<option value=""CP"" " & selected_or(26) & ">Order by Cup Played</option>"  
 outline = outline & "<option value=""CW"" " & selected_or(27) & ">Order by Cup Wins</option>"
 outline = outline & "<option value=""CD"" " & selected_or(28) & ">Order by Cup Draws</option>" 
 outline = outline & "<option value=""CL"" " & selected_or(29) & ">Order by Cup Defeats</option>" 
 outline = outline & "<option value=""CF"" " & selected_or(30) & ">Order by Cup Goals For</option>" 
 outline = outline & "<option value=""CA"" " & selected_or(31) & ">Order by Cup Goals Against</option>"      
 outline = outline & "</select>"
 outline = outline & "<input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 9 0 0 0"" value=""Select options above and click here to redisplay"" name=""B1""></p>" 
 response.write(outline)
 %>
      </form>

 <%
  
  if restrictions = "Y" or ordered = "Y" then
  	if left(heading1,9) = ", ordered" then heading1 = "Ordered" & mid(heading1,10)
  	outline = "<p style=""margin-top: 0px; margin-bottom: 0px;""><font style=""font-size: 11px;"" color=""#202020""><b>" & heading1 & "</b></font>"	
	outline = outline & "<br><a href=""http://www.greensonscreen.co.uk/gosdb-seasons.asp""> [reset to initial state]</a></p>"
  	response.write(outline)
  end if
	
  %> </td>
      <td width="260" valign="top" style="text-align: center">
      <p style="margin-bottom:3pt; margin-right:3; text-align:justify; margin-top:3">Just a mass 
      of numbers? The years link on each line leads to the full season's results, and the page 
      comes into its own when you use the 'order by' box to bring out the best 
      and worst of any column. </p>
      </td>
    </tr>
  </table>
  
  <table class="seasons">
    <tr>
      <th style="border-left: 0; border-right: 0; border-top: 0" 
      colspan="5">&nbsp;</th>
      <th rowspan="2" valign="bottom">Pla-<br>
      yers<br>
      Used</th>
      <th rowspan="2" valign="bottom">Cl'n<br>
      She-<br>
      ets</th>
      <th colspan="6"><b>League Matches</b></th>
      <th colspan="6"><b>League Home</b></th>
      <th colspan="6"><b>League Away</b></th>
      <th colspan="3"><b>Attendance</b></th>
      <th colspan="6"><b>Cups</b></th>
    </tr>
    <tr>
      <th><b>Season</b></th>
      <th><b>Div</b></th>
      <th style="padding: 2 1 2 1;"><b>Tier</b></th>
      <th><b>Pos</b></th>
      <th style="padding: 2 1 2 1;"><b>FL<br>Pos*</b></th>
      <th><b>P</b></th>
      <th><b>W</b></th>
      <th><b>D</b></th>
      <th><b>L</b></th>
      <th><b>F</b></th>
      <th><b>A</b></th>
      <th><b>P</b></th>
      <th><b>W</b></th>
      <th><b>D</b></th>
      <th><b>L</b></th>
      <th><b>F</b></th>
      <th><b>A</b></th>
      <th><b>P</b></th>
      <th><b>W</b></th>
      <th><b>D</b></th>
      <th><b>L</b></th>
      <th><b>F</b></th>
      <th><b>A</b></th>
      <th style="text-align: center"><b>Avg<br>All</b></th>
      <th style="text-align: center"><b>Avg<br>Home</b></th>
      <th style="text-align: center"><b>Avg<br>Away</b></th>
      <th><b>P</b></th>
      <th><b>W</b></th>
      <th><b>D</b></th>
      <th><b>L</b></th>
      <th><b>F</b></th>
      <th><b>A</b></th>
    </tr>

<% 
sql = "WITH seasonCTE1 AS ( "
sql = sql & "select season_no, years, division_short, tier, endpos, teams_above_div, promrel, sum(cs) as clean_sheets, "
sql = sql & "sum(p) as P, sum(w) as W, sum(d) as D, sum(l) as L, sum(f) as F, sum(a) as A, sum(po) as PO, avg(at) as AT, "
sql = sql & "sum(hp) as HP, sum(hw) as HW, sum(hd) as HD, sum(hl) as HL, sum(hf) as HF, sum(ha) as HA, sum(hpo) as HPO, avg(hat) as HAT,  "
sql = sql & "sum(ap) as AP, sum(aw) as AW, sum(ad) as AD, sum(al) as AL, sum(af) as AF, sum(aa) as AA, sum(apo) as APO, avg(aat) as AAT, "
sql = sql & "sum(cp) as CP, sum(cw) as CW, sum(cd) as CD, sum(cl) as CL, sum(cf) as CF, sum(ca) as CA "
sql = sql & "from ( "
sql = sql & "select season_no, years, division_short, tier, endpos, teams_above_div, promrel, "
sql = sql & "case when LFC <> 'C' then 1 else 0 end as p, "
sql = sql & "case when LFC <> 'C' and goalsfor > goalsagainst then 1 else 0 end as w, "
sql = sql & "case when LFC <> 'C' and goalsfor = goalsagainst then 1 else 0 end as d, "
sql = sql & "case when LFC <> 'C' and goalsfor < goalsagainst then 1 else 0 end as l, "
sql = sql & "case when LFC <> 'C' then goalsfor else 0 end as f,  "
sql = sql & "case when LFC <> 'C' then goalsagainst else 0 end as a,  "
sql = sql & "case when LFC <> 'C' and goalsagainst = 0 then 1 else 0 end as cs,  "
sql = sql & "case when LFC <> 'C' and goalsfor > goalsagainst then 3 when LFC <> 'C' and goalsfor = goalsagainst then 1 else 0 end as po,  "
sql = sql & "case when LFC <> 'C' then attendance else NULL end as at,  "
sql = sql & "case when LFC <> 'C' and homeaway = 'H' then 1 else 0 end as hp, "
sql = sql & "case when LFC <> 'C' and homeaway = 'H' and goalsfor > goalsagainst then 1 else 0 end as hw, "
sql = sql & "case when LFC <> 'C' and homeaway = 'H' and goalsfor = goalsagainst then 1 else 0 end as hd, "
sql = sql & "case when LFC <> 'C' and homeaway = 'H' and goalsfor < goalsagainst then 1 else 0 end as hl, "
sql = sql & "case when LFC <> 'C' and homeaway = 'H' then goalsfor else 0 end as hf, "
sql = sql & "case when LFC <> 'C' and homeaway = 'H' then goalsagainst else 0 end as ha, "
sql = sql & "case when LFC <> 'C' and homeaway = 'H' and goalsfor > goalsagainst then 3 when LFC <> 'C' and homeaway = 'H' and goalsfor = goalsagainst then 1 else 0 end as hpo,  "
sql = sql & "case when LFC <> 'C' and homeaway = 'H' then attendance else NULL end as hat, "
sql = sql & "case when LFC <> 'C' and homeaway = 'A' then 1 else 0 end as ap, "
sql = sql & "case when LFC <> 'C' and homeaway = 'A' and goalsfor > goalsagainst then 1 else 0 end as aw, "
sql = sql & "case when LFC <> 'C' and homeaway = 'A' and goalsfor = goalsagainst then 1 else 0 end as ad, "
sql = sql & "case when LFC <> 'C' and homeaway = 'A' and goalsfor < goalsagainst then 1 else 0 end as al, "
sql = sql & "case when LFC <> 'C' and homeaway = 'A' then goalsfor else 0 end as af, "
sql = sql & "case when LFC <> 'C' and homeaway = 'A' then goalsagainst else 0 end as aa, "
sql = sql & "case when LFC <> 'C' and homeaway = 'A' and goalsfor > goalsagainst then 3 when LFC <> 'C' and homeaway = 'A' and goalsfor = goalsagainst then 1 else 0 end as apo,  "
sql = sql & "case when LFC <> 'C' and homeaway = 'A' then attendance else NULL end as aat, "
sql = sql & "case when LFC = 'C' then 1 else 0 end as cp, "
sql = sql & "case when LFC = 'C' and goalsfor > goalsagainst then 1 else 0 end as cw, "
sql = sql & "case when LFC = 'C' and goalsfor = goalsagainst then 1 else 0 end as cd, "
sql = sql & "case when LFC = 'C' and goalsfor < goalsagainst then 1 else 0 end as cl, "
sql = sql & "case when LFC = 'C' then goalsfor else 0 end as cf, "
sql = sql & "case when LFC = 'C' then goalsagainst else 0 end as ca "
sql = sql & "from v_match_all join season on date between date_start and date_end "
sql = sql & "where season_no between " & season_no1 & " and " & season_no2 & homeaway_text
sql = sql & ") as subsel "
sql = sql & "group by season_no, years, division_short, tier, endpos, teams_above_div, promrel "
sql = sql & "), "
sql = sql & "seasonCTE2 AS ( "
sql = sql & "select season_no, count(distinct player_id) as player_count "
sql = sql & "from match_player join season on date between date_start and date_end "
sql = sql & "where season_no between " & season_no1 & " and " & season_no2 & homeaway_text
sql = sql & "group by season_no " 
sql = sql & ") "
sql = sql & "select player_count, years, division_short, tier, endpos, endpos + teams_above_div as flendpos, promrel, clean_sheets, " 
sql = sql & "P, W, D, L, F, A, PO, AT, HP, HW, HD, HL, HF, HA, HPO, HAT, AP, AW, AD, AL, AF, AA, APO, AAT, CP, CW, CD, CL, CF, CA " 
sql = sql & "from seasonCTE1 x join seasonCTE2 y on x.season_no = y.season_no " 
sql = sql & "order by" & orderby_text

rs.open sql,conn,1,2

outline = ""

Do While Not rs.EOF

	if isnull(rs.Fields("AT")) then 
		att = "?  "
	  else 
		att = rs.Fields("AT")
		if len(att) > 3 then att = left(rs.Fields("AT"),len(rs.Fields("AT"))-3) & "," & right(rs.Fields("AT"),3)
	end if
	if isnull(rs.Fields("HAT")) then 
		homeatt = "?  "
	  else 
		homeatt = rs.Fields("HAT")
		if len(homeatt) > 3 then homeatt = left(rs.Fields("HAT"),len(rs.Fields("HAT"))-3) & "," & right(rs.Fields("HAT"),3)
	end if
	if isnull(rs.Fields("AAT")) then 
		awayatt = "?  "
	  else 
		awayatt = rs.Fields("AAT")
		if len(awayatt) > 3 then awayatt = left(rs.Fields("AAT"),len(rs.Fields("AAT"))-3) & "," & right(rs.Fields("AAT"),3)
	end if
	
	outline  = outline & "<tr onmouseover=""this.className = 'rowhlt';"" onmouseout=""this.className = '';"">"	
	outline  = outline & "<td class=""season"" nowrap=""nowrap"" style=""position: relative; text-align: left;""><a href=""gosdb-season.asp?years=" & rs.Fields("years") & """>"
	outline  = outline & "<img style=""vertical-align: middle; margin:0 4px 0 0; padding:0"" src=""images/more.png"">" & rs.Fields("years")
	if rs.Fields("promrel") = "P" then outline  = outline & " <img src=""images/promote.gif"" border=""0"">"
	if rs.Fields("promrel") = "R" then outline  = outline & " <img src=""images/relegate.gif"" border=""0"">"
	outline  = outline & "</a></td>"
	if rs.Fields("years") = "1903-1904" or rs.Fields("years") = "1904-1905" or rs.Fields("years") = "1905-1906" or rs.Fields("years") = "1906-1907" or rs.Fields("years") = "1907-1908" or rs.Fields("years") = "1908-1909" then 
		division = "SL+WL"
	  elseif rs.Fields("years") = "1939-1940" then
	  	division = "D2 +<br>SWRL"
	  else
	   	division = rs.Fields("division_short") 
	end if
	outline  = outline & "<td style=""text-align: left;"">" & division & "</td>"
	outline  = outline & "<td style=""text-align: center"">" & rs.Fields("tier") & "</td>"
	outline  = outline & "<td style=""text-align: center""> " & rs.Fields("endpos") & "</td>"
	outline  = outline & "<td style=""text-align: center""> " & rs.Fields("flendpos") & "</td>"
	outline  = outline & "<td style=""text-align: center""> " & rs.Fields("player_count") & "</td>"
	outline  = outline & "<td style=""text-align: center""> " & rs.Fields("clean_sheets") & "</td>"
	
	outline  = outline & "<td style=""border-left: 2px solid #a0a0a0;"">" & rs.Fields("P") & "</td>"	
	outline  = outline & "<td>" & rs.Fields("W") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("D") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("L") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("F") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("A") & "</td>" 
	outline  = outline & "<td style=""border-left: 2px solid #a0a0a0;"">" & rs.Fields("HP") & "</td>"
	outline  = outline & "<td>" & rs.Fields("HW") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("HD") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("HL") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("HF") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("HA") & "</td>" 
	outline  = outline & "<td style=""border-left: 2px solid #a0a0a0;"">" & rs.Fields("AP") & "</td>"
	outline  = outline & "<td>" & rs.Fields("AW") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("AD") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("AL") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("AF") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("AA") & "</td>" 
	outline  = outline & "<td align=""right"" style=""border-left: 2px solid #a0a0a0;"">" & att & "</td>"
	outline  = outline & "<td align=""right"">" & homeatt & "</td>"
	outline  = outline & "<td align=""right"">" & awayatt & "</td>"
	outline  = outline & "<td style=""border-left: 2px solid #a0a0a0;"">" & rs.Fields("CP") & "</td>"
	outline  = outline & "<td>" & rs.Fields("CW") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("CD") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("CL") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("CF") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("CA") & "</td>" 
	outline  = outline & "</tr>" 
	
	rs.MoveNext
	
Loop

response.write(outline)
	
rs.close
conn.close		

%>

</tbody></table>

<p style="margin-top: 8; margin-bottom: 12"><b>*</b> 'FL Pos' out of 92 possible places, except for 1920-50 (66 places) and 1950-58 (68 places)
</p>
</div>
<!--#include file="base_code.htm"-->

</body>

</html>