<%@ Language=VBScript %> 
<% Option Explicit %>
<!doctype html>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=ISO-8859-1" />
<title>GoS-DB Miscellaneous Report</title>
<link href="http://www.greensonscreen.co.uk/images/favicon.ico" rel="shortcut icon">
<link rel="stylesheet" type="text/css" href="gos2.css">
<style>
<!--
#table1 td {border: 1px solid #c0c0c0; margin: 0; white-space:nowrap;}
#table2 td {border-bottom: 1px solid #c0c0c0; margin: 0; padding: 4px; white-space:nowrap;}
#table2 td a {padding:3px; margin:0}
#scorelist {margin:15px auto}
.score {min-width:30px; text-align:right; background-color:#e0f0e0; padding:3px 2px 3px 4px;}
.score-never {min-width:30px; text-align:right; background-color:#f4f4f4; padding:3px 2px 3px 4px;}
.scorecount {width:24px; text-align:right;  padding:3px 4px 3px 2px;}
.dummycell {width:3px; border-top-style: none !important; border-bottom-style: none !important;}
.hover {color: #000000; background-color: #c8e0c7 !important; cursor: pointer; font-weight: 700;}
.right {text-align:right;}
.nowrap {white-space:nowrap;}
.nowrap .noverticalborder {border-left:0px none; border-right:0px none;}
-->
</style>

<script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script>
<script>
$(document).ready(function(){

	$(".score").hover(function() {
    	$(this).toggleClass("hover");
	});
	
    $('#table1').on('click','.score', function(){
		var ajaxparm = "input=" + $(this).attr('id'); 
		$('#scorelist').load('gosdb-getscorelines.asp?' + ajaxparm);
		$("#scorelist").show('slow');
	});
	   
});
</script>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline, heading, season_no1, selected_s1, season_no2, selected_s2, season1opts, season2opts, selyears1, selyears2, lastyears, scorearray(10,10), i, j
Dim homeaway, HAvalue, selected_comp(2), selected_HA(2), competition, LFCvalue, result, restrictions


Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%
%>
    <table border="0" cellspacing="0" style="margin:auto; border-collapse: collapse" cellpadding="0" width="980">
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
	<p style="margin:12px 0 0; text-align:center; font-size:18px; color:#006E32">MISCELLANEOUS REPORTS</p>  
	<p style="margin:6px 0; text-align:center; font-size:13px">
    <b>Report 11: Score Counts</b></p>  
    </td>
        
	<td width="260" valign="top"  align="justify">
	<span style="font-size: 10px">'Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
    </td>
    </tr>   
	</table>
		
	<form style="font-size: 10px; padding: 3px; margin: 0 auto;" action="gosdb-misc11.asp" method="post" name="form1">
      
<%
homeaway = Request.Form("homeaway")
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
 
 select case homeaway
	case "H"
		selected_HA(1) = "selected"
		HAvalue = "'H'"
		heading = heading & ", Home"
	case "A"
		selected_HA(2) = "selected"
		HAvalue = "'A'"
		heading = heading & ", Away"
	case else
		selected_HA(0) = "selected"
		HAvalue = "'H','A'"
		heading = heading & ", Home & Away"
 end select

 outline = "<select name=""competition"" style=""font-size: 10px; margin-right:3px;"">"
 outline = outline & "<option value=""ALL"" " & selected_comp(0) & ">All Competitions</option>"
 outline = outline & "<option value=""FLG"" " & selected_comp(1) & ">Football League</option>" 
 outline = outline & "<option value=""CUP"" " & selected_comp(2) & ">All Cups</option>" 
 outline = outline & "</select>"

 outline = outline & "<select name=""homeaway"" style=""font-size: 10px;"">"
 outline = outline & "<option value=""ALL"" " & selected_HA(0) & ">Home & Away</option>"
 outline = outline & "<option value=""H"" " & selected_HA(1) & ">Home only</option>" 
 outline = outline & "<option value=""A"" " & selected_HA(2) & ">Away only</option>" 
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
   	selyears1 = rs.Fields("years") 
   else selected_s1 = ""
  end if
  season1opts = season1opts & "<option value=""" & rs.Fields("season_no") & """ " & selected_s1 & ">From " & rs.Fields("years") & "</option>"
  if CStr(rs.Fields("season_no")) = season_no2 then
    selected_s2 = "selected"
    selyears2 = rs.Fields("years") 
   else selected_s2 = ""
  end if
  season2opts = season2opts & "<option value=""" & rs.Fields("season_no") & """ " & selected_s2 & ">To " & rs.Fields("years") & "</option>"
  lastyears = rs.Fields("years")
 rs.MoveNext
 Loop
 
 if season_no1 > 1 and CInt(season_no2) < rs.RecordCount then heading = heading & ", from " & selyears1 & " to " & selyears2
 if season_no1 = 1 and CInt(season_no2) < rs.RecordCount then heading = heading & ", from 1903-1904 to " & selyears2
 if season_no1 > 1 and CInt(season_no2) = rs.RecordCount then heading = heading & ", from " & selyears1 & " to " & lastyears
 
 rs.close
 
 outline = outline & " <select name=""season1"" style=""font-size: 10px;"">" & season1opts & "</select>"  
 outline = outline & " <select name=""season2"" style=""font-size: 10px"">" & season2opts & "</select>"
 outline = outline & "<br><input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 2px 5px; margin: 9px 0 0 0"" value=""Select options and click to redisplay"" name=""B1""></p>"  
 outline = outline & "</form>"
 
 outline = outline & "<p class=""style5boldgreen"" style=""margin:12px 0 0;"">" & heading & "</p>"
 outline = outline & "<p class=""style1"" style=""margin:6px 0 0;"">Note: The scores show Argyle first, both home and away</p>"
 outline = outline & "<p class=""style1bold"" >Click on a score for the most recent results (up to 50 displayed)</p>"
 
	'initialise scorearray
	for i = 0 to 10
		for j = 0 to 10
			scorearray(i,j) = 0
		next 
	next
		
	sql = "select goalsfor, goalsagainst, count(*) as scorecount " 
	sql = sql & "from v_match_all join season on date between date_start and date_end "
	sql = sql & "where season_no between " & season_no1 & " and " & season_no2
	sql = sql & "  and LFC in (" & LFCvalue & ") "
	sql = sql & "  and homeaway in (" & HAvalue & ") "
	sql = sql & "group by goalsfor, goalsagainst "
	sql = sql & "order by goalsfor, goalsagainst "
	rs.open sql,conn,1,2
	
	outline = outline & "<table id=""table1"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"">"

	Do While Not rs.EOF
			scorearray(rs.Fields("goalsfor"),rs.Fields("goalsagainst")) = rs.Fields("scorecount")
			rs.MoveNext
	Loop
	
	rs.close
		
	'now display the array
	
	for i = 0 to 10
		for j = 0 to 10	
			if scorearray(i,j) > 0 then 
			outline = outline & "<td id=""" & i & "-" & j & "-" & replace(replace(LFCvalue,",",""),"'","") & "-" & replace(replace(HAvalue,",",""),"'","") & "-" & season_no1 & "-" & season_no2 & """ class=""score"">" & i & "-" & j & "</td>" 
			outline = outline & "<td class=""scorecount"">" & scorearray(i,j) & "</td>"
		  else
			outline = outline & "<td class=""score-never"">" & i & "-" & j & "</td>" 
			outline = outline & "<td class=""scorecount""></td>"
		end if
		if j < 10 then outline = outline & "<td class=""dummycell""></td>"
 		next
	  outline  = outline & "<tr>"
	 next

	outline  = outline & "</table>"
	response.write(outline)
	%>

<%
conn.close
%>
	
<div id="scorelist"></div>	
<br>

<!--#include file="base_code.htm"-->
</body>

</html>