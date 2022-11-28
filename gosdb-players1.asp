<%@ Language=VBScript %> 
<% Option Explicit %>
<% dim scope, rank
scope = Request.Form("scope")
if scope = "" then scope = Request.Querystring("scope")	'try for a url parameter
scope = replace(scope," ","")
if scope = "" then scope = "1,2,3,4,5,6,7" 
rank = Request.QueryString("rank")
%>
<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-19">
<title>GoS-DB Players</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
#ajaxplayerlist p {margin: 0 0 0 12px; padding: 0;  font-size: 11px; font-weight:normal; text-align: left; color:#202020; }

#ranktable1 td {text-align:left; margin: 0; padding: 0 3px;  font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 12px; }
#ranktable1 .right {text-align: right; } 
#ranktable1 .center {text-align: center; } 
#ranktable1 .bold {font-weight: bold; }
#ranktable1 .head1 {font-family: verdana,arial,sans-serif; font-size: 11px; padding: 4px; } 
#ranktable1 .head2 {font-family: verdana,arial,sans-serif; font-size: 11px; } 
#ranktable2 td {text-align:left; margin: 0; padding: 0 3px;  font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 12px; }
#ranktable2 .right {text-align: right; } 
#ranktable2 .bold {font-weight: bold; }
#ranktable2 .head1 {font-family: verdana,arial,sans-serif; font-size: 11px; padding: 4px; } 
#ranktable2 .head2 {font-family: verdana,arial,sans-serif; font-size: 11px; } 
#ranktable3 td {text-align:left; margin: 0; padding: 0 3px;  font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 12px; }
#ranktable3 .right {text-align: right; } 
#ranktable3 .bold {font-weight: bold; }
#ranktable3 .head1 {font-family: verdana,arial,sans-serif; font-size: 11px; padding: 4px; } 
#ranktable3 .head2 {font-family: verdana,arial,sans-serif; font-size: 11px; } 
#ranktable4 td {text-align:left; margin: 0; padding: 0 3px;  font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 12px; }
#ranktable4 .right {text-align: right; } 
#ranktable4 .bold {font-weight: bold; }
#ranktable4 .head1 {font-family: verdana,arial,sans-serif; font-size: 11px; padding: 4px; } 
#ranktable4 .head2 {font-family: verdana,arial,sans-serif; font-size: 11px; } 

-->
</style>
<%
' *** Ajax Overview ***
' The player name box calls GetPlayerlist which fires gosdb-getplayerspageplayerlist.asp
'  * gosdb-getplayerspageplayerlist.asp produces list of players with link to GetPlayer, which fires gosdb-getplayerdetails-full.asp
'  * gosdb-getplayerdetails-full.asp produces season list with [+] using Toggle2 and id's based on 'tag', which fires gosdb-getplayerspagematchlist.asp
'  * gosdb-getplayerspagematchlist.asp produces a match list with [+] using Toggle3 and id's based on 'mattag', which fires gosdb-getmatchdetails1.asp
%>

<script language="javascript">

function GetPlayerList(initial,scope) { 

try { 
        // Moz supports XMLHttpRequest. IE uses ActiveX. 
        // browser detction is bad. object detection works for any browser 
        xmlhttp = window.XMLHttpRequest?new XMLHttpRequest(): new ActiveXObject("Microsoft.XMLHTTP"); 
} catch (e) { 
        // browser doesn't support ajax. handle however you want 
          alert ("Sorry, your browser does not support this function.");
  		  return;
} 
// the xmlhttp object triggers an event everytime the status changes 
// triggered1() function handles the events 
xmlhttp.onreadystatechange = triggered1; 

// open takes in the HTTP method and url.
document.body.style.cursor='wait';   
var url="gosdb-getplayerspageplayerlist.asp";     
url=url+"?initial="+initial;
url=url+"&scp="+scope;
url=url+"&sid="+Math.random();

xmlhttp.open("GET", url, true);
document.getElementById('ajaxplayerlist').innerHTML = '<img style="margin: 0 0 0 12;" border="0" src="images/ajax-loader.gif"><br><img border="0"" src="images/dummbar_0.gif" height="260" width="1">' 
 
// send the request. if this is a POST request we would have 
// sent post variables: send("name=aleem&gender=male) 
// Moz is fine with just send(); but 
// IE expects a value here, hence we do send(null); 
xmlhttp.send(null);
document.body.style.cursor='auto';  
} 
 
function triggered1() { 
// if the readyState code is 4 (Completed) 
// and http status is 200 (OK) we go ahead and get the responseText 
// other readyState codes: 
// 0=Uninitialised 1=Loading 2=Loaded 3=Interactive 
if (xmlhttp.readyState == 4) { 
        // xmlhttp.responseText object contains the response.
		 document.getElementById('ajaxplayerlist').innerHTML = xmlhttp.responseText ;
} 
} 

</script>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline, tableview 
Dim competition, season_no1, selected_s1, season_no2, selected_s2, season1opts, season2opts, selyears1, selyears2, goalspergame
Dim i, j, n, work1, work2, playername, spells 

season_no1 = Request.Form("season1")
season_no2 = Request.Form("season2")

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
    
  	<td width="460" align="center" style="text-align: center">	
	<p style="margin-top:9; margin-bottom:0; text-align:center; font-size:18px; color:#006E32">
    THE PLAYERS</p>  
    
    <%
    outline = "<p style=""color: #CC3300; margin: 12 30 6 30; text-align:center; ""><b>Competitions selected: "
    
    if scope = "1,2,3,4,5,6,7" then
    
    	outline = outline & "All</b></p>"
    	
      else

		sql = "select distinct LFC, compcat, compcatname " 
		sql = sql & "from competition " 
		sql = sql & "where compcat in (" & scope & ") "
		sql = sql & "order by compcat "
		rs.open sql,conn,1,2
	
		Do While Not rs.EOF
	  	outline = outline & rs.Fields("compcatname") & ", "
	  	rs.MoveNext
		Loop
		rs.close 
	
		outline = left(outline,len(outline)-2) & "</b>"  	'remove last comma and space 
		
		outline = outline & " (excluded: "

		sql = "select distinct LFC, compcat, compcatname " 
		sql = sql & "from competition " 
		sql = sql & "where not compcat in (" & scope & ") "
		sql = sql & "order by compcat "
		rs.open sql,conn,1,2
	
		Do While Not rs.EOF
	  	outline = outline & rs.Fields("compcatname") & ", "
	  	rs.MoveNext
		Loop
		rs.close 
	
		outline = left(outline,len(outline)-2) & ")</p>"  	'remove last comma and space 
		
	end if	
	
	outline = outline & "<p style=""margin: 6 0 6 0; text-align:centre""><a href=""gosdb-players0.asp""><u>Reselect Competitions</u></a></p>"
		
    response.write(outline)
    
    %>	
    </td>
        
	<td width="260" valign="top"  align="right">
	
	<%	 
	If Request.Form("button") = "Player Rankings" or rank = "rank" then
		outline = "<p style=""margin:0; text-align:justify"">Top 100 lists for the chosen competitions, with an option to limit the results to a range of seasons. <i>Note: the first list - Time at Club - is for all competitions, regardless of selection.</i></p>"
	  else
	  	outline = "<p style=""margin:0; text-align:justify"">A full PAFC history of every first-team player since 1903, with details relevant for the competitions selected. Note that the names listed here are from the complete set; they are not limited to those playing in the chosen competitions.</p>" 

	end if
    
    response.write(outline)
	%>	
	
     
    </td>
    </tr>
    
    <tr>
    <td width="980" valign="top" style="text-align:center;" colspan="3">
    
 <%
 outline = ""
 
 If Request.Form("button") = "Player Rankings" or rank = "rank" then
 
 outline = outline & "<div>"
 outline = outline & "<center>"
 outline = outline & "<table border=""0"" cellpadding=""0"" cellspacing=""0""  style=""border-collapse: collapse"" width=""980px"">"
 outline = outline & "<tr><td colspan=""4"" style=""text-align: center; border: 0px solid #c0c0c0;"" >" 

 outline = outline & "<form style=""font-size: 10px; padding: 0; margin: 0; text-align: center;"" action=""gosdb-players1.asp"" method=""post"" name=""form1"">"

 
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
 rs.MoveNext
 Loop
 
 rs.close
 
 outline = outline & "<select name=""season1"" style=""font-size: 10px"">" & season1opts & "</select>"  
 outline = outline & "<select name=""season2"" style=""font-size: 10px"">" & season2opts & "</select>"
 outline = outline & "<input type=""hidden"" value=" & scope & " name=""scope"">" 
 outline = outline & "<input type=""hidden"" value=""Player Rankings"" name=""button"">" 

 outline = outline & "<input type=""submit"" style=""width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; padding: 1 5 1 5; margin: 0 0 0 0"" value=""Redisplay"" name=""B1""></p>" 
 
 outline = outline & "</form>"
 
 outline = outline & "</td></tr>"
   	
 outline = outline & "<tr>"
 outline = outline & "<td valign=""top"">"

	sql = "with detailCTE as "
	sql = sql & "( 	"
	sql = sql & "select spell, a.player_id_spell1, surname, forename, initials, datediff(day,d.date,e.date)+1 as duration "
	sql = sql & "from player a "
	sql = sql & "join match_player b on a.player_id = b.player_id "
	sql = sql & "join match_player c on a.player_id = c.player_id "
	sql = sql & "join v_match_all d on b.date = d.date join season f on d.date between f.date_start and f.date_end "	
	sql = sql & "join v_match_all e on c.date = e.date join season g on e.date between g.date_start and g.date_end "
	sql = sql & "and d.date = ( "
 	sql = sql & "select min(b1.date) "
 	sql = sql & "from player a1 "
	sql = sql & "	join match_player b1 on a1.player_id = b1.player_id "
	sql = sql & "   join v_match_all d1 on b1.date = d1.date join season f1 on d1.date between f1.date_start and f1.date_end "
	sql = sql & "	where a1.player_id = a.player_id "
	sql = sql & "   and f1.season_no between " & season_no1 & " and " & season_no2 & " "
	sql = sql & "	) "
	sql = sql & "and e.date = ( "
 	sql = sql & "select max(c2.date) "
 	sql = sql & "from player a2 "
	sql = sql & "	join match_player c2 on a2.player_id = c2.player_id "
	sql = sql & "   join v_match_all e2 on c2.date = e2.date join season g2 on e2.date between g2.date_start and g2.date_end "
	sql = sql & "	where a2.player_id = a.player_id "
	sql = sql & "   and g2.season_no between " & season_no1 & " and " & season_no2 & " "
	sql = sql & "	) "	
	sql = sql & "), "
	sql = sql & "  sumCTE as "
	sql = sql & "( 	"
	sql = sql & "select top 100 player_id_spell1, surname, forename, initials, sum(duration) as totduration "
	sql = sql & "from detailCTE "
	sql = sql & "group by player_id_spell1, surname, forename, initials "
	sql = sql & "order by totduration desc, surname "
	sql = sql & "), "
	sql = sql & "spellCTE as "
	sql = sql & "( "
	sql = sql & "select player_id_spell1, max(spell) as maxspell "
	sql = sql & "from player "
	sql = sql & "group by player_id_spell1 "
	sql = sql & ") "
	sql = sql & "select rank() over (order by totduration desc) as rank, a.player_id_spell1, surname, forename, initials, maxspell, "
	sql = sql & "floor(totduration/365.25) as years, cast(round(totduration - (365.25*floor(totduration/365.25)),0) as integer) as days "
	sql = sql & "from sumCTE a join spellCTE b on a.player_id_spell1 = b.player_id_spell1; "
	
	rs.open sql,conn,1,2
	
	outline = outline & "<table id=""ranktable1"" border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#e0e0e0"">"
    outline = outline & "<tr><td nowrap colspan=""2"" class=""head1""><b>Ranked by Time at Club</b><br>(see footnote)</td><td class=""head2 center"">Spells</td><td class=""head2 right"">Yrs</td><td class=""head2 right"">Days</td></tr>"

	Do While Not rs.EOF
	  if not IsNull(rs.Fields("forename")) then 
	  		playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
	  	else
	  		playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
	  end if
	  if rs.Fields("maxspell") > 1 then
			spells = rs.Fields("maxspell")
	    else
	  		spells = ""
	  end if		  	
	  outline = outline & "<tr><td class=""right"">" & rs.Fields("rank") & "</td><td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & "&scp=" & scope & """>" & playername & "</a></td><td class=""center"">" & spells & "</td><td class=""right"">" & rs.Fields("years") & "</td><td class=""right"">" & rs.Fields("days") & "</td></tr>"
	  rs.MoveNext
	Loop
	rs.close

	outline = outline & "</table>"
	
   	outline = outline & "</td><td valign=""top"">"
	
	sql = "with detailCTE as "
	sql = sql & "( 	"
	sql = sql & "select d.player_id_spell1, surname, forename, initials, 1 as starts, 0 as subs " 
	sql = sql & "from v_match_all a join season on date between date_start and date_end " 
	sql = sql & "join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where season_no between " & season_no1 & " and " & season_no2
	sql = sql & " and d.player_id <> 9000 and startpos > 0 "
	sql = sql & " and a.compcat in (" & scope & ") "
	sql = sql & "union all "
	sql = sql & "select d.player_id_spell1, surname, forename, initials, 0 as starts, 1 as subs " 
	sql = sql & "from v_match_all a join season on date between date_start and date_end " 
	sql = sql & "join match_player b on a.date = b.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where season_no between " & season_no1 & " and " & season_no2
	sql = sql & " and d.player_id <> 9000 and startpos = 0 "
	sql = sql & " and a.compcat in (" & scope & ") "
	sql = sql & "), "
	sql = sql & "  sumCTE as "
	sql = sql & "( 	"
	sql = sql & "select top 100 player_id_spell1, surname, forename, initials, sum(starts) as totstarts, sum(subs) as totsubs, sum(starts+subs) as tot "
	sql = sql & "from detailCTE "
	sql = sql & "group by player_id_spell1, surname, forename, initials "
	sql = sql & "order by tot desc, surname "
	sql = sql & ") "
	sql = sql & "select rank() over (order by tot desc) as rank, player_id_spell1, surname, forename, initials, totstarts, totsubs, tot "
	sql = sql & "from sumCTE "
	rs.open sql,conn,1,2
	
	outline = outline & "<table id=""ranktable2"" border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#c0c0c0"">"
    outline = outline & "<tr><td colspan=""2"" class=""head1 bold"">Ranked by<br>Appearances</td><td class=""head2 right"">Starts</td><td class=""head2 right"">Subs</td></tr>"

	Do While Not rs.EOF
	  if not IsNull(rs.Fields("forename")) then 
	  		playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
	  	else
	  		playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
	  end if
	  outline = outline & "<tr><td class=""right"">" & rs.Fields("rank") & "</td><td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & "&scp=" & scope & """>" & playername & "</a></td><td class=""right"">" & rs.Fields("totstarts") & "</td><td class=""right"">" & rs.Fields("totsubs") & "</td></tr>"
	  rs.MoveNext
	Loop
	rs.close

	outline = outline & "</table>"
		
   	outline = outline & "</td><td valign=""top"">"
	
	sql = "with CTE as "
	sql = sql & "( 	"
	sql = sql & "select top 100 player_id_spell1, surname, forename, initials, count(c.player_id) as goals, round(count(c.player_id)/cast(count(distinct b.date) as dec(7,3)),2) as pergame "
	sql = sql & "from v_match_all a join season on date between date_start and date_end "
	sql = sql & "join match_player b on a.date = b.date " 
	sql = sql & "left outer join match_goal c on b.player_id = c.player_id and b.date = c.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where season_no between " & season_no1 & " and " & season_no2
	sql = sql & " and d.player_id <> 9000 "
	sql = sql & " and a.compcat in (" & scope & ") "
	sql = sql & "group by player_id_spell1, surname, forename, initials "
	sql = sql & "order by goals desc, surname "
	sql = sql & ") "
	sql = sql & "select rank() over (order by goals desc) as rank, player_id_spell1, surname, forename, initials, pergame, goals "
	sql = sql & "from CTE "		
	
	rs.open sql,conn,1,2
	
	outline = outline & "<table id=""ranktable3"" border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#c0c0c0"">"
    outline = outline & "<tr><td colspan=""2"" class=""head1 bold"">Ranked by<br>Goals Scored</td><td class=""head2 right"">Goals</td><td class=""head2 right"">Goals<br>/Game</td></tr>"

	Do While Not rs.EOF
	  if not IsNull(rs.Fields("forename")) then 
	  		playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
	  	else
	  		playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
	  end if
	  
	  goalspergame = rs.Fields("pergame") 
	  if Instr(rs.Fields("pergame"),".") = 0 then goalspergame = goalspergame & "."  'no decimal point, must be a whole number, so add a dec. point
	  goalspergame = left(goalspergame & "00",4)
	  
	  outline = outline & "<tr><td class=""right"">" & rs.Fields("rank") & "</td><td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & "&scp=" & scope & """>" & playername & "</a></td><td class=""right"">" & rs.Fields("goals") & "</td><td class=""right"">" & goalspergame & "</td></tr>"
  	  rs.MoveNext
	Loop
	rs.close

	outline = outline & "</table>"
	
   	outline = outline & "</td><td valign=""top"">"

	sql = "with CTE as "
	sql = sql & "( 	"
	sql = sql & "select top 100 player_id_spell1, surname, forename, initials, "
	sql = sql & "count(c.player_id) as goals, round(count(c.player_id)/cast(count(distinct b.date) as dec(7,3)),3) as pergame "
	sql = sql & "from v_match_all a join season on date between date_start and date_end "
	sql = sql & "join match_player b on a.date = b.date " 
	sql = sql & "left outer join match_goal c on b.player_id = c.player_id and b.date = c.date " 
	sql = sql & "join player d on b.player_id = d.player_id "
	sql = sql & "where season_no between " & season_no1 & " and " & season_no2
	sql = sql & " and d.player_id <> 9000 "
	sql = sql & " and a.compcat in (" & scope & ") "
	sql = sql & "group by player_id_spell1, surname, forename, initials " 
	sql = sql & "order by pergame desc, surname "
	sql = sql & ") "
	sql = sql & "select rank() over (order by pergame desc) as rank, player_id_spell1, surname, forename, initials, pergame, goals "
	sql = sql & "from CTE "		
	rs.open sql,conn,1,2
	
	outline = outline & "<table id=""ranktable4"" border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#c0c0c0"">"
    outline = outline & "<tr><td colspan=""2"" class=""head1 bold"">Ranked by<br>Goals per Game</td><td class=""head2 right"">Goals<br>/Game</td><td class=""head2 right"">Goals</td></tr>"

	Do While Not rs.EOF
	  if not IsNull(rs.Fields("forename")) then 
	  		playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("forename"))
	  	else
	  		playername = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
	  end if
	  
	  goalspergame = rs.Fields("pergame") 
	  if Instr(rs.Fields("pergame"),".") = 0 then goalspergame = goalspergame & "."  'no decimal point, must be a whole number, so add a dec. point
	  goalspergame = left(goalspergame & "000",5)
	  
	  outline = outline & "<tr><td class=""right"">" & rs.Fields("rank") & "</td><td><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & "&scp=" & scope & """>" & playername & "</a></td><td class=""right"">" & goalspergame & "</td><td class=""right"">" & rs.Fields("goals") & "</td></tr>"
  	  rs.MoveNext
	Loop
	rs.close

	outline = outline & "</table>"
	
	outline = outline & "</td></tr></table>"
	
	outline = outline & "<p style=""margin:12px; text-align:left;""><span style=""font-weight:700"">Ranked by Time at Club:</span> the total time between first and final first-team games, whilst on PAFC's books. For players who left the club and then returned for a second or third spell, the time for each spell is accumulated. Careers interupted by war are broken into separate spells, so the war years are not included in the results. Loans from other clubs count as a separate spell; loans to other clubs have no impact.</p>" 

	outline = outline & "</div>"

	
	response.write(outline)

else 

	outline = outline & "<p style=""margin-left:0; margin-right:0; margin-top:12; margin-bottom:6"">" 
    outline = outline & "<b>To find a player,</b> type the first<br>few letters of the surname:</p>"

	outline = outline & "<form style=""padding:0; margin:0;"" name='form1'>"
	outline = outline & "<p style=""margin-left: 0; margin-top:0; margin-bottom:0"">"
	outline = outline & "<input type='text' name='surname' onKeyUp=""GetPlayerList(this.value,'" & scope & "')"" size=""10"" > </p>"
	outline = outline & "<input type=""hidden"" value=" & scope & " name=""scope"">"
	outline = outline & "</form>"
	
	outline = outline & "<p style=""margin-left:0; margin-right:0; margin-top:6; margin-bottom:6"">... " 
    outline = outline & "then select a name revealed below:</p>"		
    outline = outline & "<center><div id=""ajaxplayerlist"" style=""width:300"">"
    
    outline = outline & "<p><img border=""0"" src=""images/dummbar_0.gif"" height=""300"" width=""1""></p>"
    
    outline = outline & "</div>"
    
    response.write(outline)

end if

conn.close
%>	
	
	</td>
    
    </tr>
    
    </table>




</center><br>
</div>

<!--#include file="base_code.htm"-->
</body>

</html>