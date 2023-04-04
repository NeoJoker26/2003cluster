<%@ Language=VBScript %> <% Option Explicit %>
<%
Dim conn,sql,rs
Dim years,part3940
years = Request.QueryString("years")
part3940 = Request.QueryString("part")
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

if years = "" then
	sql = "select max(years) as years "
	sql = sql & "from season " 

	rs.open sql,conn,1,2
	years = rs.Fields("years")
	rs.close
end if
%>

<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>GoS-DB Season</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
#container {width:980px; font-family:verdana,arial,helvetica,sans-serif; font-size: 11px; text-align:left;}

#appearances td { font-family: verdana,arial,helvetica,sans-serif; line-height=1; text-align: center; font-size: 9px; padding: 0;}
#appearances th { font-family: arial,helvetica,sans-serif; line-height=1; vertical-align: top; text-align: center; font-size: 9px; padding: 1px 0;}
#appearances td.head0 { padding: 8px 0 6px 0; font-size: 12px; font-weight: bold; vertical-align: middle;}
#appearances td.head1 { padding: 8px 0 6px 20px; font-size: 11px; text-align: right; vertical-align: middle;}
#appearances .name { font-size: 11px; text-align: left; padding: 0 4px; }
#appearances .num { text-align: right; padding: 0 2px; }
#appearances .cuphead { font-size: 7px; }

#leaguetable {float:right; width:280px; margin:0; padding:0}
#leaguetable td, #leaguetable td, #leaguetable p {font-size: 10px; margin:0; padding: 0 2; text-align: right; white-space: nowrap;}
#leaguetable td:first-child {text-align: left;}

th.match {max-width: 14px}

.hover {background-color: #c8e0c7; cursor: pointer;}

.nohover a:hover {background-color: transparent;}

.apptable {
	border-collapse: collapse;
	width:980px; 
}
.apptable th, .apptable td {
	border: 1px solid #c0c0c0;
}
.restable {
	border-collapse: collapse;
	width:980px; 
	margin:12px auto;
}
.restable th, .restable td {
	border: 1px solid #c0c0c0;
}
-->
</style>

<script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script>

<script language="javascript">

$(document).ready(function(){
	
	$('#appearances').load('gosdb-season_appearances.asp?years=<%response.write(years)%>');
	
	$('#results').load('gosdb-season_results.asp?years=<%response.write(years)%>&part=<%response.write(part3940)%>');
	
	$('#container').on('mouseenter','.match,.name,.sort', function(){
    	$(this).addClass('hover');
   	});
 	
	$('#container').on('mouseleave','.match,.name,.sort', function(){
    	$(this).removeClass('hover');
   	});
   	
   	$('#appearances').on('click','.sort', function(){
   		var sortid = $(this).attr("id");
   		sortid = sortid.substring(4,5);
   		$('#appearances').load('gosdb-season_appearances.asp?years=<%response.write(years)%>&sort=' + sortid);
    });
	  	
   	$('#results').on('click','.sort', function(){
   		var sortid = $(this).attr("id");
   		sortid = sortid.substring(4,5);
   		$('#results').load('gosdb-season_results.asp?years=<%response.write(years)%>&part=<%response.write(part3940)%>&sort=' + sortid);
    });
    
});

function Toggle(item1,item2) {
   obj1=document.getElementById(item1);
   obj2=document.getElementById(item2);
   obj1.style.display="none";
   obj2.style.display="block";
}

</script>

</head>

<body><!--#include file="top_code.htm"-->
<%
Server.ScriptTimeout = 30 

Dim f, fs, photopage, playername, fromname, loanline, loanhold, tierstyle(3), startage, subage
Dim seasonarray(150), firstseason_i, lastseason_i, thisseason_i, maxseason_i, season_left, season_right, heading_done
Dim rsgoals, rslineup, rssubbedsubs, i, j, list, outline, teamline1, teamline2, teamsurnames, tab, tabsuf, datehold, att, goalscorer, thisname, cellclass
Dim sub_surname, sub_forename, sub_initials, sub_playerid, tagno, agecount, totage, orderby, limitdate

tab = Request.QueryString("tab")


Set fs=Server.CreateObject("Scripting.FileSystemObject")

'Prepare photo page name
if years = "1999-2000" then
	photopage = "teampic99-2000.asp"
	elseif left(years,1) = "2" then 
		photopage = "teampic" & left(years,4) & "-" & mid(years,8,2) & ".asp"
	else
		photopage = "teampic" & mid(years,3,2) & "-" & mid(years,8,2) & ".asp"
end if
%>

<div id="container">

    <p style="padding: 6px 0 6px; text-align: center; font-size:11px">
    <%
    sql = "select years "
    sql = sql & "from season  "
	sql = sql & "order by years "

	rs.open sql,conn,1,2
	
	outline = ""
	i = 0
	Do While Not rs.EOF
		seasonarray(i) = rs.Fields("years")
		if seasonarray(i) = years then thisseason_i = i
		i = i + 1
		rs.MoveNext
	Loop
	rs.close
	
	maxseason_i = i-1
	
	if thisseason_i > 4 then 
		firstseason_i = thisseason_i - 5
		season_left = "<a href=""gosdb-season.asp?years=" & seasonarray(firstseason_i) & """><u><<--</u></a> "
	  else 
		firstseason_i = 0
		lastseason_i = 10
		season_left = ""
	end if
	if thisseason_i < maxseason_i - 5 then 
		if thisseason_i > 4 then lastseason_i = thisseason_i + 5
		season_right = " <a href=""gosdb-season.asp?years=" & seasonarray(lastseason_i) & """><u>-->></u></a>"		
	  else 
		lastseason_i = maxseason_i
		firstseason_i = maxseason_i - 11
		season_right = ""
	end if
	
	outline = outline & season_left
	
	for i = firstseason_i to lastseason_i
		if seasonarray(i) = years then
			outline = outline & "<b>" & seasonarray(i) & "</b>; "
		  else
			outline = outline & "<a href=""gosdb-season.asp?years=" & seasonarray(i) & """><u>" & seasonarray(i) & "</u></a>; "
		end if
	next
	
	outline = left(outline,len(outline)-2)  	'remove last comma and space
	outline = outline & season_right
    
    response.write(outline)
    %>
    
	</p>
	
	<div id="top-section" style="overflow:auto;">
	
	<div style="float:left; margin:0; padding:0; width:288px;">
	<div>
    <a href="gosdb.asp"><img style="float:left; border:0" src="images/gosdb-small.jpg" ></a>
    <p style="margin:0; text-align:center; font-size: 13px; font-weight:700; color:#202020">SEASON</p>
    <p style="margin:2px 0;  text-align:center; font-size: 13px; font-weight:700; color:green"><% response.write(years) %></p>
    <p style="margin:6px 0 0; text-align:center; font-weight:700;"><a href="gosdb-seasons.asp">Back to<br>all seasons</a></p> 
    </div>
    
    <%
   	if part3940 = "D2" then
  		limitdate = " and date <= '1939-09-02' " 
   	  elseif part3940 = "SWRL" then
  		limitdate = " and date > '1939-09-02' " 
   	  else
	 	limitdate = "" 	
  	end if

'Prepare list of loanees for later

	sql = "select distinct c.player_id_spell1, surname, forename, initials, came_from, name_then_short "
    sql = sql & "from season a join match_player b on date between date_start and date_end "
	sql = sql & " join player c on b.player_id = c.player_id "
	sql = sql & " left outer join opposition d on c.came_from = d.name_then "
	sql = sql & "where years = '" & years & "' "
	sql = sql & limitdate
	sql = sql & "  and (first_game_year = last_game_year or first_game_year+1 = last_game_year)	"
	sql = sql & "  and came_from = went_to "
	sql = sql & "order by surname, forename, initials "
	
	rs.open sql,conn,1,2
	
	heading_done = 0
	Do While Not rs.EOF
		if heading_done = 0 then
   			loanline = "<p style=""margin:4px 0; line-height:140%; vertical-align: bottom;""><span class=""style1boldgreen"">LOANS: </span>"
			heading_done = 1
		end if
		loanhold = loanhold & rs.Fields("surname") & rs.Fields("player_id_spell1") & ","
		if not isnull(rs.Fields("forename")) then
			playername = trim(rs.Fields("forename")) & " " & trim(rs.Fields("surname"))
		  elseif not isnull(rs.Fields("initials")) then
			playername = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname")) 
		  else playername = trim(rs.Fields("surname"))
		end if

		if not isnull(rs.Fields("name_then_short")) then
			fromname = trim(rs.Fields("name_then_short"))
		  else 
		  	fromname = trim(rs.Fields("came_from"))
		end if
		loanline = loanline & playername & " (" & fromname & "), "
		rs.MoveNext
	Loop
	
	rs.close
	if heading_done > 0 then loanline = left(loanline,len(loanline)-2) & ".</p>"  	'replace comma with fullstop
	
'Manager(s)

    sql = "select managers, date_start "
	sql = sql & "from v_managerspell_horiz a join season b "
	sql = sql & "	on isnull(a.to_date,'9999-12-31') >= b.date_start and a.from_date <= b.date_end "
	sql = sql & "where years = '" & years & "' "
	sql = sql & limitdate
	sql = sql & "order by date_start "
	
	rs.open sql,conn,1,2
	
	heading_done = 0
	outline = ""
	Do While Not rs.EOF
		if heading_done = 0 then
			outline = outline & "<p style=""margin:12px 0 0; line-height:140%; vertical-align: bottom;""><span class=""style1boldgreen"">MANAGER: </span>"
			heading_done = 1
		end if	

		outline = outline & rs.Fields("managers") & "; "
		rs.MoveNext
	Loop
	
	rs.close
	if heading_done > 0 then outline = left(outline,len(outline)-2) & ".</p>"  	'replace last ; with fullstop

	response.write(outline)
	

'Player of the season

    sql = "select distinct tier, division, c.player_id_spell1, surname, forename, initials "
	sql = sql & "from season a join player_season b on a.season_no = b.season_no "
	sql = sql & "	join player c on b.player_id_spell1 = c.player_id "
	sql = sql & "where years = '" & years & "' "
	sql = sql & "order by surname, forename, initials "

	rs.open sql,conn,1,2
	
	heading_done = 0
	outline = ""
	Do While Not rs.EOF
		if heading_done = 0 then
			outline = outline & "<p style=""margin:4px 0; line-height:140%; vertical-align: bottom;""><span class=""style1boldgreen"">PLAYER OF THE SEASON: </span>"
			heading_done = 1
		end if	
		if not isnull(rs.Fields("forename")) then
			playername = trim(rs.Fields("forename")) & " " & trim(rs.Fields("surname"))
		  elseif not isnull(rs.Fields("initials")) then
			playername = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname")) 
		  else playername = trim(rs.Fields("surname"))
		end if

		outline = outline & playername & " and "
		rs.MoveNext
	Loop
	
	rs.close
	if heading_done > 0 then outline = left(outline,len(outline)-5) & ".</p>"  	'replace last "and" with fullstop

	response.write(outline)
	
'Debut players
		
    sql = "select distinct c.player_id_spell1, surname, forename, initials "
    sql = sql & "from season a join match_player b on date between date_start and date_end "
	sql = sql & " join player c on b.player_id = c.player_id "
	sql = sql & "where years = '" & years & "' "
	sql = sql & " and not exists ( select * from match_player d join player e on d.player_id = e.player_id"
	sql = sql & "					where e.player_id_spell1 = c.player_id_spell1 "
	sql = sql & "					and d.date < b.date ) "
	sql = sql & limitdate
	sql = sql & "order by surname, forename, initials "

	rs.open sql,conn,1,2
	
	heading_done = 0
	outline = ""
	Do While Not rs.EOF
		if heading_done = 0 then
   			outline = outline & "<p style=""margin:4px 0; line-height:140%; vertical-align: bottom;""><span class=""style1boldgreen"">DEBUTS: </span>"
   			heading_done = 1
		end if	
		if not isnull(rs.Fields("forename")) then
			playername = trim(rs.Fields("forename")) & " " & trim(rs.Fields("surname"))
		  elseif not isnull(rs.Fields("initials")) then
			playername = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname")) 
		  else playername = trim(rs.Fields("surname"))
		end if

		'exclude if already listed as a loan
		if instr(loanhold, rs.Fields("surname") & rs.Fields("player_id_spell1")) = 0 then
			outline = outline & playername & ", "
		end if	
		rs.MoveNext
	Loop
	
	rs.close
	if heading_done > 0 then outline = left(outline,len(outline)-2) & ".</p>"  	'replace last comma with fullstop

	response.write(outline)
	
'Final game players

    sql = "select distinct c.player_id_spell1, surname, forename, initials "
    sql = sql & "from season a join match_player b on date between date_start and date_end "
	sql = sql & " join player c on b.player_id = c.player_id "
	sql = sql & "where years = '" & years & "' "
	sql = sql & " and not exists ( select * from match_player d join player e on d.player_id = e.player_id"
	sql = sql & "					where e.player_id_spell1 = c.player_id_spell1 "
	sql = sql & "					and d.date > b.date ) "
	sql = sql & " and last_game_year < 9999 "
	sql = sql & limitdate
	sql = sql & "order by surname, forename, initials "

	rs.open sql,conn,1,2

	heading_done = 0
	outline = ""
	Do While Not rs.EOF
		if heading_done = 0 then
   			outline = outline & "<p style=""margin:4px 0; line-height:140%; vertical-align: bottom;""><span class=""style1boldgreen"">FINAL GAMES: </span>"
   			heading_done = 1
		end if	
		if not isnull(rs.Fields("forename")) then
			playername = trim(rs.Fields("forename")) & " " & trim(rs.Fields("surname"))
		  elseif not isnull(rs.Fields("initials")) then
			playername = trim(rs.Fields("initials")) & " " & trim(rs.Fields("surname")) 
		  else playername = trim(rs.Fields("surname"))
		end if

		'exclude if already listed as a loan
		if instr(loanhold, rs.Fields("surname") & rs.Fields("player_id_spell1")) = 0 then
			outline = outline & playername & ", "
		end if	
		rs.MoveNext
	Loop
	
	rs.close
	if heading_done > 0 then outline = left(outline,len(outline)-2) & ".</p>"  	'replace last comma with fullstop

	response.write(outline)
	
' Output pre-prepared loan players

	response.write(loanline)
	
    %>
    </div>
    
    <div style="float:right; width:692px;"> 
    
    <div style="float:left; width:400px; text-align:center;"> 
    <%
 
 	for i = 0 to 3
		tierstyle(i) = "style=""color:#d0d0d0"""
	next
	
    sql = "select tier, division "
	sql = sql & "from season  "
	sql = sql & "where years = '" & years & "' "
	rs.open sql,conn,1,2
	
	outline = "<p style=""margin:0; font-size: 14px; font-weight:700; color:#000000"">" & rs.Fields("division") & "</p>"
	
	if not isnull(rs.Fields("tier")) then 
		outline = outline & "<p style=""margin: 0; font-size: 13px; font-weight:700;"">Tier "
	
		tierstyle(rs.Fields("tier")-1) = "style=""color:green"""
	
		for i = 0 to 3
			outline = outline & "<span " & tierstyle(i) & ">" & i+1 & " <span>"
		next
		outline = outline & "</p>"
	end if
	rs.close

    if fs.FileExists(Server.MapPath("gosdb/photos/" & years & ".jpg")) then
	    outline = outline & "<a href=""" & photopage & """>"
        outline = outline & "<img style=""margin-top: 6px"" src=""gosdb/photos/" & years & ".jpg"">"
    	outline = outline & "<p style=""margin:0; text-align:center;"">Click for larger team photo</p></a>"
      end if
    response.write(outline)
    %>
    </div>
      
	<table id="leaguetable">	
	
	<%
	outline = ""
	if years = "1903-1904" or years = "1904-1905" or years = "1905-1906" or years = "1906-1907" or years = "1907-1908" or years = "1908-1909" then 
		if tab = "WL" then
			tabsuf = "WL"
			outline = outline & "<tr><td colspan=""8"" align=""center""><b>WESTERN LEAGUE</b> | <a href=""http://www.greensonscreen.co.uk/gosdb-season.asp?years=" & years & "&list=" & list & "&tab=SL""><u>Southern League</u></a></td></tr>"
	  	  else 
			tabsuf = ""
			outline = outline & "<tr><td colspan=""8"" align=""center""><b>SOUTHERN LEAGUE</b> | <a href=""http://www.greensonscreen.co.uk/gosdb-season.asp?years=" & years & "&list=" & list & "&tab=WL""><u>Western League</u></a></td></tr>"
		end if
	end if
	if years = "1939-1940" then 
		if tab = "SWRL" then
			tabsuf = "SWRL"
			outline = outline & "<tr><td colspan=""8"" align=""center""><b>SW REGIONAL LEAGUE</b> | <a href=""http://www.greensonscreen.co.uk/gosdb-season.asp?years=" & years & "&list=" & list & "&tab=""><u>Division 2 (Abnd)</u></a></td></tr>"
	  	  else 
			tabsuf = ""
			outline = outline & "<tr><td colspan=""8"" align=""center""><b>DIVISION 2</b> | <a href=""http://www.greensonscreen.co.uk/gosdb-season.asp?years=" & years & "&list=" & list & "&tab=SWRL""><u>SW Regional League</u></a></td></tr>"
		end if
	end if

	if fs.FileExists(Server.MapPath("gosdb/gosdb-table" & years & ".txt")) then
		Set f=fs.OpenTextFile(Server.MapPath("gosdb/gosdb-table" & years & tabsuf & ".txt"),1)
		Do While Not f.AtEndOfStream
			outline = outline & f.Readline
		Loop
		f.close
	end if
	response.write(outline)
	%>
		
	</table>
	
	</div>
	
	</div>


<div id="appearances" style="clear:both; margin:24px auto; text-align: center">Fetching appearance data</div>

<div id="results" style="clear:both; margin:24px auto; text-align: center">Fetching results</div>

<!--#include file="base_code.htm"-->

</body>

</html>