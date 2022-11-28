
<%@ Language=VBScript %>
<% Option Explicit %>

<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>Greens on Screen Database</title>

<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
<style>
<!--

.rowhlt { }

#gottable1 td {text-align:left; margin: 0; padding: 0 4 0 4; font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 11px; }
#gottable1 p {margin: 3 0 3 0; padding: 0; font-family: verdana,arial,sans-serif; font-size: 11px; } 
#gottable1 .right {text-align: right; } 
#gottable1 .tah {font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; } 

#gottable2 td {text-align:left; margin: 0; padding: 0 4 0 4; font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 11px; } 
#gottable2 p {margin: 3 0 3 0; padding: 0; font-family: verdana,arial,sans-serif; font-size: 11px; } 
#gottable2 .right {text-align: right; }
#gottable2 .tah {font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; } 


#table1 tr { border: 0px; }
#table1 tr.border { border: 1px solid #202020; background-color: #e0f0e0; }
#table1 td {padding: 2 8;}
#table1 td.bold {font-weight: 700;}
#table1 td.a {}
#table1 td.b {}
#table1 td.c {}
#table1 td.d {}
#table1 td.t {}
#table1 td.head {padding-bottom: 3; border-bottom: 1px solid #c0c0c0; font-size: 11px; font-weight: bold; color:#006e32; }

-->

</style>

<%
Dim conn,sql,rs,rs1,lastmon,lastday,thisdate,i,j,datearray(11,31),urldate
Dim rslineup,rsgoals, displaydate, years, yearslast, opposition, outline, outlinematch, called, fullteam, crowd, venue, work1, tagno, selday, selmon, seldate, seldateshort, totwins, totmatches
Dim latestdates(5,1), counts(8,2),tableview, heading1, heading2, headtext, competition, check1, check2, check3, check4, clickon, boldclass, borderclass, season_years1, season_years2, season_text, restrictions

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

sql = "select distinct day, "
sql = sql & "case when month = 'Jan' then 1 when month = 'Feb' then 2 when month = 'Mar' then 3 "
sql = sql & "     when month = 'Apr' then 4 when month = 'May' then 5 when month = 'Jun' then 6 "
sql = sql & "     when month = 'Jul' then 7 when month = 'Aug' then 8 when month = 'Sep' then 9 "
sql = sql & "     when month = 'Oct' then 10 when month = 'Nov' then 11 when month = 'Dec' then 12 "
sql = sql & "end as monthno "
sql = sql & "from onthisday " 
sql = sql & "order by monthno, day " 
rs.open sql,conn,1,2

i = -1
lastmon = ""
Do While Not rs.EOF
	if rs.Fields("monthno") <> lastmon then
		lastmon = rs.Fields("monthno")
		i = i + 1
		j = 1
		datearray(i,0) = rs.Fields("monthno")
	end if
	datearray(i,j) = rs.Fields("day")
	j = j + 1
	rs.MoveNext
Loop
rs.close 
%>


<script language="javascript">

function setOptions(chosen) { 
 var Days=document.form1.matday;
 Days.options.length=0; 
 
 <%
 for i = 0 to 11
  	if datearray(i,0) = "" then exit for
 	response.write("if(chosen==" & datearray(i,0) & ") { ")
 	for j = 1 to 31
 		if datearray(i,j) = "" then exit for
 		response.write("Days.options[Days.options.length] = new Option('" & datearray(i,j) & "','" & datearray(i,j) & "'); ")
 	next
 	response.write("} ")
 next
 %>		
}

function GetTable(side) { 

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
// triggered() function handles the events 
xmlhttp.onreadystatechange = triggered; 

// open takes in the HTTP method and url.
//document.body.style.cursor='wait';        
var url="gosdb-getdatespagedetails.asp";
url=url+"?side="+side;
url=url+"&sid="+Math.random();

xmlhttp.open("GET", url, true); 
 
// send the request. if this is a POST request we would have 
// sent post variables: send("name=aleem&gender=male) 
// Moz is fine with just send(); but 
// IE expects a value here, hence we do send(null); 
xmlhttp.send(null);
//document.body.style.cursor='auto';  
} 

function triggered() { 
// if the readyState code is 4 (Completed) 
// and http status is 200 (OK) we go ahead and get the responseText 
// other readyState codes: 
// 0=Uninitialised 1=Loading 2=Loaded 3=Interactive 
if (xmlhttp.readyState == 4) { 
        // xmlhttp.responseText object contains the response.
        var textsplit = xmlhttp.responseText.split("^");
        if (textsplit[0] == "left" || textsplit[0] == "both") { document.getElementById('left').innerHTML = textsplit[1]; } 
        if (textsplit[0] == "right" || textsplit[0] == "both") { document.getElementById('right').innerHTML = textsplit[2]; } 
} 
} 

function Toggle(item,clickon) {
   obj=document.getElementById(item);
   objtr=document.getElementById("tr" + item);
   objdate=document.getElementById("d" + item);
   GetDetails(objdate.innerHTML);

   visible=(obj.style.display!="none")
   key=document.getElementById("x" + item);
   if (visible) {
     obj.style.display="none";
     objtr.style.backgroundColor="";
     key.innerHTML= '[+<span style="font-family:verdana;">' + clickon + '</span>]';
   } else {
      obj.style.display="block";
      objtr.style.backgroundColor="#e0f0e0";
      key.innerHTML='[-<span style="font-family:verdana;">' + clickon + '</span>]';
   }
}

var xmlHttp

function GetDetails(str)
{ 
xmlHttp=GetXmlHttpObject();

if (xmlHttp==null)
  {
  alert ("Sorry, your browser does not support this function.");
  return;
  }
document.body.style.cursor='wait';        
var url="gosdb-getmatchdetails1.asp";
url=url+"?q="+str;
url=url+"&sid="+Math.random();
xmlHttp.onreadystatechange=stateChanged;
xmlHttp.open("GET",url,true);
xmlHttp.send(null);
document.body.style.cursor='auto';   
}

function stateChanged() 
{ 
if (xmlHttp.readyState==4)
   { 
   obj.innerHTML = xmlHttp.responseText;
   }
}

function GetXmlHttpObject()
{
var xmlHttp=null;
try
  {
  // Firefox, Opera 8.0+, Safari
  xmlHttp=new XMLHttpRequest();
  }
catch (e)
  {
  // Internet Explorer
  try
    {
    xmlHttp=new ActiveXObject("Msxml2.XMLHTTP");
    }
  catch (e)
    {
    xmlHttp=new ActiveXObject("Microsoft.XMLHTTP");
    }
  }
return xmlHttp;
}

</script>

</head>

<body onLoad="javascript:GetTable('both','all');
<%'response.write("javascript:Toggle('tag" & eval(onload_tagno) & "','" & eval(onload_clickon) & "');")%>
">
<!--#include file="top_code.htm"-->

<center>

<table border="0" cellpadding="0" cellspacing="0" 
style="border-collapse: collapse; margin-top:20 0;" bordercolor="#111111" 
width="984">
<tr>
<td width="290px" style="text-align: left" valign="top">
<div style="width:260;">
<p style="text-align: center; margin-top:0; margin-bottom:3">
<a href="gosdb.asp"><font color="#404040"><img border="0" src="images/gosdb-small.jpg" align="left"></font></a><font 
color="#404040"><b><font style="font-size: 15px">Search by<br>
Date</font></b></font><p style="text-align: center; margin-top:0; margin-bottom:0">
<b>
<a href="gosdb.asp">Back to<br>GoS-DB Hub</a></b>
</div>
    
<div id="left" align="left" style="margin-top:12px">Retrieving data ...</div>
	
</td>
<td align="center" valign="top" style="text-align: center; padding: 0 10px;">
<p style="margin-top: 0; margin-bottom: 3; text-align:center">
<span style="font-size: 18px"><font color="#006E32">
PAFC BY DATE</font></p>
</span>
	
<%
urldate = Request.QueryString("date")
selday = Request.Form("matday")
selmon = Request.Form("matmonth")

if urldate > "" then
		work1 = split(urldate," ")
		for i = 1 to 12
			if MonthName(i) = trim(work1(1)) then
				selmon = i
				exit for
			end if
		next
		if selmon > 0 and selmon < 13 then 
			selday = work1(0)
			seldate = selday & " " & MonthName(selmon)
			'if a year was included in the urldate, rebuild in shortened form of month for later use
			if ubound(work1) = 2 then
				urldate = selday & " " & Monthname(selmon,True) & " " & work1(2)
			end if
			response.write("<p style=""margin-top: 0; margin-bottom: 6; text-align:center; font-size: 16px""><b>" & seldate & "</b></p>")
		end if
	elseif selday = "" then
		'if no passed parameters, must be first time in; use today's date
		selday = DatePart("d",Date)
		selmon = DatePart("m",Date)
		seldate = selday & " " & MonthName(selmon)
		seldateshort = selday & " " & MonthName(selmon,1)
		response.write("<p style=""margin-top: 0; margin-bottom: 6; text-align:center; font-size: 14px""><b>Today: " & seldate & "</b></p>")
  	else
  		seldate = selday & " " & MonthName(selmon)
  		response.write("<p style=""margin-top: 0; margin-bottom: 6; text-align:center; font-size: 14px""><b>" & ucase(seldate) & "</b></p>")
end if	
%>

	<p style="margin: 12px 0 3px; font-size: 12px;">Select a new date ...</p>
    <form style="padding: 0; margin: 0;" 
    action="gosdb-dates.asp" method="post" name="form1">
	<p style="margin-top: 0; margin-bottom: 0">
	<select name="matmonth" style="font-size: 12px" onchange="setOptions(document.form1.matmonth.options[document.form1.matmonth.selectedIndex].value);">
     	
	<%
	outline = ""
	for i = 1 to 12
		outline = outline & " <option value=""" & i & """"
		if i = cInt(selmon) then outline = outline & " selected"
		outline = outline & ">" & MonthName(i) & "</option>"
	next
	response.write(outline)
	%>
	</select>
	 
	<select name="matday" style="font-size: 12px">
	<%
	outline = ""
	for i = 1 to 31
		outline = outline & " <option value=""" & i & """"
		if i = cInt(selday) then outline = outline & " selected"
		outline = outline & ">" & i & "</option>"
	next
	response.write(outline)
	%> 
	
  	</select> 
  	</p>
  		
 	<p style="text-align: center; margin-top:0; margin-bottom:0">
    <input type="submit" 
    style="width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 11px; margin-left:0; margin-right:0; margin-top:3; margin-bottom:0; padding-left:5; padding-right:5; padding-top:1; padding-bottom:1" 
    value="Display selected date" name="B1"></p> 
 	</form>

<p style="font-weight: bold; margin: 24px auto 0;">ON THIS DAY</p>

<%
Dim texthold

sql = "select year, fact "
sql = sql & "from onthisday "  
sql = sql & "where month = '" & left(monthname(selmon),3) & "' "
sql = sql & "  and day = " & selday & " " 
sql = sql & "order by seqno "
rs.open sql,conn,1,2
		
Do While Not rs.EOF	
		
	if rs.Fields("year") > "" then
		texthold = "<b>" & rs.Fields("year") & ":</b> " & rs.Fields("fact")
	  else 
		texthold = rs.Fields("fact")
	end if
			
	if instr(texthold,"^^") > 0 then texthold = replace(texthold, "^^", year(date) - rs.Fields("year"))
	 
	response.write("<p style=""text-align:left; margin: 0 6px; padding: 4px 0;"">" & texthold & "</p>")
			
	rs.MoveNext
Loop
rs.close
%>

<p style="font-weight: bold; margin: 12px auto 0;">BORN THIS DAY</p>

<%
Dim games, goals, penpichold
texthold = ""

sql = "select a.player_id, a.forename, a.surname, year(a.dob) as year, a.first_game_year, max(b.last_game_year) as last_game_year, left(a.penpic,160) as penpic, a.prime_photo "
sql = sql & "from player a left outer join player b on a.player_id = b.player_id_spell1 "  
sql = sql & "where month(a.dob) = " & selmon & " "
sql = sql & "  and day(a.dob) = " & selday & " " 
sql = sql & "  and a.spell = 1 "
sql = sql & "group by a.player_id, a.forename, a.surname, a.dob, a.first_game_year, a.penpic, a.prime_photo " 
sql = sql & "order by a.dob "
rs.open sql,conn,1,2
		
If rs.EOF then
		
	texthold = "<p style=""margin: 6px 0;"">We know of no first-team players born on this day.</p>"

  else
		
   	Do While Not rs.EOF
   		
   		sql = "with cte as ( "
		sql = sql & "select count(*) as starts, 0 as subs, 0 as goals "
		sql = sql & "from player a join match_player b on a.player_id = b.player_id "
		sql = sql & "where player_id_spell1 = " & rs.Fields("player_id")
		sql = sql & "  and startpos > 0 "
		sql = sql & "union all "
		sql = sql & "select 0, count(*), 0 "
		sql = sql & "from player a join match_player b on a.player_id = b.player_id "
		sql = sql & "where player_id_spell1 = " & rs.Fields("player_id")
		sql = sql & "  and startpos = 0 "
		sql = sql & "union all "
		sql = sql & "select 0, 0, count(*) "
		sql = sql & "from player a join match_goal b on a.player_id = b.player_id "
		sql = sql & "where player_id_spell1 = " & rs.Fields("player_id")
		sql = sql & ") "
		sql = sql & "select sum(starts) + sum(subs) as games, sum(goals) as goals "
		sql = sql & "from cte " 
	
		rs1.open sql,conn,1,2
		games = rs1.Fields("games")
		if games = 1 then 
			games = "1 game"
		  else games = games & " games"
		end if
		goals = rs1.Fields("goals")
		if goals = 0 then 
			goals = "no goals"
		  elseif goals = 1 then 
			goals = "1 goal"
		  else goals = goals & " goals"
		end if

		rs1.close
    				

		texthold = texthold & "<p style=""text-align:left; margin: 6px 3px 0 6px;""><b>" & rs.Fields("year") & ":</b> <a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id") & """>" & rs.Fields("forename") & " " & rs.Fields("surname") & "</a> "  
		texthold = texthold & "(" & games & ", " & goals 
		  	
		if rs.Fields("last_game_year") = 9999 then
		  	texthold = texthold & " so far)"
		  elseif rs.Fields("first_game_year") = rs.Fields("last_game_year") then 
		  	texthold = texthold & " in " & rs.Fields("first_game_year") & ")" 	
		  else texthold = texthold & " from " & rs.Fields("first_game_year") & " to " & rs.Fields("last_game_year") & ")"
		end if
		  	
		texthold = texthold & "</p><p style=""margin: 2 3 0 6;"">" & penpichold & "</p>"
			
	rs.MoveNext
	Loop
	
end if
		
rs.close

response.write(texthold)
%>

<p style="font-weight: bold; margin: 20px auto 0;">MATCH RECORD</p>

<%
tagno = 1
outlinematch = ""
for i = 0 to ubound(counts,1)
	for j = 0 to ubound(counts,2)
		counts(i,j) = 0
	next
next

sql = "select opposition, opposition_qual, date, "
sql = sql & "case homeaway when 'H' then 'H' else 'A' end as homeawayHA"
sql = sql & ", shortcomp, subcomp, goalsfor, goalsagainst, attendance, notes, ground_name "
sql = sql & "from v_match_all " 
sql = sql & " join opposition on opposition = name_then "
sql = sql & " left outer join venue on name_then = club_name_then and date between first_game and last_game "
sql = sql & "where day(date) = " & selday & " and month(date) = " & selmon & " " 
sql = sql & "order by date "

rs.open sql,conn,1,2

Do While Not rs.EOF

	'Accumulate stats
	if rs.Fields("homeawayHA") = "H" then 
		j = 0
		else j = 1
	end if 
	
	counts(0,j) = counts(0,j) + 1  'played
	if rs.Fields("goalsfor") > rs.Fields("goalsagainst") then 
		counts(1,j) = counts(1,j) + 1 'wins
		if latestdates(1,j) < rs.Fields("date") then latestdates(1,j) = rs.Fields("date")
	end if
	if rs.Fields("goalsfor") = rs.Fields("goalsagainst") then 
		counts(2,j) = counts(2,j) + 1 'draws
		if latestdates(2,j) < rs.Fields("date") then latestdates(2,j) = rs.Fields("date")
	end if
	if rs.Fields("goalsfor") < rs.Fields("goalsagainst") then 
		counts(3,j) = counts(3,j) + 1 'defeats
		if latestdates(3,j) < rs.Fields("date") then latestdates(3,j) = rs.Fields("date")
	end if
	counts(4,j) = counts(4,j) + rs.Fields("goalsfor")  'goals for
	counts(5,j) = counts(5,j) + rs.Fields("goalsagainst")  'goals against
	if IsNumeric(rs.Fields("attendance")) then 
		counts(6,j) = counts(6,j) + rs.Fields("attendance")  'attendance
		counts(7,j) = counts(7,j) + 1  'games with attendance
	end if	

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

	clickon = ""
	if rs.Fields("notes") > "" then clickon = clickon & "N"

	venue = rs.Fields("ground_name")
	
	boldclass = ""
	borderclass = ""
	
	'if a full date (inc. year) came in, it's for a specific date to be highlit
	if displaydate = urldate then 
		boldclass = " bold"
		borderclass = " class=""border"""
	end if

	outlinematch  = outlinematch & "<tr" & borderclass & " id=""trtag" & tagno & """ onmouseover=""this.className = 'rowhlt';"" onmouseout=""this.className = '';""><td class=""a" & boldclass & """ nowrap=""nowrap""><a style=""font-family:courier;"" id=""xtag" & tagno & """ href=""javascript:Toggle('tag" & tagno & "','" & clickon & "');"">[+<span style=""font-family:verdana;"">" & clickon & "</span>]</a></td>" 
	outlinematch  = outlinematch & "<td class=""a" & boldclass & """ nowrap=""nowrap""><a href=""gosdb-match.asp?date=" & rs.Fields("date") & """>" & left(displaydate,6) & " <b>" & right(displaydate,4) & "</b></a><span id=""dtag" & tagno & """ style=""display:none;"">" & rs.Fields("date") & "</span></td>"	
	if rs.Fields("homeawayHA") = "H" then
		venue = "Home Park"
		if displaydate = "18 Mar 1961" then venue = "Plainmoor"
		outlinematch  = outlinematch & "<td class=""b" & boldclass & """ align=""right"">" & "Argyle" & "</td>"   
		outlinematch  = outlinematch & "<td class=""c" & boldclass & """ nowrap=""nowrap"">" & rs.Fields("goalsfor") & " - " & rs.Fields("goalsagainst") & "</td>" 
		outlinematch  = outlinematch & "<td class=""d" & boldclass & """>" & opposition & " " & rs.Fields("opposition_qual") & "</td>"
		outlinematch  = outlinematch & "<td class=""a" & boldclass & """ nowrap=""nowrap"">" & rs.Fields("shortcomp") & " " & rs.Fields("subcomp") & "</td>"
		outlinematch  = outlinematch & "<td class=""a" & boldclass & """ align=""right"">" & crowd & "</td>" 
		outlinematch  = outlinematch & "<td class=""a" & boldclass & """>" & venue & "</td>"   
		outlinematch  = outlinematch & "</tr>"
	  else  
		outlinematch  = outlinematch & "<td class=""b" & boldclass & """ align=""right"">" & opposition & " " & rs.Fields("opposition_qual") & "</td>"  
		outlinematch  = outlinematch & "<td class=""c" & boldclass & """ nowrap=""nowrap"">" & rs.Fields("goalsagainst") & " - " & rs.Fields("goalsfor") & "</td>"
		outlinematch  = outlinematch & "<td class=""d" & boldclass & """>" & "Argyle" & "</td>"	
		outlinematch  = outlinematch & "<td class=""a" & boldclass & """ nowrap=""nowrap"">" & rs.Fields("shortcomp") & " " & rs.Fields("subcomp") & "</td>"
		outlinematch  = outlinematch & "<td class=""a" & boldclass & """ align=""right"">" & crowd & "</td>" 
		outlinematch  = outlinematch & "<td class=""a" & boldclass & """>" & venue & "</td>" 
		outlinematch  = outlinematch & "</tr>"		
	end if
	
	outlinematch  = outlinematch & "<tr><td class=""a""></td><td  class=""a"" colspan=""8"" width=""500px"" style=""font-size:10px;""><span id=""tag" & tagno & """ style=""display:none; padding: 1 6 4 6; background-color: #e0f0e0; ""></span></td></tr>"	  
	
	tagno = tagno + 1
	
	rs.MoveNext 
Loop 

rs.close

if counts(0,0) + counts(0,1) > 0 then

	'Games have been played on this day - display summary figures

	sql = "select count(date) as wins "
	sql = sql & "from match " 
	sql = sql & "where goalsfor > goalsagainst "
	rs.open sql,conn,1,2

	totwins = rs.Fields("wins")

	rs.close

	sql = "select count(date) as totmatches, count(date)/ count(distinct cast(day(date) as varchar) + ' ' + cast(month(date) as varchar)) as matchesperday  "
	sql = sql & "from match " 
	rs.open sql,conn,1,2

	totmatches = rs.Fields("totmatches")
	outline = ""
	outline = outline & "<p style=""margin: 6px 0;"">" & counts(0,0) + counts(0,1) & " games played (norm: " & rs.Fields("matchesperday") & ")</p>"
	outline = outline & "<p style=""margin: 6px 0 12px;"">" & round(100*(counts(1,0) + counts(1,1))/(counts(0,0) + counts(0,1))) & "% of games won (norm: " & round(100*totwins/totmatches) & "%)</p>"

	rs.close 
	conn.close

	outline = outline & "<table style=""border-collapse: collapse;"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""250px"" align=""center"" >"

	outline = outline & "<tr>"
	outline = outline & "<td>" & seldateshort & "</td>"
	outline = outline & "<td>P</td>"
	outline = outline & "<td>W</td>"
	outline = outline & "<td>D</td>"
	outline = outline & "<td>L</td>"
	outline = outline & "<td>F</td>"
	outline = outline & "<td>A</td>"
	outline = outline & "<td align=""right"">Avg Att</td>"
	outline = outline & "</tr><tr>"
	outline = outline & "<td>Home</td>"
	outline = outline & "<td>" & counts(0,0) & "</td>"
	outline = outline & "<td>" & counts(1,0) & "</td>"
	outline = outline & "<td>" & counts(2,0) & "</td>"
	outline = outline & "<td>" & counts(3,0) & "</td>"
	outline = outline & "<td>" & counts(4,0) & "</td>"
	outline = outline & "<td>" & counts(5,0) & "</td>"
	if counts(7,0) > 0 then
		outline = outline & "<td align=""right"">" & FormatNumber(Round(counts(6,0)/counts(7,0)),0,0,0,-1) & "</td>"
			else outline = outline & "<td align=""right"">0</td>"
	end if
	outline = outline & "</tr><tr>"
	outline = outline & "<td>Away</td>"
	outline = outline & "<td>" & counts(0,1) & "</td>"
	outline = outline & "<td>" & counts(1,1) & "</td>"
	outline = outline & "<td>" & counts(2,1) & "</td>"
	outline = outline & "<td>" & counts(3,1) & "</td>"
	outline = outline & "<td>" & counts(4,1) & "</td>"
	outline = outline & "<td>" & counts(5,1) & "</td>"
	if counts(7,1) > 0 then
		outline = outline & "<td align=""right"">" & FormatNumber(Round(counts(6,1)/counts(7,1)),0,0,0,-1) & "</td>"
			else outline = outline & "<td align=""right"">0</td>"
	end if
	outline = outline & "</tr><tr>"
	outline = outline & "<td>Both</td>"
	outline = outline & "<td>" & counts(0,0) + counts(0,1) & "</td>"
	outline = outline & "<td>" & counts(1,0) + counts(1,1) & "</td>"
	outline = outline & "<td>" & counts(2,0) + counts(2,1) & "</td>"
	outline = outline & "<td>" & counts(3,0) + counts(3,1) & "</td>"
	outline = outline & "<td>" & counts(4,0) + counts(4,1) & "</td>"
	outline = outline & "<td>" & counts(5,0) + counts(5,1) & "</td>"
	if counts(7,0)+counts(7,1) > 0 then
		outline = outline & "<td align=""right"">" & FormatNumber(Round((counts(6,0)+counts(6,1))/(counts(7,0)+counts(7,1))),0,0,0,-1) & "</td>"
		else outline = outline & "<td align=""right"">0</td>"
	end if
	outline = outline & "</tr></table>"
	
  else
  
  	outline = "<p style=""margin: 6px 0;"">No games have been played on this day</p>"
  
end if

response.write(outline)
%>
	
      </td>
      <td width="290px" valign="top" style="text-align: center">
      <p style="margin-bottom:0; margin-right:2px; margin-left:20px; margin-top:3px" 
      align="justify">This page displays events, birthdays and match details for games played on 
      today's date, in all competitions. Select a different day and month for 
      the overall club record and individual match details for that day.</p>

	  <div id="right" align="right" style="margin-top:12px">Retrieving data ...</div>
      
      </td>
    </tr>
  	<tr><td colspan="3">
  	 
      <% if	counts(0,0) + counts(0,1) > 0 then response.write("<p style=""margin:12px 0 0; text-align:center"">Below is the full record for the chosen day. Click on [+] for more, or the date for the full match page.</p>") %>
      </p>
  	
  	<center>

	<%
	if counts(0,0) + counts(0,1) > 0 then
		outline = ""
		outline = outline & "<div id=""table1"" style=""margin-top:18px"">"
		outline = outline & "<table style=""border-collapse: collapse; cell-spacing: 0;"">"
		outline = outline & "<tr>"
		outline = outline & "<td class=""head""></td>"
		outline = outline & "<td class=""head"">MATCH DATE</td>"
		outline = outline & "<td class=""head"" align=""center"" colspan=""3"">RESULT</td>"
		outline = outline & "<td class=""head"">COMP.</td>"
		outline = outline & "<td class=""head"">ATT.</td>"
		outline = outline & "<td class=""head"">VENUE</td>"
		outline = outline & "</b></tr>"
		outline = outline & outlinematch
		outline = outline & "</table>"			
		response.write(outline)
	end if
  	%>
  	</td></tr>
  	</tbody>
    </table>
  	</div>
<br>
 
<!--#include file="base_code.htm"-->  		
</body></html>