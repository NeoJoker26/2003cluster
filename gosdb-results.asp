<%@ Language=VBScript %>
<% Option Explicit %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
   <meta http-equiv="Content-Language" content="en-gb">
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
   <meta name="Author" content="Trevor Scallan">
   <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>GoS-DB Results</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
div#table1 tr td { border: 0px none; }
div#table1 td.a {border-left: 1px dotted #c0c0c0; border-right: 1px dotted #c0c0c0; }
div#table1 td.b {border-bottom-style: none; border-right-style: none;}
div#table1 td.c {border-bottom-style: none; border-left-style: none; border-right-style: none;}
div#table1 td.d {border-bottom-style: none; border-left-style: none; border-right: 1px dotted #c0c0c0; }
div#table1 td.t {border-top: 1px solid #c0c0c0; }
div#table1 td.pa {font-size: 10px; color: red;}
div#table1 td.head {border: 0px solid #c0c0c0; padding-bottom: 6px; font-size: 11px; font-weight: bold; color:#006e32; }

-->
   </style>
<script language="javascript">
function HeadToggle(item) {
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

function Toggle(item,clickon) {
   obj=document.getElementById(item);
   objdate=document.getElementById("d" + item);
   GetDetails(objdate.innerHTML);

   visible=(obj.style.display!="none")
   key=document.getElementById("x" + item);
   if (visible) {
     obj.style.display="none";
     key.innerHTML= '[+<span style="font-family:verdana;">' + clickon + '</span>]';
   } else {
      obj.style.display="block";
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
<body>
<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs,rslineup,rsgoals, displaydate, years, yearslast, opposition, outline, outline1, outline2, outlinematch, called, fullteam, crowd, venue, work1, tagno, scoresep
Dim latestdates(5,1), counts(8,2),i,j, tableview, heading1, heading2, headtext, competition, check1, check2, check3, check4, clickon, topclass, season_years1, season_years2, season_text, restrictions, notplayed_qual

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rsgoals = Server.CreateObject("ADODB.Recordset")
Set rslineup = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_read.inc"--><%
%>

<div id=sv>
<%
if called = "opposition" then 
	else
		fullteam = Request.QueryString("team")
end if
competition = Request.QueryString("comp")
season_years1 = Request.QueryString("s1")
season_years2 = Request.QueryString("s2")

check1 = ""
check2 = ""
check3 = ""
check4 = ""
if Request.Form("R1") > "" then competition = Request.Form("R1")
check1 = "checked"
tableview = "v_match_all"
notplayed_qual = ""
heading1 = "All Competitions"
headtext = "all competitive first-team games since the club turned professional, including the Southern League [1903-1920]; the Western League, a mid-week league including south-east clubs [1903-08]; the Football League [1920-39, 1946-present]; the South West Regional League [1939-40]; the Football League South [1945-46]; and all Cup competitions (see the Cup option for details)." 

select case competition
	case "FLG"
		check1 = ""
		check2 = "checked"
		tableview = "v_match_FL"
		notplayed_qual = " and LFC = 'F' "
		heading1 = "Football League"
		headtext = "all matches in tier 2 [Div 2 to 1991, Div 1 to 2003 and the Championship]; tier 3 [Div 3 South to 1958, Div 3 to 1991, Div 2 to 2003]; and tier 4 [Div 3, 1992-2003], all from 1920 to 1939 and 1946 to the present day." 
	case "LGS"
		check1 = ""
		check3 = "checked"
		tableview = "v_match_all_league"
		notplayed_qual = " and LFC <> 'C' "
		heading1 = "All Leagues"
		headtext = "all matches in all league competitions, including the Southern league [1903-20]; the Western League, a mid-week league including south-east clubs [1903-08]; the Football League [1920-39, 1946-present]; the South West Regional League [1939-40]; and the Football League South [1945-46]."
	case "CUP"
		check1 = ""
		check4 = "checked"
		tableview = "v_match_cups"
		notplayed_qual = " and LFC = 'C' "
		heading1 = "All Cups"
		headtext = "all matches in all knock-out competitions, including the FA Cup [1903-present]; the Football League War Cup [1940]; the Football League Cup (various sponsors) [1960-present]; the Full Members Cup [1986] (a competition for tiers 1 and 2, also known as the Simod Cup [1987-88] and Zenith Data Systems Cup [1989-91]); the Football League Trophy (a generic name for a competition for tiers 3 and 4, including the Associate Members Cup [1984], the Freight Rovers Trophy [1985-86], the Autoglass Trophy[1993], the Auto Windshields Shield [1994-2000], the LDV Vans Trophy [2000-03] and the Johnstone's Paint Trophy [from 2010]); and official pre-season competitions (the Watney Cup [1973], the Anglo Scottish Cup [1977-79] and the Football League Group Cup [1981])." 		
end select
heading2 = heading1
heading1 = heading1 & "<a style=""font-family:courier;"" id=""xheadtext"" href=""javascript:HeadToggle('headtext');"">[+]</a>"

restrictions = ""
season_text = ""
if season_years1 > "" then
	heading1 = heading1 & " from " & season_years1
	season_text = season_text & " and years >= '" & season_years1 & "' " 
	restrictions = "Y"
end if			
if season_years2 > "" then
	heading1 = heading1 & " to " & season_years2
	season_text = season_text & " and years <= '" & season_years2 & "' "
	restrictions = "Y"
end if

outline = ""
outline = outline & "<center>"
 
  outline = outline & "<table border=""0"" cellspacing=""0"" style=""border-collapse: collapse"" cellpadding=""0"" width=""980"">"
    outline = outline & "<tr>"
    outline = outline & "<td width=""260"" valign=""top"" style=""text-align:center;"">"

	if called <>  "opposition" then
		outline = outline & "<p style=""text-align: center; margin-top:0; margin-bottom:3"">"
		outline = outline & "<a href=""gosdb.asp""><font color=""#404040""><img border=""0"" src=""images/gosdb-small.jpg"" align=""left""></font></a><font color=""#404040"">"
		outline = outline & "<b><font style=""font-size: 15px"">Search by<br>"
		outline = outline & "</font></b><span style=""font-size: 15px""><b>Opposition</b></span></font><p style=""text-align: center; margin-top:0; margin-bottom:0"">"
		outline = outline & "<b>"
		outline = outline & "<a href=""gosdb.asp"">Back to<br>GoS-DB Hub</a></b></p>"
	end if
	
	outline = outline & "</td>"
    
  	outline = outline & "<td width=""460"" align=""center"" style=""text-align: center"">"	

	Dim action

	if called =  "opposition" then
			outline = outline & "<p style=""margin-top: 6px; margin-bottom: 3px;""><font style=""font-family: verdana,arial,helvetica,sans-serif; font-size: 14px;"" color=""#006e32""><b>HEAD TO HEAD RESULTS</b>" & "</font></p>"
			action = "opposition.asp?team=" & idteam & "&s1=" & season_years1 & "&s2=" & season_years2
		else
			outline = outline & "<p style=""margin-top: 9px; margin-bottom: 3px;""><font style=""font-size: 16px;"" color=""#006e32"">HEAD TO HEAD RESULTS</font></p>"
			outline = outline & "<p style=""margin-top: 3px; margin-bottom: 3px;""><font style=""font-size: 14px;""><b>" & ucase(fullteam) & "</b></font></p>"		
			action = "gosdb-results.asp?team=" & fullteam & "&s1=" & season_years1 & "&s2=" & season_years2
	end if

	outline = outline & "<form style=""font-size: 11px; padding: 0; margin: 0 0 6 0;"" action=""" & action & """ method=""post"" name=""form1"">"
	outline = outline & "All Competitions"
	outline = outline & "<input type=""radio"" name=""R1"" "  & check1 & " value=""ALL"" onClick=""javascript:document.forms.form1.submit()"" onMouseOver=""style.cursor='hand'"">" 
	outline = outline & "&nbsp; Football League"
	outline = outline & "<input type=""radio"" name=""R1"" "  & check2 & " value=""FLG"" onClick=""javascript:document.forms.form1.submit()"" onMouseOver=""style.cursor='hand'"">"
	outline = outline & "&nbsp; All Leagues"
	outline = outline & "<input type=""radio"" name=""R1"" "  & check3 & " value=""LGS"" onClick=""javascript:document.forms.form1.submit()"" onMouseOver=""style.cursor='hand'"">"
	outline = outline & "&nbsp; All Cups"
	outline = outline & "<input type=""radio"" name=""R1"" "  & check4 & " value=""CUP"" onClick=""javascript:document.forms.form1.submit()"" onMouseOver=""style.cursor='hand'"">"
	outline = outline & "</form>"
	outline = outline & "<p style=""margin-top: 0px; margin-bottom: 3px;""><font style=""font-size: 12px;"" color=""#202020""><b>" & heading1 & "</b></font></p>"
	outline = outline & "<p style=""margin-top: 0px; margin-bottom: 9px;"">"
	if restrictions = "Y" then 
  	outline = outline & "<font style=""font-size: 11px;"" color=""#900033""><b>Reminder: the results are limited to the seasons you selected </b></font>"
	end if
	if called =  "opposition" then
		else
			outline = outline & "<a href=""http://www.greensonscreen.co.uk/gosdb-headtohead.asp"">[return to overview]</a></p>"
	end if


    outline = outline & "</td>"
    outline = outline & "<td width=""260"" valign=""top""  align=""justify"">"
        
	if called <>  "opposition" then
		outline = outline & "The match details for the selected team. Use [+] for further information."
    end if
    
    outline = outline & "</td>"
    outline = outline & "</tr>"
    
    outline = outline & "</table>"
    	  

outline = outline & "<div id=""headtext"" style=""width:900px; display:none""><p style=""margin:0 0 3 0; text-align: justify"">"
outline = outline & "<b>" & heading2 & ":</b> " & headtext
outline = outline & "</p></div>"
response.write(outline)	

tagno = 1
outlinematch = ""
for i = 0 to ubound(counts,1)
	for j = 0 to ubound(counts,2)
		counts(i,j) = 0
	next
next


sql = "select years, opposition, opposition_qual, date, "
sql = sql & "case homeaway when 'H' then 'H' else 'A' end as homeawayHA"
sql = sql & ", shortcomp, subcomp, goalsfor, goalsagainst, pensfor, attendance, notes, ground_name, ground_name_trad, NULL as not_played_type "
sql = sql & "from " & tableview 
sql = sql & " join opposition on opposition = name_then "
sql = sql & " left outer join venue on name_then = club_name_then and date between first_game and last_game "
sql = sql & " join season on date >= date_start and date <= date_end "
sql = sql & "where name_now = '" & replace(fullteam,"'","''") & "' " & season_text 			'Double-up any apostrophes inside the club name (e.g. Lovell's)

sql = sql & "union all "
sql = sql & "select years, opposition, opposition_qual, date, "
sql = sql & "case homeaway when 'H' then 'H' else 'A' end as homeawayHA"
sql = sql & ", shortcomp, subcomp, NULL, NULL, NULL, NULL, NULL, ground_name, ground_name_trad, not_played_type "
sql = sql & "from match_not_played a   "
sql = sql & " join competition b on a.compcode = b.compcode "
sql = sql & " join opposition on opposition = name_then "
sql = sql & " left outer join venue on name_then = club_name_then and date between first_game and last_game "
sql = sql & " join season on date >= date_start and date <= date_end "
sql = sql & "where name_now = '" & replace(fullteam,"'","''") & "' " & season_text			'Double-up any apostrophes inside the club name (e.g. Lovell's)
sql = sql & notplayed_qual  
sql = sql & "order by date "

rs.open sql,conn,1,2

Do While Not rs.EOF

	'Accumulate stats
	if isnull(rs.Fields("not_played_type")) then
	
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
		crowd = " "
	end if
	
	displaydate = FormatDateTime(rs.Fields("date"),1)
	work1 = split(displaydate," ")
	displaydate = work1(0) & " " & left(work1(1),3) & " " & work1(2) 

	clickon = ""
	if rs.Fields("notes") > "" then clickon = clickon & "N"

	if rs.Fields("years") <> yearslast then 
		topclass = " t"
	 	else topclass = ""
	end if
	
	if rs.Fields("years") <> yearslast then
		outlinematch  = outlinematch & "<tr><td class=""a" & topclass & """ nowrap=""nowrap""><b>" & rs.Fields("years") & "</b></td>" 
		else outlinematch  = outlinematch & "<tr><td class=""a" & topclass & """></td>"
	end if	
	
	venue = rs.Fields("ground_name")
	
	if isnull(rs.Fields("not_played_type")) then
		outlinematch  = outlinematch & "<td class=""a" & topclass & """ nowrap=""nowrap""><a style=""font-family:courier;"" id=""xtag" & tagno & """ href=""javascript:Toggle('tag" & tagno & "','" & clickon & "');"">[+<span style=""font-family:verdana;"">" & clickon & "</span>]</a></td>" 
	  else
	   	outlinematch  = outlinematch & "<td></td>"
	end if   	
	
	outlinematch  = outlinematch & "<td class=""a" & topclass & """ nowrap=""nowrap""><a href=""gosdb-match.asp?date=" & rs.Fields("date") & """>"  & displaydate & "</a><span id=""dtag" & tagno & """ style=""display:none;"">" & rs.Fields("date") & "</span></td>"	

  	if isnull(rs.Fields("pensfor")) then 
  		scoresep = "-"
  	  else
		scoresep = "<a href=""javascript:Toggle('tag" & tagno & "','" & clickon & "');"">+</a>"	  
	end if
		  	 
	if rs.Fields("homeawayHA") = "H" then
		venue = "Home Park"
		if displaydate = "18 Mar 1961" then venue = "Plainmoor"
		outlinematch  = outlinematch & "<td class=""b" & topclass & """ align=""right"">" & "Argyle" & "</td>" 
		if rs.Fields("not_played_type") = "P" then
			outlinematch  = outlinematch & "<td class=""c pa"">POST</td>" 
		  elseif rs.Fields("not_played_type") = "A" then
			outlinematch  = outlinematch & "<td class=""c pa"">ABND</td>"
		  elseif rs.Fields("not_played_type") = "C" then
			outlinematch  = outlinematch & "<td class=""c pa"">CANC</td>"		
		  else
			outlinematch  = outlinematch & "<td class=""c" & topclass & """ nowrap=""nowrap"">" & rs.Fields("goalsfor") & scoresep & rs.Fields("goalsagainst") & "</td>"
		end if
		outlinematch  = outlinematch & "<td class=""d" & topclass & """>" & opposition & " " & rs.Fields("opposition_qual") & "</td>"
		outlinematch  = outlinematch & "<td class=""a" & topclass & """ nowrap=""nowrap"">" & rs.Fields("shortcomp") & " " & rs.Fields("subcomp") & "</td>"
		outlinematch  = outlinematch & "<td class=""a" & topclass & """ align=""right"">" & crowd & "</td>" 
		outlinematch  = outlinematch & "<td class=""a" & topclass & """>" & venue & "</td>"   
		outlinematch  = outlinematch & "</tr>"
	  else  
		outlinematch  = outlinematch & "<td class=""b" & topclass & """ align=""right"">" & opposition & " " & rs.Fields("opposition_qual") & "</td>"  
		if rs.Fields("not_played_type") = "P" then
			outlinematch  = outlinematch & "<td class=""c pa"">POST</td>" 
		  elseif rs.Fields("not_played_type") = "A" then
			outlinematch  = outlinematch & "<td class=""c pa"">ABND</td>"
		  elseif rs.Fields("not_played_type") = "C" then
			outlinematch  = outlinematch & "<td class=""c pa"">CANC</td>"		
		  else
			outlinematch  = outlinematch & "<td class=""c" & topclass & """ nowrap=""nowrap"">" & rs.Fields("goalsagainst") & scoresep & rs.Fields("goalsfor") & "</td>"
		end if
		outlinematch  = outlinematch & "<td class=""d" & topclass & """>" & "Argyle" & "</td>"	
		outlinematch  = outlinematch & "<td class=""a" & topclass & """ nowrap=""nowrap"">" & rs.Fields("shortcomp") & " " & rs.Fields("subcomp") & "</td>"
		outlinematch  = outlinematch & "<td class=""a" & topclass & """ align=""right"">" & crowd & "</td>" 
		outlinematch  = outlinematch & "<td class=""a" & topclass & """>" & venue 
		if not isnull(rs.Fields("ground_name_trad")) then outlinematch  = outlinematch & " (aka " & rs.Fields("ground_name_trad") & ")"  
		outlinematch  = outlinematch & "</td></tr>"		
	end if
	
	outlinematch  = outlinematch & "<tr><td class=""a"" style=""margin: 0 0 0 0; padding: 0 0 0 0;""></td><td style=""margin: 0 0 0 0; padding: 0 0 0 0;""></td><td  class=""a"" colspan=""7"" width=""500px"" style=""border-top-style: none; margin: 0 0 0 0; padding: 0 2 0 6; font-size:10px;""><span id=""tag" & tagno & """ style=""display:none;""></span></td></tr>"	  
	
	yearslast = rs.Fields("years")
	tagno = tagno + 1
	
	rs.MoveNext 
Loop 

rs.close

outline2 = outline2 & "<center>"
outline2 = outline2 & "<div id=""table1"">"
outline2 = outline2 & "<table style=""border-collapse: collapse;"" border=""0"" bordercolor=""#c0c0c0"" cellpadding=""0"" cellspacing=""0"">" 
outline2  = outline2 & "<tbody>"
outline2  = outline2 & "<tr>"
outline2  = outline2 & "<td class=""head"">SEASON</td>"
outline2  = outline2 & "<td class=""head""><a href=""#""><img hspace=""6"" border=""0"" src=""images/help.gif"" onclick=""showtip('Click on [+] to expand the details.<br>\\\'N\\\' indicates a special note.')"" onmouseout=""hidetip()""></a></td>"
outline2  = outline2 & "<td class=""head"">MATCH DATE</td>"
outline2  = outline2 & "<td class=""head"" align=""center"" colspan=""3"">RESULT</td>"
outline2  = outline2 & "<td class=""head"">COMP.</td>"
outline2  = outline2 & "<td class=""head"">ATT.</td>"
outline2  = outline2 & "<td class=""head"">VENUE</td>"
outline2  = outline2 & "</b></tr>"
		
if outlinematch > "" then
	outline2  = outline2 & outlinematch
  else
  	outline2  = outline2 & "<tr><td></td><td></td><td colspan=""5""><b>No matches for this selection</b></td></tr>"
end if		
outline2  = outline2 & "</tbody>"
outline2  = outline2 & "</table>"
outline2  = outline2 & "</div>"

'Construct summary tables

Dim besthome, bestaway, worsthome, worstaway, rankhome, rankaway, counthome, countaway

besthome = ""
bestaway = ""
worsthome = ""
worsthome = ""
rankhome = ""
rankaway = ""
counthome = ""
countaway = ""


sql = "WITH CTE1 AS "
sql = sql & "(select top 1 with ties 1 as extreme,homeaway, year(date) as year, goalsfor, goalsagainst, goalsfor as majorgoals " 
sql = sql & "from " & tableview 
sql = sql & " join opposition on opposition = name_then join season on date between date_start and date_end "
sql = sql & " where name_now = '" & replace(fullteam,"'","''") & "' " & season_text			'Double-up any apostrophes inside the club name (e.g. Lovell's)
sql = sql & " and homeaway = 'H'"
sql = sql & " order by goalsfor-goalsagainst desc, goalsfor desc)"
sql = sql & ",CTE2 AS"
sql = sql & "(select top 1 with ties 1 as extreme,homeaway, year(date) as year, goalsfor, goalsagainst, goalsfor as majorgoals " 
sql = sql & "from " & tableview 
sql = sql & " join opposition on opposition = name_then join season on date between date_start and date_end "
sql = sql & " where name_now = '" & replace(fullteam,"'","''") & "' " & season_text			'Double-up any apostrophes inside the club name (e.g. Lovell's)
sql = sql & " and homeaway = 'A'"
sql = sql & " order by goalsfor-goalsagainst desc, goalsfor desc)"
sql = sql & ",CTE3 AS"
sql = sql & "(select top 1 with ties 0 as extreme,homeaway, year(date) as year, goalsagainst, goalsfor, goalsagainst as majorgoals "
sql = sql & "from " & tableview 
sql = sql & " join opposition on opposition = name_then join season on date between date_start and date_end "
sql = sql & " where name_now = '" & replace(fullteam,"'","''") & "' " & season_text			'Double-up any apostrophes inside the club name (e.g. Lovell's)
sql = sql & " and homeaway = 'H'"
sql = sql & " order by goalsagainst-goalsfor desc, goalsagainst desc)"
sql = sql & ",CTE4 AS"
sql = sql & "(select top 1 with ties 0 as extreme,homeaway, year(date) as year, goalsagainst, goalsfor, goalsagainst as majorgoals "
sql = sql & "from " & tableview 
sql = sql & " join opposition on opposition = name_then join season on date between date_start and date_end "
sql = sql & " where name_now = '" & replace(fullteam,"'","''") & "' " & season_text			'Double-up any apostrophes inside the club name (e.g. Lovell's)
sql = sql & " and homeaway = 'A'"
sql = sql & " order by goalsagainst-goalsfor desc, goalsagainst desc)"
sql = sql & "select year, extreme, homeaway, goalsfor, goalsagainst"
sql = sql & " from CTE1"
sql = sql & " union all "
sql = sql & "select year, extreme, homeaway, goalsfor, goalsagainst"
sql = sql & " from CTE2"
sql = sql & " union all "
sql = sql & "select year, extreme, homeaway, goalsfor, goalsagainst"
sql = sql & " from CTE3"
sql = sql & " union all "
sql = sql & "select year, extreme, homeaway, goalsfor, goalsagainst"
sql = sql & " from CTE4"
sql = sql & " order by year"

rs.open sql,conn,1,2

Do While Not rs.EOF

	if rs.Fields("extreme") = 1 and rs.Fields("homeaway") = "H" then besthome = besthome & "<b>" & rs.Fields("year") & "</b>, " & rs.Fields("goalsfor") & " - " & rs.Fields("goalsagainst") & "<br>"
	if rs.Fields("extreme") = 1 and rs.Fields("homeaway") = "A" then bestaway = bestaway & "<b>" & rs.Fields("year") & "</b>, " & rs.Fields("goalsfor") & " - " & rs.Fields("goalsagainst") & "<br>"
	if rs.Fields("extreme") = 0 and rs.Fields("homeaway") = "H" then worsthome = worsthome & "<b>" & rs.Fields("year") & "</b>, " & rs.Fields("goalsfor") & " - " & rs.Fields("goalsagainst") & "<br>"
	if rs.Fields("extreme") = 0 and rs.Fields("homeaway") = "A" then worstaway = worstaway & "<b>" & rs.Fields("year") & "</b>, " & rs.Fields("goalsfor") & " - " & rs.Fields("goalsagainst") & "<br>"
	
	rs.MoveNext 
Loop 

rs.close

if besthome > "" then besthome = left(besthome,len(besthome)-4)	 'remove last <br>
if bestaway > "" then bestaway = left(bestaway,len(bestaway)-4)	 'remove last <br>
if worsthome > "" then worsthome = left(worsthome,len(worsthome)-4)	 'remove last <br>
if worstaway > "" then worstaway = left(worstaway,len(worstaway)-4)	 'remove last <br>
	

outline1 = "<table border=""0"" cellpadding=""0"" cellspacing=""6"" width=""900px"">"
outline1 = outline1 & "<tr><td align=""right"" valign=""top"" width=""33%"">"

outline1 = outline1 & "<table style=""border-collapse: collapse;"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
outline1 = outline1 & "<tr><td style=""font-size: 11px; font-weight: bold; color:#006e32"" colspan=""3"" align=""center"">"
outline1 = outline1 & "OUR BEST AND WORST"
outline1 = outline1 & "</td></tr>"
outline1 = outline1 & "<tr>"
outline1 = outline1 & "<td></td>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">Home</td>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">Away</td>"
outline1 = outline1 & "</td></tr>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"" valign=""top"">Best</td>"
outline1 = outline1 & "<td valign=""top"">" & besthome & "</td>"
outline1 = outline1 & "<td valign=""top"">" & bestaway & "</td>"
outline1 = outline1 & "</td></tr>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32; margin-top:3px;"" valign=""top"">Worst</td>"
outline1 = outline1 & "<td style=""margin-top:3px;"" valign=""top"">" & worsthome & "</td>"
outline1 = outline1 & "<td style=""margin-top:3px;"" valign=""top"">" & worstaway & "</td>"
outline1 = outline1 & "</tr></table>"

outline1 = outline1 & "</td>"
outline1 = outline1 & "<td style=""padding-left:40px; padding-right:40px;"" align=""center"" valign=""top"" width=""33%"">"

outline1 = outline1 & "<table style=""border-collapse: collapse;"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""240px"">"
outline1 = outline1 & "<tr><td style=""font-size: 11px; font-weight: bold; color:#006e32"" colspan=""8"" align=""center"">"
outline1 = outline1 & "SELECTION TOTALS"
outline1 = outline1 & "</td></tr>"
outline1 = outline1 & "<tr>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32""></td>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">P</td>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">W</td>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">D</td>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">L</td>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">F</td>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">A</td>"
outline1 = outline1 & "<td  style=""font-weight: bold; color:#006e32"" align=""right"">Avg Att</td>"
outline1 = outline1 & "</tr><tr>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">Home</td>"
outline1 = outline1 & "<td>" & counts(0,0) & "</td>"
outline1 = outline1 & "<td>" & counts(1,0) & "</td>"
outline1 = outline1 & "<td>" & counts(2,0) & "</td>"
outline1 = outline1 & "<td>" & counts(3,0) & "</td>"
outline1 = outline1 & "<td>" & counts(4,0) & "</td>"
outline1 = outline1 & "<td>" & counts(5,0) & "</td>"
if counts(7,0) > 0 then
	outline1 = outline1 & "<td align=""right"">" & FormatNumber(Round(counts(6,0)/counts(7,0)),0,0,0,-1) & "</td>"
		else outline1 = outline1 & "<td align=""right"">0</td>"
end if
outline1 = outline1 & "</tr><tr>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">Away</td>"
outline1 = outline1 & "<td>" & counts(0,1) & "</td>"
outline1 = outline1 & "<td>" & counts(1,1) & "</td>"
outline1 = outline1 & "<td>" & counts(2,1) & "</td>"
outline1 = outline1 & "<td>" & counts(3,1) & "</td>"
outline1 = outline1 & "<td>" & counts(4,1) & "</td>"
outline1 = outline1 & "<td>" & counts(5,1) & "</td>"
if counts(7,1) > 0 then
	outline1 = outline1 & "<td align=""right"">" & FormatNumber(Round(counts(6,1)/counts(7,1)),0,0,0,-1) & "</td>"
		else outline1 = outline1 & "<td align=""right"">0</td>"
end if
outline1 = outline1 & "</tr><tr>"
outline1 = outline1 & "<td style=""font-weight: bold; color:#006e32"">Both</td>"
outline1 = outline1 & "<td>" & counts(0,0) + counts(0,1) & "</td>"
outline1 = outline1 & "<td>" & counts(1,0) + counts(1,1) & "</td>"
outline1 = outline1 & "<td>" & counts(2,0) + counts(2,1) & "</td>"
outline1 = outline1 & "<td>" & counts(3,0) + counts(3,1) & "</td>"
outline1 = outline1 & "<td>" & counts(4,0) + counts(4,1) & "</td>"
outline1 = outline1 & "<td>" & counts(5,0) + counts(5,1) & "</td>"
if counts(7,0)+counts(7,1) > 0 then
	outline1 = outline1 & "<td align=""right"">" & FormatNumber(Round((counts(6,0)+counts(6,1))/(counts(7,0)+counts(7,1))),0,0,0,-1) & "</td>"
	else outline1 = outline1 & "<td align=""right"">0</td>"
end if
outline1 = outline1 & "</tr></table>"

outline1 = outline1 & "</td>"
outline1 = outline1 & "<td align=""left"" valign=""top"" width=""33%"">"

sql = "select homeaway, count(distinct name_now) as count "
sql = sql & "from ("
sql = sql & " select homeaway, name_now, count(*) as count "
sql = sql & " from " & tableview
sql = sql & " join opposition on opposition = name_then join season on date between date_start and date_end "
sql = sql & " where homeaway in ('H','A') " & season_text
sql = sql & " group by homeaway, name_now "
sql = sql & " having count(*) > 3 "
sql = sql & " ) as subsel "
sql = sql & "group by homeaway"

rs.open sql,conn,1,2

Do While Not rs.EOF

	if rs.Fields("homeaway") = "H" then
		counthome = rs.Fields("count") 
		else countaway = rs.Fields("count") 
	end if
	
	rs.MoveNext 
Loop 

rs.close

sql = "WITH CTE1 AS "
sql = sql & "( "
sql = sql & "select homeaway, name_now, cast(sum(points) as dec(7,2))/SUM(p) as pointspergame "
sql = sql & "from ( "
sql = sql & " select homeaway, name_now, 1 as p, "
sql = sql & " case when goalsfor > goalsagainst then 3 " 
sql = sql & " when goalsfor = goalsagainst then 1 " 
sql = sql & " when goalsfor < goalsagainst then 0 end as points "
sql = sql & " from " & tableview 
sql = sql & " join opposition on opposition = name_then join season on date between date_start and date_end "
sql = sql & " where homeaway in ('H','A') " & season_text
sql = sql & ") as sub "
sql = sql & "group by homeaway, name_now "
sql = sql & "having count(*) > 3 "
sql = sql & "), "
sql = sql & "CTE2 as "
sql = sql & "( "
sql = sql & "select rank() over(partition by homeaway order by homeaway, pointspergame desc) as rank, homeaway, pointspergame, name_now "
sql = sql & "from CTE1 "
sql = sql & ") "
sql = sql & "select rank, homeaway "
sql = sql & "from CTE2 "
sql = sql & "where name_now = '" & replace(fullteam,"'","''") & "' " 		'Double-up any apostrophes inside the club name (e.g. Lovell's)

rs.open sql,conn,1,2

rankhome = "-"
rankaway = "-"

Do While Not rs.EOF

	if rs.Fields("homeaway") = "H" then
		rankhome = rs.Fields("rank") 
		else rankaway = rs.Fields("rank") 
	end if
	
	rs.MoveNext 
Loop 

rs.close

outline1 = outline1 & "<table border=""0"" cellspacing=""0"" style=""border-collapse: collapse;"" bordercolor=""#c0c0c0"" width=""200"">"
outline1 = outline1 & "<tr>"
outline1 = outline1 & "<td align=""center"" style=""font-size: 11px; font-weight: bold; color:#006e32;"" colspan=""3"">"
outline1 = outline1 & "SUCCESS RANKING</td>"
outline1 = outline1 & "</tr>"
outline1 = outline1 & "<tr>"
outline1 = outline1 & "<td  style=""font-weight: bold; color:#006e32"" width=""90"" align=""right"">Home</td>"
outline1 = outline1 & "<td width=""20"" align=""center""><img border=""0"" src=""images/help.gif"" onclick=""showtip('If we\\\'ve played this team more than 3 times in the chosen category, these numbers show how well we\\\'ve done compared with all other clubs. Our best record shows as 1. Any number lower than half-way indicates results that are better than average. A high number suggests a bogey side.')"" onmouseout=""hidetip()""></td>"
outline1 = outline1 & "<td  style=""font-weight: bold; color:#006e32"" width=""90"" align=""left"">Away</td>"
outline1 = outline1 & "</tr>"
outline1 = outline1 & "<tr>"
outline1 = outline1 & "<td align=""right""><span style=""font-size: 18px; font-weight: bold; color:#61A76D;"">" & rankhome & "</span><br>of " & counthome & "</td>"
outline1 = outline1 & "<td align=""center""></td>"
outline1 = outline1 & "<td align=""left""><span style=""font-size: 18px; font-weight: bold; color:#404040;"">" & rankaway & "</span><br>of " & countaway & "</td>"
outline1 = outline1 & "</tr>"
outline1 = outline1 & "</table>"

outline1 = outline1 & "</td></tr></table>"


response.write(outline1) 'summary table
%>
<table border="0" cellspacing="0" style="border-collapse: collapse; margin: 9 0 9 0;">
  <tr>
    <td nowrap><b><font color="#457B44">SINCE LAST </font></b></td>
    <td style="font-weight: bold; color:#006e32">Win</td>
    <td style="font-weight: bold; color:#006e32" padding-left=12px">Draw</td>
    <td style="font-weight: bold; color:#006e32" padding-left=12px">Defeat</td>
  </tr>
  <tr>
    <td style="font-weight: bold; color:#006e32">Home</td>
    <td nowrap><% response.write(periodfunc(latestdates(1,0))) %>&nbsp;</td>
    <td nowrap style="padding-left=12px"><% response.write(periodfunc(latestdates(2,0))) %>&nbsp;</td>
    <td nowrap style="padding-left=12px"><% response.write(periodfunc(latestdates(3,0))) %>&nbsp;</td>
  </tr>
  <tr>
    <td style="font-weight: bold; color:#006e32">Away</td>
    <td nowrap><% response.write(periodfunc(latestdates(1,1))) %>&nbsp;</td>
    <td nowrap style="padding-left=12px"><% response.write(periodfunc(latestdates(2,1))) %>&nbsp;</td>
    <td nowrap style="padding-left=12px"><% response.write(periodfunc(latestdates(3,1))) %>&nbsp;</td>
  </tr>
</table><br>
<%
response.write(outline2) 'match tables

conn.close

Function periodfunc(matchdate) 
Dim displaydate, mnths, yrs, remmnths
 if isdate(matchdate) = "True" then
	displaydate = FormatDateTime(matchdate,1)
	mnths = DateDiff("m",displaydate,Date)
	yrs = Int(mnths/12)
	remmnths = mnths - 12*yrs
	select case yrs
 		case 0
 			yrs = ""
 		case 1	
  			yrs = yrs & " yr "
 		case else 
 			yrs = yrs & " yrs "
 	end select
 	select case remmnths
 		case 0
 			remmnths = ""
 		case 1	
  			remmnths = remmnths & " mth "
 		case else 
 			remmnths = remmnths & " mths "
 		end select 
	periodfunc = yrs & remmnths 
  else periodfunc = "Never"
 end if 	
End Function

%><%'="a" %><%		

%>

<br>
<!--#include file="base_code.htm"-->
</p>
</body>
</html>