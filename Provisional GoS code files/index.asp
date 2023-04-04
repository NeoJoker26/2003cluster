<%@ Language=VBScript %> 
<% Option Explicit %> 
<% Dim f1, f2, objFolder, objFile, File, fact, fact1, byte1, byte2, thisdate, basedate, days, i, pos 
   Dim opposition, opp_qual, lastdate, lastdatefull, goals_for, goals_against, competition, game_no, homeaway 
   Dim thisday, thismon, todaypic, todaycaption, todayline, sameday, minusdays, todaypart, audiodone, parm
   Dim lastpublish, lasteventdate, lasteventtype, lastmaterialseq, lastheading, lasttext, lastlinktext
   Dim seconds, days1, hours, mins, kickoff, dateparts, motdhold, motdday, motdmonth, motdfound, daycode1, daycode2, thumb1, thumb2
   
   Const timespan = 10
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="imagetoolbar" content="no">
<title>Greens on Screen</title>
<link href="https://www.greensonscreen.co.uk/images/favicon.ico" rel="shortcut icon">
<link rel="stylesheet" type="text/css" href="gos2.css">
<style>
<!--
#outside {background: #ffffff; width:1000px}
#welcome {width:600px; color: #43784D; font: 15px Arial; font-weight: 700; margin:0 0 12px;}
			
#box1 {clear:both; width: 336px; float:left; margin: 0 4px 0 0; padding: 0;}

#box1a {width: 330px; margin: 0 auto 18px; padding: 0;}

#box1ba {width: 166px; float:left; margin: 0 4px 0 0; padding: 0;}
#box1ba p {text-align:left; margin: 0; padding: 0 0 0 4px; font-size:11px}
#box1ba ul {width:99%; margin: 0; padding: 0; list-style-type: none; }
#box1ba li {border:1px solid #c0c0c0; border-top:1px solid #ffffff; background: #fffff8;}
#box1ba a:link {color: #006c00; text-decoration: none;  font-weight:normal;}
#box1ba a:visited {color: #006c00; text-decoration: none; font-weight:normal;}
#box1ba a:hover {background: #e0f0e0; color: #000000; font-weight:normal;}

#box1bb {width: 166px; float:right; margin: 0; padding: 0;}
#box1bb p {text-align:left; margin: 0; padding: 6px 4px; font-size:11px}
#box1bb ul {width:100%; margin: 0; padding: 0; list-style-type: none; }
#box1bb li {border:1px solid #c0c0c0; border-top:1px solid #ffffff; background: #fffff8;}
#box1bb a {display: block; width: 100%; color: black; text-decoration: none; background: #e0f0e0; text-align: left;}
#box1bb a:link, a:visited {color: black; }
#box1bb a:hover {background: transparent; color: #004438; letter-spacing: 0em; font-weight:normal; }

#box2 {width: 656px; float:right; margin: 0 4px 0 0; padding: 0;}

#box2a {width: 320px; float:left; margin: 0; padding: 0; font-size:11px;  position: relative;}
#box2a a:link { color: #457b44; text-decoration: none; }
#box2a a:visited {color: #457b44; text-decoration: none; }
#box2a a:hover {background: transparent; color: #457b44; background: #e0f0e0; }

#box2a1 {width: 320px; margin: 0; padding: 0;}
#box2a1 .date {text-align: center; margin: 2px; }
#box2a1 .headline {font-size: 11px; color: #457b44; font-weight: bold; margin: 2px 0; text-align: center;}
#box2a1 .score {font-size: 11px; color: #404040; font-weight: bold; margin: 2px 0; text-align: center;}
#box2a1 p {text-align: center; margin: 6px; }
#box2a1 ul {width: 312px; margin: 0; padding: 0; list-style-type: none; }
#box2a1 a:link { color: #457b44; text-decoration: none; text-align: center;}
#box2a1 a:visited {color: #457b44; text-decoration: none; text-align: center;}
#box2a1 a:hover {background: transparent; color: #457b44; font-weight:bold;}

#box2a2 {width: 310px; margin: 0 0 12px 0; padding: 0;}
#box2a2 p {text-align: left; margin: 0}
#box2a2 ul {width: 99%; margin: 0; padding: 0; list-style-type: none; }
#box2a2 a:link {color: #457b44; text-decoration:none;}
#box2a2 a:visited {color: #457b44; text-decoration:none;}
#box2a2 a:hover {background: #e0f0e0; color: #000000; font-weight:normal;}

#box2b {width: 336px; float:right; margin: 0; padding: 0;}
#box2b p {text-align:left; margin: 0; padding: 4px 4px; font-size:11px}
#box2b ul {width: 99%; margin: 0; padding: 0; list-style-type: none; }

#box2ba {width: 300px; margin: 0 auto 24px; padding: 0;}

#box2ba1 {width: 166px; float:left; margin: 0 4px 0; padding: 0;}
#box2ba1 p {text-align:left; margin: 0; padding: 6px 4px; font-size:11px}
#box2ba1 ul {width:99%; margin: 0; padding: 0; list-style-type: none; }
#box2ba1 li {border:1px solid #c0c0c0; border-top:1px solid #ffffff; background: #fffff8;}
#box2ba1 a {display: block; width: 100%; color: black; text-decoration: none; background: #e0f0e0; text-align: left;}
#box2ba1 a:link, a:visited {color: black; }
#box2ba1 a:hover {background: transparent; color: #004438; letter-spacing: 0em; font-weight:normal; }

#box2ba2 {width: 162px; float:right; margin: 0; padding: 0;}
#box2ba2 p {text-align:left; margin: 0; padding: 6px 4px; font-size:11px}
#box2ba2 ul {width:99%; margin: 0; padding: 0; list-style-type: none; }
#box2ba2 li {border:1px solid #c0c0c0; border-top:1px solid #ffffff; background: #fffff8;}
#box2ba2 a {display: block; width: 100%; color: black; text-decoration: none; background: #e0f0e0; text-align: left;}
#box2ba2 a:link, a:visited {color: black; }
#box2ba2 a:hover {background: transparent; color: #004438; letter-spacing: 0em; font-weight:normal; }

.button1 {
    padding: 4px 8px; 
    margin: 0 0 6px;
    font-family: verdana, sans-serif;
    display: inline-block;
    white-space: nowrap;
    font-size: 11px;
    position: relative;
    outline: none;
    overflow: visible;
    cursor: pointer;
    border-radius: 3px;
    border: 1px solid #808080;
    color: #000000 !important; 
    background: linear-gradient(#e0f0e0,#d0e0d0);  
    text-transform: uppercase;
}

.button1:hover {
    background: transparent;
    border: 1px solid #202020; 
}

#miscreports { 
	display:none;
	border:0px none;
	border-collapse: collapse; 
	margin-left:40px; 
	margin-right:0; 
	margin-top:0; 
	margin-bottom:12px;
}
#miscreports a:link {color: #006c00; text-decoration: none;  font-weight:normal;}
#miscreports a:visited {color: #006c00; text-decoration: none; font-weight:normal;}
#miscreports a:hover {background: #ffffff; color: #000000; font-weight:normal;}
#miscreports td {padding: 2px 3px;}

#histchapters {
	display:none;
	border:0px none;
	border-collapse: collapse; 
	margin: 0 0 12px;
}
#histchapters a:link {color: #006c00; text-decoration: none;  font-weight:normal;}
#histchapters a:visited {color: #006c00; text-decoration: none; font-weight:normal;}
#histchapters a:hover {background: #ffffff; color: #000000; font-weight:normal;}
#histchapters td {padding: 2px 3px;}

#matchdetails {width:550px; margin: 0 auto; padding: 15px 25px; text-align: left; font-size: 11px; line-height: 130%; border:1px solid #c0c0c0; background: #fffffd;}
#matchdetails h1 {font-size: 14px; color: #404040; font-weight: bold; margin: 0 0 10px;}
#matchdetails h2 {font-size: 12px; color: #3f7855; margin-top: 6px;}
#matchdetails .score {font-size: 14px; font-weight: 700; margin: 9px 0 6px;}
#matchdetails .penalties { font-size: 11px; margin-top: 6px;}
#matchdetails a:link {color: #006c00; text-decoration: none;  font-weight:normal}
#matchdetails a:visited {color: #006c00; text-decoration: none; font-weight:normal;}
#matchdetails a:hover {background: #e0f0e0; color: #000000; font-weight:normal;}

#report {
	text-align: justify;
	max-width: 100%;
	line-height: 130%;
}

#milestones ul {
	list-style-type: circle; 
	padding-left: 18px;
	margin: 0;
}

#milestones ul.nav  {
	margin: 18px 0 24px -18px;
}
 
#milestones ul li.cell {
	font-size: 11px;
	margin: 0 12px 6px 0;
	padding: 0 5px;
}

.score { font-size: 11px; font-weight: 700; margin: 10px 0 6px; }
.penalties { font-size: 11px; margin-top: 6px; font-weight: 700; }
.team { margin: 6px 0 0; }
.opp { color: #606060; }
.goals { margin: 9px 0 0; }
.venue, .attendance, .visitors, .totpoints { margin-right: 18px; }

.bold { font-weight: 700; }
.green { color: #40703f; }
.grey { color: #606060; }
.hover {width:fit-content; background: #e0f0e0; color: black; cursor: pointer;}

#notice {width:600px; margin:0 auto 10px; }
#notice p {font-size: 11px; margin: 0 10px 4px; text-align:justify; line-height:1.3;}

#display1a {width:800px; margin-bottom:12px; display:none; position: relative}
#display1b {width:800px; margin-bottom:12px; display:none; position: relative}
.potdheading {margin:0 auto 8px; font-size:11px; font-weight:bold;}
.caption {width:fit-content; width:-moz-fit-content; position:absolute; bottom:10px; margin:0 auto; padding:1px 8px 2px; left:0; right:0; background:#fffffd; font-size:11px; border-radius:6px;}
#display2 {width:628px; margin:0 auto 25px; display:none;}

a:link .ordlink  {color: #457b44; text-decoration: underline; font-weight:normal;}
a:visited .ordlink  {color: #457b44; text-decoration: none; font-weight:normal;}
a:hover .ordlink  {background: transparent; color: #004438; letter-spacing: 0em; font-weight:normal; }

.WNtag {background: #3f7855; color: #FFFFFF; padding: 0px 3px 1px 3px; margin: 0 4px 0 0; font-weight: bold;}

-->
</style>

<script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script>
<script>
$(document).ready(function(){

	$('.showimage1').css('cursor', 'pointer');
	$('.showimage2').css('cursor', 'pointer');
	
	$('.showimage1').click(function() {
		$('#notice').hide('fast');
		$('#display1b').hide('fast');
		$('#display1a').show('slow');
		ga('send','event','PotD','show',$(this).attr('id'));
	});
	
	$('.showimage2').click(function() {
		$('#notice').hide('fast');
		$('#display1a').hide('fast');
		$('#display1b').show('slow');
		ga('send','event','PotD','show',$(this).attr('id'));
	});
	
	$("#display1a").on("click",".close2", function(){
		$('#display1a').hide('fast');
		$('#notice').show('fast');
	});
	$("#display1b").on("click",".close2", function(){
		$('#display1b').hide('fast');
		$('#notice').show('fast');
	});

		
	$("#display2").on("click",".close", function(){
		$("#display2").hide('fast');
		$('#notice').show('fast');
		$(".potdheading").show('fast');
		$(".showimage1").show('fast');
    	$(".showimage2").show('fast');
	});
	   
	$(".motdmatch").hover(function() {
    $(this).toggleClass("hover");
	});
	
    $(".motdmatch").click(function(){
    	$(".potdheading").hide();
    	$(".showimage1").hide();
    	$(".showimage2").hide();
    	$("#display1a").hide();
  		$("#display1b").hide();
		$('#notice').hide();
        $("#display2").append('<img style="margin-top:10px" src="images/ajax-loader.gif">');
    	$("#display2").show();
    	var matchdate = $(this).attr('id');
    	$.ajax({url: "gosdb-getmotd.asp?date=" + matchdate, success: function(result){
    		$("#display2").html(result);
    		}});
    	$("#display2").show();
    	$(window).scrollTop(0);
    	ga('send','event','MotD','show',matchdate);
	});
	
	$(".open_gosdb_menu").on("click",function(){
		$(".open_gosdb_menu").hide(100);
		$(".close_gosdb_menu").show(100);
		$("#miscreports").show(100);
	});
	$(".close_gosdb_menu").on("click",function(){
		$(".close_gosdb_menu").hide(100);
		$(".open_gosdb_menu").show(100);
		$("#miscreports").hide(100);
	});
	
	$(".open_hist_chapters").on("click",function(){
		$(".open_hist_chapters").hide(100);
		$(".close_hist_chapters").show(100);
		$("#histchapters").show(100);
	});
	$(".close_hist_chapters").on("click",function(){
		$(".close_hist_chapters").hide(100);
		$(".open_hist_chapters").show(100);
		$("#histchapters").hide(100);
	});

});
</script>

</head>

<body>

<% 
Dim conn,sql,sqlorder,rs,rs1 
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

minusdays = Request.QueryString("day")
if minusdays = "" then minusdays = 0
thisdate = Date - minusdays
if IsDate(Request.QueryString("altdate")) then thisdate = Request.QueryString("altdate")

%>

<!--#include file="top_code.htm"-->

<%	
sql = ""
sql = "with CTE as ("
sql = sql & "select row_number() over(order by date) as gameno, date "
sql = sql & "from season_this a join competition b on a.compcode = b.compcode "
sql = sql & "where lfc = 'F' "
sql = sql & ") "
sql = sql & "select a.date, homeaway, opposition, opposition_qual, goalsfor, goalsagainst, competition, subcomp, name_then_short "
sql = sql & ", gameno "
sql = sql & "from match a join competition b on a.compcode = b.compcode join opposition c on a.opposition = c.name_then "
sql = sql & "left outer join CTE d on a.date = d.date "
sql = sql & "where a.date = (select max(event_date) from event_control where event_type = 'M' and event_published = 'Y') " 
sql = sql & "  and a.date >= (select max(date_start) from season) "   
		
rs.open sql,conn,1,2
		
If Not rs.EOF then
		
	lastdatefull = WeekdayName(Weekday(rs.Fields("date"))) & " " & Day(rs.Fields("date")) & " " & MonthName(Month(rs.Fields("date"))) & ", " & Year(rs.Fields("date"))

	lastdate = rs.Fields("date")
	competition = rs.Fields("competition")
	if IsNull(rs.Fields("subcomp")) then
		game_no = rs.Fields("gameno")
	  else
		game_no = trim(rs.Fields("subcomp"))
	end if	 
	if len(rs.Fields("opposition")) < 14 then
		opposition = rs.Fields("opposition")
	  else
		opposition = rs.Fields("name_then_short")
	end if				
	homeaway = rs.Fields("homeaway")
	goals_for = rs.Fields("goalsfor")
	goals_against = rs.Fields("goalsagainst")
	opp_qual = rs.Fields("opposition_qual")
		
end if	
rs.close
%>


<div id="outside">
  <div id="welcome">Welcome to the sights, sounds and history of Plymouth Argyle Football Club</div>
  
  <%
  daycode1 = right("0" & month(dateadd("d",-1,thisdate)),2) & right("0" & day(dateadd("d",-1,thisdate)),2) 		'yesterdays daycode
  daycode2 = right("0" & month(thisdate),2) & right("0" & day(thisdate),2)										'today's daycode	
 
  sqlorder = ""
  if daycode1 > daycode2 then	sqlorder = " desc"	' this happens for Dec 31st (1231) and Jan 1st (0101)
 			
  sql = "select filename, caption "
  sql = sql & "from potd "
  sql = sql & "where daycode in (" & daycode1 & "," & daycode2 & ") " 
  sql = sql & "order by daycode" & sqlorder
  		
  rs.open sql,conn,1,2
		
  if Not rs.EOF then	
  	response.write("<div id=""display1a"">")		 
	response.write("<img style=""width:800px;"" src=""images/dailyphoto/" & rs.Fields("filename") & """>") 
	if not isnull(rs.Fields("caption")) then response.write("<p class=""caption"">" & rs.Fields("caption") & "</p>") 
	thumb1 = "<img id=""daycode1"" class=""showimage1"" style=""width:150px; margin-right:5px; position: relative;"" src=""images/dailyphoto/" & rs.Fields("filename") & """>"
	response.write("<img class=""close2"" style=""position:absolute; top:15px; right:15px; cursor: pointer;"" src=""images/close2.png"">")
	response.write("</div>")
 			
	rs.MoveNext
 	
	response.write("<div id=""display1b"">")		 
	response.write("<img style=""width:800px;"" src=""images/dailyphoto/" & rs.Fields("filename") & """>") 
	if not isnull(rs.Fields("caption")) then response.write("<p class=""caption"">" & rs.Fields("caption") & "</p>")
	thumb2 = "<img id=""daycode2"" class=""showimage2"" style=""width:150px;"" src=""images/dailyphoto/" & rs.Fields("filename") & """>"
	response.write("<img class=""close2"" style=""position:absolute; top:15px; right:15px;"" src=""images/close2.png"">")
	response.write("</div>")
 			
  end if
 		
  rs.close

  Call Notices

  %>

    
  <div id="box1">
  
    <div id="box1a">
  
    <img border="0" src="images/gosdb-small.jpg" width="180" height="80">
    <p style="width:200px; margin: 4px 0; text-align:center; font-size: 10px;">GoS-DB is a relational database of over 100,000 facts and stats from 1903 to <%response.write(lastdatefull)%></p>
 	
 	<div style="margin: 0 auto">
 	<p class="style1bold" style="text-align:center; margin: 6px 0;">Search by ...</p>
	<a class="button1" href="gosdb-match.asp">Match Date</a>
	<a class="button1" href="gosdb-dates.asp">Day</a>
    <a class="button1" href="gosdb-seasons.asp">Seasons</a>
    <a class="button1" href="gosdb-headtohead.asp">Opposition</a>
    <a class="button1" href="gosdb-players1.asp?rank=rank">Top Players</a>
    <a class="button1" href="gosdb-players1.asp">All Players</a>
    <a class="button1" href="gosdb-managers.asp">Managers</a>
	</div>
 
 	<p class="style1bold" style="text-align:center; padding:0; margin:-2px auto 4px;">or</p>
    <p class="style1 open_gosdb_menu" style="text-align:center; padding:0; margin:4px auto 9px;"><a class="button1" style="padding: 3px 6px" href="#">Show Reports</a> <a class="button1" style="padding: 3px 6px" href="gosdb.asp">Full G<span style="text-transform:lowercase;">o</span>S-DB</a></p>
    <p class="style1 close_gosdb_menu" style="display:none; text-align:center; padding:0; margin:4px auto 9px;"><a class="button1" style="padding: 3px 6px" href="#">Hide Reports</a></p>
        
    <table id="miscreports">
      <tr>
        <td style="text-align: right">1</td>
        <td><a href="gosdb-misc1.asp">Competition Totals</a></td>
      </tr>
      <tr>
        <td style="text-align: right">2</td>
        <td><a href="gosdb-misc2.asp">Consecutive Results</a></td>
      </tr>
      <tr>
        <td style="text-align: right">3</td>
        <td><a href="gosdb-misc3.asp">Attendance Highs and Lows</a></td>
      </tr>
      <tr>
        <td style="text-align: right">4</td>
        <td><a href="gosdb-misc4.asp">Top Substitutes</a></td>
      </tr>
      <tr>
        <td style="text-align: right">5</td>
        <td><a href="gosdb-misc5.asp">Youngest and Oldest</a></td>
      </tr>
      <tr>
        <td style="text-align: right">6</td>
        <td><a href="gosdb-misc6.asp">Best and Worst Starts</a></td>
      </tr>
      <tr>
        <td style="text-align: right">7</td>
        <td><a href="gosdb-misc7.asp">Football League by Decade</a></td>
      </tr>
      <tr>
        <td style="text-align: right">8</td>
        <td><a href="gosdb-misc8.asp">Football League by Calendar Year</a></td>
      </tr>
      <tr>
        <td style="text-align: right">9</td>
        <td><a href="gosdb-misc9.asp">Success Rankings by Opposition</a></td>
      </tr>
      <tr>
        <td style="text-align: right">10</td>
        <td><a href="gosdb-misc10.asp">Goalscorers by Season</a></td>
      </tr>
      <tr>
        <td style="text-align: right">11</td>
        <td><a href="gosdb-misc11.asp">Score Counts</a></td>
      </tr>
      <tr>
        <td style="text-align: right">12</td>
        <td><a href="gosdb-misc12.asp">Consecutive Appearances</a></td>
      </tr>
    </table>
    
    </div>
  
  	<div id="box1ba">

      <ul>   
        <li style="border-top:1px solid #c0c0c0; background:#e0f0e0;">
        <p style="font-size: 11px; color: #101010; border-top: medium none; padding: 5px">
        <b>WHAT&#39;S NEW?</b></p>
	   	</li>

        <% Call WhatsNew %>
       
	   </ul>
	</div>
	   
	<div id="box1bb">
    
      <ul>
        <li style="border-top:1px solid #c0c0c0;">
        <p style="padding: 5px; border-top: medium none; font-size: 11px; color:#101010;">
        <b>HERE AND NOW</b></p>
        </li>
        
        <li><a href="dailydiary.asp"><p>Daily Diary &amp; OTD</p></a></li>
        <li><a href="squad.asp"><p>Current Squad</p></a></li>
        <li>	
<%	
		If lastdate > "" then
			response.write("<a href=""gosdb-match.asp?date=" & lastdate & """>")
			response.write("<p>The Latest Match:<br><span style=""font-style:italic"">" & competition)
			response.write(" #" & game_no)
			response.write("</span><br>")
			if homeaway = "H" then
				response.write("Argyle " & goals_for & " " & opposition & " " & opp_qual & " " & goals_against)
			  else
				response.write(opposition & " " & opp_qual & " " & goals_against & " Argyle " & goals_for)
			end if
			response.write("</p></a>")
  		  else
			response.write("<a href=""#""><p>The Latest Match:<br><span style=""font-style:italic"">None played so far this season</span></p></a>")
		end if	
%>        					
	   </li> 
	   <li><a href="progresstables.asp"><p>League Table Plus</p></a></li>
       <li><a href="gosdb-season.asp"><p>Appearance Chart</p></a></li>
       <li><a href="progressgraphs.asp"><p>Progress Graphs</p></a></li>

       </ul>
       
     </div>
	   
  </div>
   
  <div id="box2">
   
  <div id="box2b">

	<div id="box2ba">
	
	<img border="0" src="images/history-small.jpg" width="180" height="80">
	<p style="margin: 1px 12px 12px; width:255px; text-align:center; font-size: 10px;">
	The History of Argyle is an original, comprehensive and thoroughly researched account of Plymouth Argyle 
	- formerly Argyle FC - from its earliest roots in 1886 to, so far, 1957</p>
	
	<a class="button1" href="argylehistory.asp?era=1886-1890">How It All Began</a>
	<a class="button1" href="argylehistorymenu.asp">History Menu</a>
    <a class="button1" href="argylehistory.asp?era=anx2">Argyle FC Facts</a>
    <a class="button1" href="argylehistory_shirt.pdf">Green Jersey Journey</a>
    <p class="style1 open_hist_chapters" style="text-align:center; padding:0; margin:4px auto 9px;"><a class="button1" style="padding: 3px 6px" href="#">Show Chapters</a></p>
    <p class="style1 close_hist_chapters" style="display:none; text-align:center; padding:0; margin:4px auto 9px;"><a class="button1" style="padding: 3px 6px" href="#">Hide Chapters</a></p>

 	
	<table id="histchapters">
		<tr><td><a href="argylehistory.asp?era=1886-1890">1886-90: In the Beginning</a></td></tr>
		<tr><td><a href="argylehistory.asp?era=1890-1895">1890-95: From Struggle to Demise</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1895-1899">1895-99: Rising from the Phoenix</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1899-1900">1899-00: The Argyle Athletic Umbrella</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1900-1901">1900-01: Home Park Home</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1901-1902_1">1901-02 (1): The Argyle Affair</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1901-1902_2">1901-02 (2): The Big Boys Come to Town</a></td></tr>
		<tr><td><a href="argylehistory.asp?era=1902-1903_1">1902-03 (1): Argyle FC Becomes Semi-Pro</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1902-1903_2">1902-03 (2): Birth of Plymouth Argyle</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1902-1903_3">1902-03 (3): Argyle FC's Final Triumph</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1903-1910">1903-10: Plymouth Argyle's Early Years</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1910-1920">1910-20: Argyle and the Great War</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1920-1930">1920-30: Into the Football League</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1930-1934">1930-34: Life in the Second Division</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1934-1939">1934-39: The End of an Era</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1939-1945">1939-45: The Second World War Years</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1945-1950">1945-50: From the Ashes</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1950-1953">1950-53: Into the Fifties</a></td></tr>
      	<tr><td><a href="argylehistory.asp?era=1953-1957">1953-57: From Best to Worst, via Hollywood</a></td></tr>
  
	</table>				
    
    </div>
    
    <div id="box2ba1">
    
    <ul>
        <li style="border-top:1px solid #c0c0c0;">
        <p style="padding: 5px; border-top: none; font-size: 11px; color:#101010;">
        <b>THEN AND NOW</b></p>
        </li>
        <li><a href="achievements.asp"><p>Records &amp; Achievements</p></a></li>
        <li><a href="teampics.asp"><p>Team Photos</p></a></li>
        <li><a href="sv-friendlies.asp"><p>Friendly Matches</p></a></li>
        <li><a href="youtubeview.asp?parm=1"><p>Videos Found Online</p></a></li>
        <li><a href="photoarchive.asp"><p>Non-Match Photo Archive</p></a></li>
    </ul>
    
    </div>
    
    <div id="box2ba2">
    
    <ul>
        <li style="border-top:1px solid #c0c0c0;">
        <p style="padding: 5px 0 5px 5px; border-top: none; font-size: 11px; color:#101010;">
        <b>OTHER SITES</b><span style="font-size:10px"> (in new tab)</span></p>
        </li>
        <li><a target="_blank" href="https://www.pafc.co.uk"><p>PAFC Official</p></a></li>
        <li><a target="_blank" href="https://www.facebook.com/ArgyleWFC"><p>Argyle Women</p></a></li>        
        <li><a target="_blank" href="https://www.argylearchive.org.uk/"><p>The Argyle Archive</p></a></li>
        <li><a target="_blank" href="https://www.argylefanstrust.com/"><p>Argyle Fans' Trust</p></a></li>
        <li><a target="_blank" href="https://pasoti.co.uk/talk/"><p>PASOTI</p></a></li>
        <li><a target="_blank" href="https://www.greentaverners.co.uk"><p>Green Taverners</p></a></li>
        <li><a target="_blank" href="https://www.facebook.com/groups/1325246504231190"><p>Senior Greens</p></a></li>
        <li><a target="_blank" href="https://www.plymouthherald.co.uk/all-about/plymouth-argyle"><p>Plymouth Live</p></a></li>
        <li><a target="_blank" href="https://www.bbc.co.uk/sport/football/teams/plymouth-argyle"><p>BBC Sport</p></a></li>
        <li><a target="_blank" href="https://www.skysports.com/plymouth-argyle"><p>Sky Sports</p></a></li>
        <li><a target="_blank" href="https://www.newsnow.co.uk/h/Sport/Football/League+One/Plymouth+Argyle?type=ln"><p>News Now</p></a></li>
        <li><a target="_blank" href="http://www.historicalkits.co.uk/Plymouth_Argyle/Plymouth_Argyle.htm"><p>Historical Football Kits</p></a></li>
    </ul>
    
    </div>

  </div>
  
  <div id="box2a">
  			
  		<div id="display2"></div>
         
	  	<p class="potdheading style4boldgrey" style="margin: 8px 0">PICTURES OF THE DAY<br><span class="style2">Click to expand</span></p>
	  
	  	<% response.write(thumb1) %>
	  	<% response.write(thumb2) %>
      
		<% Call CentrePanel %> 
	
  </div>
  
 </div>
  
</div>
  
<div style="clear: both; margin-top:18px;"><br></div>     

<%
conn.close
%>

<!--#include file="base_code.htm"-->

</body>
</html>

<% 
		Sub WhatsNew()
		
		Dim video_this_batch_done
	
		sql = "select cast(publish_timestamp as date) as publish_date, left(cast(publish_timestamp as time),5) as publish_time, publish_timestamp, "
  		sql = sql & "	event_date, event_type, material_type, material_seq, publish_by, updateno, material_details1, material_details2, "
  		sql = sql & "	name_then_short, homeaway, goalsfor, goalsagainst, whatsnew_heading, a.whatsnew_text, NULL as whatsnew_linktext, shortname, "
  		sql = sql & "	case material_type when 'S' then 1 when 'A' then 2 else 3 end as material_order "
		sql = sql & "from event_control a "
		sql = sql & "	left outer join match b on a.event_date = b.date "
		sql = sql & "	left outer join opposition c on b.opposition = c.name_then "
		sql = sql & "	left outer join match_extra d on b.date = d.date "
		sql = sql & "	left outer join contributor e on a.publish_by = e.initials "
		sql = sql & "where event_published = 'Y' "
		sql = sql & " and event_type in ('M','F','O','E','H') "
		sql = sql & "and datediff(day, event_date, getdate()) < " & timespan
	
		sql = sql & "union all "	
		sql = sql & "select cast(publish_timestamp as date), NULL, NULL, NULL, event_type, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, whatsnew_text, NULL, NULL, NULL "
		sql = sql & "from event_control "
		sql = sql & "where event_published = 'Y' "
		sql = sql & " and event_type = 'V' "
		sql = sql & "and datediff(day, cast(publish_timestamp as date), getdate()) < " & timespan

		sql = sql & "union all "
		sql = sql & "select date, left(time,5), NULL, date as event_date, ' ', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, heading, text, linktext, NULL, NULL " 
		sql = sql & "from whatsnew "
		sql = sql & "where datediff(day, date, getdate()) > -1 "
		sql = sql & "   and datediff(day, date, getdate()) < " & timespan
		sql = sql & "order by publish_date desc, event_date desc, event_type, material_order, updateno "
		
		rs.open sql,conn,1,2
					
		Do While Not rs.EOF
		
			if NOT (rs.Fields("event_type") = "V" and video_this_batch_done = rs.Fields("publish_date")) then		'ignore video entries when the first for a publish date has already been processed

				if rs.Fields("event_type") & rs.Fields("event_date") & rs.Fields("publish_date") <> lastpublish then 
			
					'if not the first, finish off the previous entry
					if lastpublish > "" then Call WhatsNewEndEntry
				
				 	lastpublish = rs.Fields("event_type") & rs.Fields("event_date") & rs.Fields("publish_date")
				 	lasteventdate = rs.Fields("event_date")
				 	lasteventtype = rs.Fields("event_type")
				 	lastlinktext = rs.Fields("material_seq")
				 	
				 	lastheading = rs.Fields("whatsnew_heading")
				 	lasttext = rs.Fields("whatsnew_text")
					lastlinktext = rs.Fields("whatsnew_linktext")
				 	
				 	audiodone = ""
				 				
					'start the new entry
					response.write("<li style=""background: #fffff8;"">")
	
					response.write("<p style=""margin-top:4px; font-weight:bold"">") 
					Call WNtag(rs.Fields("publish_date"))
					response.write(weekdayname(weekday(rs.Fields("publish_date"))) & " " & day(rs.Fields("publish_date")) & " " &  monthname(month(rs.Fields("publish_date")),true) & "</p>")
					
					Select Case rs.Fields("event_type")
					
						case "M"	'a conventional first-team game
										
							response.write("<p style=""margin-top:2px; font-weight:bold"">")
							
							if IsNull(rs.Fields("goalsfor")) then	'material has been found before the match details have been added to the match table, so go to season_this instead 
								sql = "select opposition, homeaway "
								sql = sql & "from season_this "
								sql = sql & "where date = '" & rs.Fields("event_date") & "' " 
								rs1.open sql,conn,1,2
								if rs1.RecordCount > 0 then
									if rs1.Fields("homeaway") = "H" then
										response.write("Argyle v " & rs1.Fields("opposition"))
				  	  	  	  	  	  else
										response.write(rs1.Fields("opposition") & " v Argyle")
									end if
								end if	
								rs1.close
							  else
								if rs.Fields("homeaway") = "H" then
									response.write("Argyle " & rs.Fields("goalsfor") & " " & rs.Fields("name_then_short") & " " & rs.Fields("goalsagainst"))
				  	  	  	  	  else
									response.write(rs.Fields("name_then_short") & " " & rs.Fields("goalsagainst") & " Argyle " & rs.Fields("goalsfor"))
								end if
							end if	
							response.write("</p>")	
							
	
						case "E"	'other event
						
							response.write("<p style=""margin-top:2px; font-weight:bold"">" & rs.Fields("whatsnew_heading") & "</p>")	
							response.write("<p style=""margin-top:2px; margin-bottom:3px;"">" & rs.Fields("whatsnew_text") & "</p>")
	
							
						case "F"	'pre-season friendly
						
							response.write("<p style=""margin-top:2px; font-weight:bold"">" & rs.Fields("whatsnew_heading") & "</p>")	
							response.write("<p style=""margin-top:2px; margin-bottom:3px;"">Pictures from this pre-season friendly.</p>")
	
						'case "H"	'Home Park Redevelopment
						
							'response.write("<p style=""margin-top:2px; font-weight:bold"">" & rs.Fields("whatsnew_heading") & "</p>")	
							'response.write("<p style=""margin-top:2px; margin-bottom:3px;"">")
							'if lasteventdate = "2018-01-30" then 
								'response.write("The first ")
							  'else
							  	'response.write("More ")
							'end if
							'response.write("redevelopment pictures from Home Park's south side.</p>")
						
						case "H"	'Home Park Redevelopment
						
							response.write("<p style=""margin-top:2px; font-weight:bold"">" & rs.Fields("whatsnew_heading") & "</p>")	
							response.write("<p style=""margin-top:2px; margin-bottom:3px;"">")
							response.write("The latest redevelopment pictures from Home Park.</p>")

						
						case "O"," "	'a blank indicates a non-match, non-photo related entry on the whatsnew table
											
							response.write("<p style=""margin-top:2px; font-weight:bold"">" & rs.Fields("whatsnew_heading")  & "</p>")
							response.write("<p style=""margin-top:2px;"">" & rs.Fields("whatsnew_text") & " ")
							response.write(rs.Fields("whatsnew_linktext")  & "</p>")
							
						case "V"	'YouTube videos added
											
							response.write("<p style=""margin-top:2px; font-weight:bold"">Videos Found Online</p>")
							sql = "select count(*) as videocount "
							sql = sql & "from event_control "
							sql = sql & "where cast(publish_timestamp as date) = '" & rs.Fields("publish_date") & "' "
							sql = sql & "  and datediff(""dd"", publish_timestamp, getdate()) < " & timespan
							rs1.open sql,conn,1,2
							if rs1.Fields("videocount") > 1 then response.write("<p style=""margin-top:2px;"">" & rs1.Fields("videocount") & " new links on the Videos Found Online page, and also on the associated match pages. ")
							rs1.close
							if not isnull(rs.Fields("whatsnew_text")) then response.write("<p style=""margin-top:2px;"">" & rs.Fields("whatsnew_text")) 'Added comment in the first video for this batch
							video_this_batch_done = rs.Fields("publish_date")							
					
					End Select
			
				end if	
								 
				response.write("<p style=""margin: 2px 0;"">")

				Select Case rs.Fields("event_type")
				
					case "M"	'a conventional first-team game
				
						Select Case rs.Fields("material_type")
							Case "A"
								if audiodone <> "y" then	 response.write("Audio clips from " & rs.Fields("shortname"))
								audiodone = "y"
							Case "I"
								response.write(rs.Fields("material_details1"))
							Case "P"
								response.write("Panorama (")
							'Case "S"
								'response.write("Report from " & rs.Fields("shortname"))
							Case "Y"
								response.write(rs.Fields("material_details2"))
						End Select
					
					case "O","F","E","H"
						response.write("[<a href=""photoarchive.asp"">Non-Match Photos</a>]")
						'response.write("[<a href=""")
						'parm = datepart("yyyy", lasteventdate) & "-" & right("0" & datepart("m", lasteventdate),2) & "-" & right("0" & datepart("d", lasteventdate),2) & rs.Fields("event_type") & rs.Fields("material_seq")
						'response.write("photos.asp?parm=" & parm & """>" & rs.Fields("material_details1") & "</a>]")
						
					'case "H"
					
						'response.write("[<a href=""homepark.asp"">HP Redevelopment</a>]")
						
					case "V"
					
						response.write("[<a href=""youtubeview.asp?parm=1"">Recently Found Videos</a>]")
					
				End select	
				
				response.write("</p>")	
			
			end if		
				
			rs.MoveNext		
						
		Loop

		Call WhatsNewEndEntry	'finish off the last entry
		
		rs.close
		
%><% End Sub %>

<% 
		Sub WhatsNewEndEntry()
			
		Select Case lasteventtype
		
			Case "M"
				response.write("<p style=""margin-top:3px; margin-bottom:4px;"">[<a href=""")
				response.write("gosdb-match.asp?date=" & lasteventdate & """>Here &amp; Now: Match Page</a>]</p>")

		End Select
							
		response.write("</li>")
								
%><% End Sub %>

<% Sub CentrePanel() %>

	<div id="box2a1">
		<ul> 
			<li style="width:90%; margin: 20px 0 12px; padding: 6px 0 5px; color:#101010; text-align:center; border:1px solid #c0c0c0; background: #fffff8;">
			<span class="style4boldgrey">MATCH OF THE DAY: <% Response.Write (" " & Day(thisdate) & " " & UCase(MonthName(Month(thisdate)))) %></span>
        	</li> 		
 		
 			<% Call MotD(thisdate) %> </li>

		</ul>
      </div>
      
      <div id="box2a2">
      
      	<% Call OnThisDay(thisdate) %>

		<ul>
          <li style="width:fit-content; margin: 12px auto 10px; padding: 5px 26px; color:#101010; text-align:center; border:1px solid #c0c0c0; background: #fffff8;">
          <span class="style4boldgrey">BORN THIS DAY</span></li>
			
          <% Call BornToday(thisdate) %>
                  
        </ul>
         
     </div>

</a>

<% End Sub %><% Sub MotD (motddate) %><%

		motdday = day(motddate)
		motdmonth = month(motddate)

		sql = "select a.date, opposition, opposition_qual, name_then_short, lfc, homeaway, goalsfor, goalsagainst, pensfor, pensagainst, "
		sql = sql & "competition, subcomp, headline "
		sql = sql & "from v_match_season a left outer join match_extra b on a.date = b.date join opposition c on a.opposition = c.name_then "
		sql = sql & "where day(a.date) = " & motdday & " and month(a.date) = " & motdmonth & " and motd is not null and len(report) > 100 "
		sql = sql & "order by a.date desc "

		rs.open sql,conn,1,2
		
		motdhold = "<ul>"
	
		Do While Not rs.EOF 
		
			motdfound = "Y"	
       		motdhold = motdhold & "<li>"
       		motdhold = motdhold & "<p class = ""date"">" & FormatDateTime(rs.Fields("date"),1) & "<span style=""margin-left:20px;"" >" & rs.Fields("competition")
       		if not IsNull(rs.Fields("subcomp")) then motdhold = motdhold & " " & trim(rs.Fields("subcomp"))
   			motdhold = motdhold & "</span></h1>"
       		motdhold = motdhold & "<p class=""headline"">" & trim(replace(rs.Fields("headline"),"|","<br>")) & "</p>"
   			
   			if rs.Fields("homeaway") = "H" then
				motdhold = motdhold & "<p class=""score"">Argyle &nbsp;" & rs.Fields("goalsfor") & " - " & rs.Fields("goalsagainst") & "&nbsp; " & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & "</p>"
 		  		if not isnull(rs.Fields("pensfor")) then
		  		   	motdhold = motdhold & "<p class=""penalties"">Penalties: Argyle " & rs.Fields("pensfor") & " - " & rs.Fields("pensagainst") & " " & rs.Fields("name_then_short") & "</p>"
				end if
		  	  else 
				motdhold = motdhold & "<p class=""score"">" & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " &nbsp;" & rs.Fields("goalsagainst") & " - " & rs.Fields("goalsfor") & "&nbsp; Argyle" & "</p>"
				if not isnull(rs.Fields("pensfor")) then
				   	motdhold = motdhold & "<p class=""penalties"">Penalties: " & rs.Fields("name_then_short") & " " & rs.Fields("pensagainst") & " - " & rs.Fields("pensfor") & " Argyle" & "</p>"	
				end if		   		  
			end if
			
			motdhold = motdhold & "<p style=""margin: 2px 0 9px"">[<span id=""" & rs.Fields("date") & """ class=""motdmatch style1green"">See More</span>]</p>"
			
			motdhold = motdhold & "</li>"

		rs.MoveNext
	
		Loop
		
		if motdfound <> "Y" then motdhold = motdhold & "<li><p>No first-team matches have been played on this day</p></li>"

		response.write(motdhold)
	
		rs.close
	
%><% End Sub %><% Sub OnThisDay(thisdate) %><%	

 		Dim texthold
 					
   		sql = "select year, fact "
		sql = sql & "from onthisday "  
		sql = sql & "where month = '" & monthname(month(thisdate),-1) & "' "
		sql = sql & "  and day = " & day(thisdate) & " " 
		sql = sql & "  and seqno < 99 " 
		sql = sql & "order by seqno "
		rs.open sql,conn,1,2
		
		if rs.RecordCount > 0 then
		
			response.write("<ul><li style=""width:fit-content; margin: 15px auto 6px; padding: 5px 26px; color:#101010; text-align:center; border:1px solid #c0c0c0; background: #fffff8;"">")
			response.write("<span class=""style4boldgrey"">ON THIS DAY</span></li>")
		
   			Do While Not rs.EOF	
				if rs.Fields("year") > "" then
					texthold = "<b>" & rs.Fields("year") & ":</b> " & rs.Fields("fact")
				  else 
			  		texthold = rs.Fields("fact")
				end if
			
				if instr(texthold,"^^") > 0 then texthold = replace(texthold, "^^", year(thisdate) - rs.Fields("year"))
	 
				response.write("<li style=""width:99%; position:relative; top:-1px;""><p style=""margin: 0 6; padding: 4 0;"">" & texthold & "</p></li>")
			
				rs.MoveNext
			Loop
			
		end if
		response.write("</ul>")
		rs.close
		
%><% End Sub %><% Sub BornToday(thisdate) %><%	

		Dim texthold, photohold, penpichold, photoname, primephoto, games, goals

   		sql = "select a.player_id, a.forename, a.surname, year(a.dob) as year, a.first_game_year, max(b.last_game_year) as last_game_year, left(a.penpic,160) as penpic, a.prime_photo "
		sql = sql & "from player a left outer join player b on a.player_id = b.player_id_spell1 "  
		sql = sql & "where month(a.dob) = " & month(thisdate) & " "
		sql = sql & "  and day(a.dob) = " & day(thisdate) & " " 
		sql = sql & "  and a.spell = 1 "
		sql = sql & "group by a.player_id, a.forename, a.surname, a.dob, a.first_game_year, a.penpic, a.prime_photo " 
		sql = sql & "order by a.dob "
		rs.open sql,conn,1,2
		
		If rs.EOF then
		
		response.write("<p style=""margin: 6 0; text-align:center; "">We know of no first-team players born on this day.</p>")

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

   		
   			if not IsNull(rs.Fields("prime_photo")) then
				primephoto = "_" & rs.Fields("prime_photo")
	  		else 
				primephoto = ""
			end if
		
			if len(rs.Fields("player_id")) < 4 then 
				photoname = right("00" & rs.Fields("player_id"),3)
	  		else
	  			photoname = rs.Fields("player_id")
			end if
   		
			photohold = "<div style=""clear: both;""></div>"	

			Set f1=Server.CreateObject("Scripting.FileSystemObject")
			if f1.FileExists(Server.MapPath("gosdb/photos/players/" & photoname & primephoto & ".jpg")) then
				if Request.QueryString("mode") = "app" then
         				photohold = photohold & "<img style=""width:100px; margin: 0 6; float: left"" border=""0"" src=""gosdb/photos/players/" & photoname & primephoto & ".jpg"" name=""photo"">"
   				  else
	        			photohold = photohold & "<img style=""width:70px; margin: 0 6; float: left"" border=""0"" src=""gosdb/photos/players/" & photoname & primephoto & ".jpg"" name=""photo"">"
            		end if
    			end if
    				
			penpichold = ""
			if len(rs.Fields("penpic")) > 0 then penpichold = left(rs.Fields("penpic"), instrrev(rs.Fields("penpic")," ")) & " ... <a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id") & """> more</a>"
			penpichold = replace(penpichold,"|p|"," ")	'remove any new paragraph markers in the snippet
			
			texthold = "<p style=""margin: 6 3 0 6;""><b>" & rs.Fields("year") & ":</b> " & rs.Fields("forename") & " " & rs.Fields("surname") & " - "  
		  	texthold = texthold & games & ", " & goals 
		  	
		  	if rs.Fields("last_game_year") = 9999 then
		  		texthold = texthold & " so far."
		  	  elseif rs.Fields("first_game_year") = rs.Fields("last_game_year") then 
		  		texthold = texthold & " in " & rs.Fields("first_game_year") & "." 	
		  	  else texthold = texthold & " between " & rs.Fields("first_game_year") & " and " & rs.Fields("last_game_year") & "."
		  	end if
		  	
		  	texthold = texthold & "</p><p style=""margin: 2 3 0 6;"">" & penpichold & "</p>"
			response.write(photohold & texthold)
			
			rs.MoveNext
		Loop
		end if
		
		rs.close
		
%><% End Sub %><% Sub WNtag(itemdate) %><%	

		response.write("<span class=""WNtag"">" & DateDiff("d", Now(), itemdate) & "</span>")
 		
%><% End Sub 

Sub Notices 
%> 
<!--
    <div id="notice">
    <p style="font-weight:bold; text-align:left;">WALKING WITH WELICAR 
	<p>Whenever you see photos on GoS, more often than not it's "Thanks to Bob". Bob has been an extraordinary support for me and GoS over many years - when you think of the thousands of matchday photos and of course the countless number during the grandstand redevelopment and now the work going on this summer.
	<p>What better way to thank him than to support his selfless efforts now to raise money for the school in Africa that his charity supports. It's all here: <a href="https://www.justgiving.com/fundraising/steppingoutwithwelicar"><u>JustGiving to Welicar</u></a>. Times are hard of course, but please do what you can - it really will be appreciated.
	</div>
 
    <div id="notice">
	<p style="font-weight:bold; text-align:left;">SEASON 1944-1945 
	<p>I appreciate that this will be of little interest to many, but just for the record, some new matches and players have been added to Greens on Screen's underlying database.
	<p>I've always been pleased to say that GoS-DB has a record of every first-team competitive match played by the club, except for one isolated competition, mentioned in the chapter 16 of GoS's History of Argyle (the 1944-45 section).
	<p>Most think that post-war football resumed at Home Park in late August 1945, but in fact the club entered a Football League tournament in the final two months of the 1944-45 season. These were the days of regional football, which took place throughout the Second World War but without Argyle's participation after 1940. Record books say that Argyle's first post-war match was at Southampton in the autumn of 1945 but they did in fact play nine formal games that spring. This was the Football League West Cup, in which the Pilgrims played nine matches in a group format, with the eventual top two (not Argyle) playing for the trophy. The teams were Bristol City, Cardiff City, Swansea Town, Bath City, Lovell's Athletic (from Newport) and Aberamon Athletic, although, it still being wartime, our fixtures against Cardiff were never played and there was only one encounter with Bristol City.
	<p>The reason that GoS-DB has been lacking until now is that I could find no record of those games, even in the official ledgers at the National Football Museum's archive. However, digging deep into the British Newspaper Archive, there is sufficient to add these matches to GoS-DB, although the newspaper reports are short and none provide the Argyle line-up. Some names are mentioned, but of the 99 (9x11) possibilities I can only identify 51, and so to satisfy the database, I have  unfortunately had to create a fictitious player called 'Not Known' to fill in the 48 empty slots.
	<p>Of the 26 players known to be involved in the Football League West Cup, 11 are new to GoS-DB, so they only played in that competition. There might be other new players but I doubt we'll ever know. That leaves 15 already on GoS, so their total appearances (and possibly goals) have now changed. Of these, Ellis Stuttard, who went on to manage the club, is probably the most well known.
	<p>So there is now a new season and the club's all-time match total has increased by nine, with goals scored and conceded increasing too. However, this was a cup tournament, so the Football League (EFL) totals have not changed.
	<p>Finally, a variety of Match Milestones have regretfully changed, such as Argyle's 5000th game in all competitions, which used to be at home to Portsmouth in August 2015 but is now at home to Mansfield in the April before.
	</div>
 
    <div id="notice">
	<p style="font-weight:bold; text-align:left;">SEASON 1944-1945
	<p>Following my announcement a few weeks ago that I shall be hanging up my GoS keyboard in three years' time, I'm delighted to say that trustees of Plymouth Argyle Heritage Archive have unanimously agreed to become its new owners at that time, assuming that over the next few years we can find a way to make it a practical proposition. This is great news, and I'm also heartened that none of us are under any illusion; there's a great deal to be done to make this a practical proposition. Wish us luck!    
	<p><a href="https://argylearchive.org.uk/uncategorized/greens-on-screen/">The Argyle Archive's announcement is here.</a>
	<p>Steve
	</div>
	
    <div id="notice">
	<p style="font-weight:bold; text-align:center;">THE FUTURE OF GREENS ON SCREEN
	<p>The end of a season brings thoughts again about the future of this website, which you might know has been on my mind for many years. Back in 2017 I made an effort to seek a successor (it even made the local papers), but there were no takers. The more I've thought about it since then, the more I realise that it's a huge ask, and I can see very few realistic options. But equally, Steve Dean won't last forever.
	<p>So I have made a decision, and the only way to settle the matter in my head is to make it public. GoS has just seen the end of its 22nd full season, and I will aim for 25 and then stop. Not only does that seem like a good number, but, all being well, I'll be 72, an age when GoS will be far from the exciting hobby of the early days and much more like a ball and chain.
	<p>What will happen then? Unless an unexpected solution comes forth, I will freeze the site but continue to pay the annual fees to the hosting provider for as long as I can, so it will still be available on the internet as a reference for 138 years of Argyle. But sadly, only 138. No more matches, no new players, no further updates of any form.
	<p>Why am I giving three seasons' notice? Perhaps I still hold out hope that a rescuer will come riding over the horizon, but I can't stress enough how enormous the job will be for that person (or group), and that three years to learn the data, code, processes and the underlying technology that I've moulded over 22 years is not that daft. Needless to say, no one will be more delighted than me if a serious commitment emerges. More realistically, I want to use the remaining time to tidy up a great deal within the site, to reduce it down to what I think will be worth freezing (mostly GoS-DB and history of the club), and even to improve matters, such as attempting to finish the history chapters and add to the team photo collection.
	<p>Whatever happens, it won't be for a while, but having thought about it for some years, I now feel the time has come to make a decision. Not good news and I'm sorry about that, but I hope you understand.
	</div>
--><%
End Sub

%>