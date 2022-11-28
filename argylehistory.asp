

<%@ Language=VBScript %>
<% Option Explicit %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
   <meta name="Author" content="Greens on Screen">
   <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>Greens on Screen: Complete History of Plymouth Argyle</title>

<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--

html, body {
text-align: center;
}

form {padding: 0; margin: 0; }

#container
{
width: 980px;
margin: 10px auto;
border: 0px solid gray;
line-height: 130%
}

#chapters
{
float: left;
width: 220px;
font-size: 12px;
margin: 0;
}


#history
{
float: right;
width: 744px;
}


#historyhead
{
padding: 6px; margin-bottom: 20px;
border: 1px solid gray;
text-align:left;
}

#copyright
{
padding: 6px; margin-bottom: 20px;
border: 1px solid gray;
}

#chapterhead
{
padding: 6px 12px 2px 12px; margin: 0;
line-height: 1.3; 
border: 1px solid gray;
}

#otherchapters
{
text-align:left; 
margin: 0 0 9px 0;
}

#otherchapters td {padding:3px 12px 3px 0; align=left; valign=top; font-size: 12px; font-family: verdana, arial;}

.imageleft
{
clear:left;
float: left;
margin: 4px 12px 8px 0;
background-color: #fff;
padding: 5px;
border-top: 1px solid #999;
border-right: 2px solid #555;
border-bottom: 2px solid #555;
border-left: 1px solid #999; 
}

.imageright
{
clear:right;
float: right;
margin: 4px 0 8px 12px;
background-color: #fff;
padding: 5px;
border-top: 1px solid #999;
border-right: 2px solid #555;
border-bottom: 2px solid #555;
border-left: 1px solid #999; 
}

.imagecenter
{
margin: 18px auto;
background-color: #fff;
padding: 5px;
border-top: 1px solid #999;
border-right: 2px solid #555;
border-bottom: 2px solid #555;
border-left: 1px solid #999;
clear: both;
}

#chapters p {margin: 6px 0; font-family: verdana, arial; text-align: left; }
#history p {margin: 6px 0; font-family: verdana, arial; text-align: justify; font-size: 12px; line-height: 1.4;}
#history td {margin: 3px 5px; font-family: verdana, arial; text-align: left; vertical-align: top; white-space: nowrap; font-size: 12px; line-height: 1.4;}
#history p.heading {margin: 32px 0 10px 0; font-family: verdana, arial; font-size: 14px; font-weight:bold; text-align: left; color: #457B44; }
#history p.heading2 {margin: 36px 0 6px 0; font-family: verdana, arial; font-size: 12px; font-weight:bold; text-align: left; }
#history p.heading3 {margin: 12px 0 6px 0; font-family: verdana, arial; font-size: 12px; font-weight:bold; text-align: left; }
#history p.feedback1 {margin: 12 0 6 12; font-family: verdana, arial; font-size: 14px; font-weight:bold; text-align: left; color: #457B44; }
#history p.feedback2 {margin: 0 0 6px 12px; font-family: verdana, arial; font-size: 11px; font-weight:bold; text-align: left; }
#history p.feedback3 {margin: 0 0 6px 12px; font-family: "Trebuchet MS",verdana, arial; font-size: 13px; text-align: left; }
#history p.list1 {margin: 6px 0 6px 0; padding-left:18px; text-indent:-18px; font-family: verdana, arial; text-align: justify; font-size: 12px; line-height: 1.4;}
#history p.list2 {margin: 6px 0 6px 18px; padding-left:18px; text-indent:-18px; font-family: verdana, arial; text-align: justify; font-size: 12px; line-height: 1.4;}
#history p.inscription {margin: 6px 0; padding:0 144px 0 120px; font-family: verdana, arial; text-align: center; font-size: 12px; line-height: 1.4;}
#history p.inscription2 {margin: 3px 0; padding:0 0 0 120px; font-family: verdana, arial; text-align: left; font-size: 12px; line-height: 1.3;}

#historyhead p.head1{margin: 3px 4px 6px 4px; font-family: verdana, arial; line-height: 1.2; font-size: 18px; font-weight:bold; text-align: left; color: #457B44; }
#historyhead p.head2{margin: 0 4px 6px 4px; font-family: verdana, arial; line-height: 1.2; font-size: 12px; font-weight:bold; text-align: left;  }

#chapterhead p.head1{margin: 3px 0 4px 0; font-family: verdana, arial; line-height: 1.2; font-size: 18px; font-weight:bold; text-align: left; color: #457B44; }
#chapterhead p.head2{margin: 0 0 6px 0; font-family: verdana, arial; line-height: 1.2; font-size: 12px; font-weight:bold; text-align: left; }

div.extract {width: 740px; margin: 18px 0; padding: 12px; background-color: #eee; border: 1px solid black}
#history p.extract { font-family:"Times New Roman",verdana,arial,helvetica,sans-serif; font-size: 14px; text-align: left; line-height: 1.2;}

div.letter {width: 740px; margin: 18px 0; padding: 12px; background-color: #eee; border: 1px solid black}
#history p.letter { font-family:courier,verdana,arial,helvetica,sans-serif; font-size: 14px; text-align: left; line-height: 1.2;}

.bold {font-weight: 700;}
.right {text-align: right !important}
.italic {font-style: italic;}
.clear {clear: both;}
.indent {padding: 0 36px;} 

#chart1
{
margin: 20px 0 10px 0;
position: relative;
}

#chart1 p {position:absolute;top:13; 
		   font-size: 10px; 
		   font-weight:normal; 
		   text-align: left; 
		   color: black; }
	
#chart2
{margin: 0;
position: relative;
}

.letred {position:absolute; font-size:9px; color:red; z-index:10;}
.letblack {position:absolute; font-size:9px; color:black; z-index:10;}
.letgreen {position:absolute; font-size:9px; color:green; z-index:10;}

#chart2 img
{
padding: 0; 
border: 0px none; 
margin-left:0; margin-right:1px; margin-top:0; margin-bottom:0
}

#chart2 p {position:absolute;top:12;
		   font-size: 10px; 
		   font-weight:normal; 
		   text-align: left; 
		   }
		   
#chart2 td {font-size: 11px; 
		   font-weight:normal; 
		   text-align: left; 
		   }

#chart3 {position:relative;
		top:-60;
		left:85		   
		}
		
#chart3 td {font-size: 11px; 
		  font-weight:normal; 
		  text-align: left;
		  padding:2px; 
		  }
		
#chart4 {position:absolute; 
		top:72; 
		left:288; 
		width:280px;
		}
		
#chart4 p, #chart4 td {font-size: 11px; 
		   font-weight:normal; 
		   text-align: left; 
		   }

#printhead {display: none}
#printhead p {font-family: verdana, arial;}
#printhead p.head1{font-family: verdana, arial; margin: 0 0 6pt; font-size: 15pt; font-weight:bold; text-align: left; color: #457B44; }
#printhead p.head2{font-family: verdana, arial; margin: 0 0 6pt; font-size: 12pt; font-weight:bold; text-align: left;  }
#printhead p.head3{margin: 0 0 6pt; font-size: 12pt; text-align: left;  }


@media print {

	#banner, #chapters, #base, .noprint {display: none}

	#printhead {display: block; margin: 30pt 0;}	

	#container {width: auto; font-size: 12pt; margin:0 20pt;}

	#chapterhead p.head1 {line-height: 150%; margin-bottom: 9pt; font-size: 14pt; font-weight: 700}
	#chapterhead p.head2 {line-height: 150%; margin-bottom: 0pt; font-size: 11pt; font-weight: 700}
	
	#history {margin; 0; width: auto; float: none;}
	#history p {margin: 6pt 0; font-family: 'Times New Roman', serif; text-align: justify; font-size: 12pt; line-height: 1.3;}
	#history p.heading {font-size: 14pt;}
	
	p.imageright {font-size: 11pt;}
	p.imagecenter {font-size: 11pt;}
	
}

-->
   </style>
   
</head>
<body><!--#include file="top_code.htm"-->

<% 
Dim fs, f, filename, era, chapter, chapterlist, seasons, season, author, line, output, output1, histpara, i, j, headlinesplit, n1, n2, n3, work1, work2, extractinprogress, letterinprogress
Dim chapterarray(39,2), activechaparray(23), activechaphigh, chapterno, chaptertitle, chapterfulltitle, chapterhold1, chapterhold2, feedbackchapter, formheading, part3940
Dim conn,sql,rs, n, l, r, avgpos, maxpos, minpos, tiers(4)
Dim formnum1, formnum2, formsum

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

era = Request.QueryString("era")

if era <> "feedback" then

	chapterarray(1,0) = "1886-1890"
	chapterarray(1,1) = "C"					'indicates a chapter rather than annex
	chapterarray(1,2) = "In the Beginning"
 
	chapterarray(2,0) = "1890-1895"
	chapterarray(2,1) = "C"
	chapterarray(2,2) = "From Struggle to Demise"

	chapterarray(3,0) = "1895-1899"
	chapterarray(3,1) = "C"
	chapterarray(3,2) = "Rising from the Phoenix"

	chapterarray(4,0) = "1899-1900"
	chapterarray(4,1) = "C"
	chapterarray(4,2) = "The Argyle Athletic Umbrella"

	chapterarray(5,0) = "1900-1901"
	chapterarray(5,1) = "C"
	chapterarray(5,2) = "Home Park Home"

	chapterarray(6,0) = "1901-1902_1"
	chapterarray(6,1) = "C"
	chapterarray(6,2) = "The Argyle Affair"
	
	chapterarray(7,0) = "1901-1902_2"
	chapterarray(7,1) = "C"
	chapterarray(7,2) = "The big boys come to town"

	chapterarray(8,0) = "1902-1903_1"
	chapterarray(8,1) = "C"
	chapterarray(8,2) = "Argyle FC becomes Semi-pro"

	chapterarray(9,0) = "1902-1903_2"
	chapterarray(9,1) = "C"
	chapterarray(9,2) = "The Birth of Plymouth Argyle"

	chapterarray(10,0) = "1902-1903_3"
	chapterarray(10,1) = "C"
	chapterarray(10,2) = "Argyle FC's Final Triumph"
		
	chapterarray(11,0) = "1903-1910"
	chapterarray(11,1) = "C"
	chapterarray(11,2) = "Plymouth Argyle's Early Years"

	chapterarray(12,0) = "1910-1920"
	chapterarray(12,1) = "C"
	chapterarray(12,2) = "Argyle and The Great War"

	chapterarray(13,0) = "1920-1930"
	chapterarray(13,1) = "C"
	chapterarray(13,2) = "Into the Football League"

	chapterarray(14,0) = "1930-1934"
	chapterarray(14,1) = "C"
	chapterarray(14,2) = "Life in the Second Division"
	
	chapterarray(15,0) = "1934-1939"
	chapterarray(15,1) = "C"
	chapterarray(15,2) = "The End of an Era"
	
	chapterarray(16,0) = "1939-1945"
	chapterarray(16,1) = "C"
	chapterarray(16,2) = "Argyle and the Second World War"
	
	chapterarray(17,0) = "1945-1950"
	chapterarray(17,1) = "C"
	chapterarray(17,2) = "From the Ashes"
	
	chapterarray(18,0) = "1950-1953"
	chapterarray(18,1) = "C"
	chapterarray(18,2) = "Into the Fifties"	
	
	chapterarray(19,0) = "1953-1957"
	chapterarray(19,1) = "C"
	chapterarray(19,2) = "From Best to Worst, via Hollywood"

	chapterarray(25,0) = "anx1"
	chapterarray(25,1) = "1"				'annex number
	chapterarray(25,2) = "An Argyle Timeline"

	chapterarray(26,0) = "anx2"
	chapterarray(26,1) = "2"				'annex number
	chapterarray(26,2) = "Facts & Figures"


	for i = 0 to Ubound(chapterarray)
		if chapterarray(i,0) = era then 
			chaptertitle = chapterarray(i,0)
			chaptertitle = replace(chaptertitle,"_"," Part ")
			chapterfulltitle = chapterarray(i,2)
			chapterno = i
			exit for
		end if
	next

	j = 1
	for i = 0 to Ubound(chapterarray)
		if chapterarray(i,0) > "" then 
			activechaparray(j) = i
			activechaphigh = j
			j = j + 1
		end if
	next

end if	
		
	if left(era,3) = "bib" or left(era,2) = "ps" or era="authors" then

		filename = "argylehistory_" & era & ".txt"

	  else
	  
		chapter = 99	'posit a high number - later added: not sure why!
		for n = 1 to Ubound(chapterarray)
			if chapterarray(n,0) = era then
				chapter = n
				filename = "argylehistory" & chapterarray(n,0) & ".txt"
			end if
		next
	
		' list chapters and annexes

		chapterlist = ""
		for n = 1 to int((activechaphigh+1)/2)		'e.g. for 5 active chapters, n would go to 3.
			if n = 1 then
				chapterlist = chapterlist & "<tr><td><b>Other chapters:</b></td>"
		  	else  
		  		chapterlist = chapterlist & "<tr><td>&nbsp;</td>"
			end if 
		
			l = n
			r = n + int((activechaphigh+1)/2)
		
			if activechaparray(l) <> chapterno then
				chapterhold1 = "<a href=""argylehistory.asp?era=" & chapterarray(activechaparray(l),0) & """>" & chapterarray(activechaparray(l),2)
				chapterhold2 = "</a>"
		 	 else
		  		chapterhold1 = chapterarray(activechaparray(l),2)
		  		chapterhold2 = ""
		 	end if	
			if chapterarray(activechaparray(l),1) = "C" then
				chapterlist = chapterlist & "<td>" & activechaparray(l) & ". " & chapterhold1 & " (" & left(chapterarray(activechaparray(l),0),5) & replace(mid(chapterarray(activechaparray(l),0),8),"_"," P") & ")" & chapterhold2 & "</td>"
		  	 else
				chapterlist = chapterlist & "<td>Annex " & chapterarray(activechaparray(l),1) & ". " & chapterhold1 & "</td>"
			end if 		
		
			if r <= activechaphigh then
				if activechaparray(r) <> chapterno then
					chapterhold1 = "<a href=""argylehistory.asp?era=" & chapterarray(activechaparray(r),0) & """>" & chapterarray(activechaparray(r),2)
					chapterhold2 = "</a>"
		  	  	 else
		  			chapterhold1 = chapterarray(activechaparray(r),2)
		  			chapterhold2 = ""
		  		end if
				if chapterarray(activechaparray(r),1) = "C" then
					chapterlist = chapterlist & "<td>" & activechaparray(r) & ". " & chapterhold1 & " (" & left(chapterarray(activechaparray(r),0),5) & replace(mid(chapterarray(activechaparray(r),0),8),"_"," P") & ")" & chapterhold2 & "</td>"
			   	 else
					chapterlist = chapterlist & "<td>Annex " & chapterarray(activechaparray(r),1) & ". " & chapterhold1 & "</td>"
			 	end if 	
		  	else 
		  		chapterlist = chapterlist & "<td></td>" 
			end if
		
			chapterlist = chapterlist & "</tr>"
		next
	end if
	
if era = "anx1" then
	feedbackchapter = "Annex 1"
'build annex 1
%>
	<!--#include file="argylehistory_annex1.asp"-->
<%
elseif era = "anx2" then
	feedbackchapter = "Annex 2"
'build annex 2
%>
	<!--#include file="argylehistory_annex2.asp"-->
<%
else
' not an annex	

	if era = "ps1" then 
		output = "<p class=""head1"">ARGYLE FC POSTSCRIPT</p>"
	 elseif left(era,3) = "bib" then 
		output = "<p class=""head1"">BIBLIOGRAPHY</p>"
	 elseif era = "authors" then 
		output = "<p class=""head1"">AUTHORS</p>"
	 else	
	 	feedbackchapter = "Chapter " & chapter & ": " & chaptertitle
		output = "<p class=""head1"">" & feedbackchapter & "</p>"
		output = output & "<p class=""head1"">" & chapterfulltitle & "</p>"
	end if

	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set	f=fs.OpenTextFile(Server.MapPath("historytext/" & filename),1)
	
	Do While Not f.AtEndOfStream
	line=f.Readline
	
	if left(line,9) = "|comment|" then
	
	elseif left(line,8) = "|author|" then
		author = mid(line,9)
		if instr(author," and ") > 0 then
			output = output & "<p class=""head2"">Authors: "
	      else
			output = output & "<p class=""head2"">Author: "
		end if
		output = output & author & "&nbsp;&nbsp; <span class=""noprint"" style=""font-size:12px; font-weight:normal;"">[<a href=""argylehistory.asp?era=authors"">about the authors</a>]</span></p>"
	
	elseif left(line,9) = "|preface|" then
		output = output & "<p>" & mid(line,10) & "</p>"
		
	elseif mid(line,4,9) = "|preface|" then
		' NB mid(line,4,9) is to get over a problem of rogue unicode character ﻿ appearing at the start of the line
		output = output & "<p>" & mid(line,13) & "</p>"
	
	elseif left(line,11) = "|headlines|" then
		headlinesplit = split(mid(line,12)," ... ")
		for i = 0 to UBound(headlinesplit)
   			output1 = output1 & " ... <a href=""#" & i & """>" & headlinesplit(i) & "</a>" 
		next
		output = output & "<p><b>In this chapter:</b> " & mid(output1,6) & "</p>"		'remove first three dots and finish off the div
		
	elseif left(line,9) = "|version|" then
		output = output & "<p><b>Version:</b> " & mid(line,10) & "</p>"
		
	elseif left(line,18) = "|acknowledgements|" then
		output = output & "<p class=""noprint""><b>Grateful thanks:</b> " & mid(line,19) & "</p>"

	elseif left(line,12) = "|versionlog|" then
		output = output & "<p class=""noprint"" style=""margin:-3 0 6 0; font-size: 11px; line-height: 1.2;"">" & mid(line,13) & "</p>"
	
	elseif left(line,6) = "|date|" then
		output = output & "<p><b>Date:</b> " & mid(line,7) & "</p>"
	
	elseif left(line,9) = "|sources|" then
		output = output & "<p class=""noprint""><b>Sources:</b> <a href=""#" & i & """>" & mid(line,10) & "</a></p>"	
		
	elseif left(line,14) = "|endheadlines|" then
		output = output & "<div id=""otherchapters"" class=""noprint"">"
		output = output & "<table style=""border-collapse:collapse"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & chapterlist	
		output = output & "</table></div>"
		
		output = output & "<p class=""noprint"" style=""margin:0 0 6 0;""><a href=""argylehistorymenu.asp"">Return to History Contents</a></p>"
		output = output & "</div>"		
		i = 0
		
	elseif left(line,9) = "|heading|" then
		if extractinprogress = 1 then
			output = output & "</div>"
			extractinprogress = 0
		end if
		if letterinprogress = 1 then
			output = output & "</div>"
			letterinprogress = 0
		end if
		histpara = replace(line,"|heading|","<p class=""heading"">")
		output = output & "<a name=""" & i & """>" & histpara & "</a></p>"
		i = i + 1
		'Check if heading is for a new season
		if mid(line,10,2) = "18" or mid(line,10,2) = "19" then
			if mid(line,10,4) < "1895" or (mid(line,10,4) > "1896" and mid(line,10,4) < "1903") then
				season = mid(line,12,2) & mid(line,15,2)
				output = output & "<div class=""imageright noprint"" style=""padding: 0 10 0 10; margin-top: 6; margin-bottom: 6; text-align: left;"">" 
				output = output & "<p>Stats for " & mid(line,10,7) & ": "
				output = output & "<a href=""argylehistory_amateurstats.asp?stat=" & season & "F""><u>First XI</u></a> - "
				if mid(line,10,4) < "1891" then output = output & "<a href=""argylehistory_amateurstats.asp?stat=" & season & "S""><u>Second XI</u></a> - "
				if mid(line,10,4) > "1896" and mid(line,10,4) < "1903" then output = output & "<a href=""argylehistory_amateurstats.asp?stat=" & season & "R""><u>Reserve XI</u></a> - "
				if mid(line,10,4) > "1899" and mid(line,10,4) < "1903" then output = output & "<a href=""argylehistory_amateurstats.asp?stat=" & season & "WF""><u>Wed XI</u></a> - "				
				if mid(line,10,4) = "1900" then output = output & "<a href=""argylehistory_amateurstats.asp?stat=" & season & "WA""><u>Wed A XI</u></a> - "				
				output = output & "<a href=""argylehistory_amateurstats.asp?stat=" & season & "P""><u>Players</u></a>"
				output = output & "</p></div>"
			end if
		end if
				
	elseif left(line,9) = "|subhead|" then
		output = output & replace(line,"|subhead|","<p class=""heading3"">") & "</p>"
	
	elseif left(line,13) = "|seasonstats|" then
				output = output & "<div class=""imagecenter noprint"" style=""padding: 0 10px; margin: 18px 20px; text-align: left;"">" 
				output = output & "<p style=""font-size: 11px; margin: 6 0 8 0; line-height: 1.2;"">Links to GoS-DB for this period: " 
				
				seasons = split(mid(line,14),",")
				
				part3940 = ""
				if uBound(seasons) = 0 and left(seasons(0),9) = "1939-1940" and mid(seasons(0),10,1) = "*" then part3940 = mid(seasons(0),11)
			 	
			 	for each season in seasons
			 		if part3940 = "" then
			 			output = output & "<a href=""gosdb-season.asp?years=" & season & """>" & season & " Results and Table</a> | "
			 		  else	
			 			output = output & "<a href=""gosdb-season.asp?years=" & left(season,9) & "&part=" & part3940 & "&tab=" & part3940 & """>" & left(season,9) & " (" & part3940 & ") Results and Table</a> | " 
					end if
				next
				
				%><!--#include file="conn_read.inc"--><%

				sql = "select rtrim(isnull(forename, isnull(initials,'?'))) + ' ' + rtrim(surname) as name, a.player_id_spell1, count(*) as goals "
				sql = sql & "from player a join match_player b on a.player_id = b.player_id " 
				sql = sql & "join season on date >= date_start and date <= date_end "
				if part3940 = "D2" then
					sql = sql & "where years = '1939-1940' "
					sql = sql & "  and date <= '1939-09-02' "
				  elseif part3940 = "SWRL" then
				  	sql = sql & "where years = '1939-1940' "
					sql = sql & "  and date > '1939-09-02' "
				  else
				  	sql = sql & "where years between '" & seasons(0) & "' and '" & seasons(ubound(seasons)) & " |' "	
				end if	
				sql = sql & "group by forename, initials, surname, a.player_id_spell1 "
				sql = sql & "order by surname, initials "
				
				rs.open sql,conn,1,2
 				
 				Do While Not rs.EOF
  					output = output & "<a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & """>" & rs.Fields("name") & "</a> | " 
				 	rs.MoveNext
 				Loop
 				
				rs.close 
				conn.close
				output = left(output,len(output)-3) & "</p></div>"
				
	elseif left(line,11) = "|imageleft|" then
		if extractinprogress = 1 then
			output = output & "</div>"
			extractinprogress = 0
		end if
		if letterinprogress = 1 then
			output = output & "</div>"
			letterinprogress = 0
		end if
		work1 = split(line & "width=0 ","width=",2)
		work2 = split(work1(1)," ")
		output = output & "<div class=""imageleft"" style=""width:" & work2(0)+13 & "px"">" & replace(line,"|imageleft|","")	
						
	elseif left(line,12) = "|imageright|" then
		if extractinprogress = 1 then
			output = output & "</div>"
			extractinprogress = 0
		end if
		if letterinprogress = 1 then
			output = output & "</div>"
			letterinprogress = 0
		end if
		work1 = split(line & "width=0 ","width=",2)
		work2 = split(work1(1)," ")
		output = output & "<div class=""imageright"" style=""width:" & work2(0)+13 & "px"">" & replace(line,"|imageright|","")	
			
	elseif left(line,13) = "|imagecenter|" then
		if extractinprogress = 1 then
			output = output & "</div>"
			extractinprogress = 0
		end if
		if letterinprogress = 1 then
			output = output & "</div>"
			letterinprogress = 0
		end if
		work1 = split(line & "width=0 ","width=",2)
		work2 = split(work1(1)," ")
		output = output & "<div class=""imagecenter"" style=""width:" & work2(0)+13 & "px"">" & replace(line,"|imagecenter|","")	
			
	elseif left(line,9) = "|caption|" then
		output = output & replace(line,"|caption|","<p style=""font-size: 11px; margin: 4px 3px 2px 6px; line-height: 1.2; text-align:left;"">") & "</p>"
		
	elseif left(line,15) = "|captioncenter|" then
		output = output & replace(line,"|captioncenter|","<p style=""font-size: 11px; font-style: italic; margin: 4px 3px 2px 4px; text-align:center;"">") & "</p>"

	elseif left(line,10) = "|endimage|" then
		output = output & "</div>"
		
	elseif left(line,15) = "|sourceheading|" then
		histpara = "<a name=""" & i & """>" & replace(line,"|sourceheading|","<p class=""heading2"">")	
		output = output & "<a name=""" & i & """>" & histpara & "</a></p>"
		
	elseif left(line,12) = "|sourcehead|" then
		output = output & replace(line,"|sourcehead|","<p class=""heading3"">")

	elseif left(line,7) = "|quote|" then
		output = output & replace(line,"|quote|","<p class=""italic"">")
		
	elseif left(line,13) = "|extracthead|" then
		if extractinprogress = 0 then
			output = output & "<div class=""extract"">"
			extractinprogress = 1
		end if
		output = output & replace(line,"|extracthead|","<p class=""extract bold"">") & "</p>"

	elseif left(line,9) = "|extract|" then
		if extractinprogress = 0 then
			output = output & "<div class=""extract"">"
			extractinprogress = 1
		end if
		output = output & replace(line,"|extract|","<p class=""extract"">") & "</p>"

	elseif left(line,8) = "|letter|" then
		if letterinprogress = 0 then
			output = output & "<div class=""letter"">"
			letterinprogress = 1
		end if
		output = output & replace(line,"|letter|","<p class=""letter"">") & "</p>"
				
	elseif left(line,12) = "|endchapter|" then
		'formheading = "<p class=""noprint"" style=""margin: 0 0 6 0;"">Feel free to leave your thoughts about this chapter by completing the boxes below. If you have a general comment to make about the history project as a whole, please use the feedback facility on the <a href=""argylehistorymenu.asp""><u>contents page</u></a>.</p>"
		%>
		<!--#include file="argylehistory_form.asp"-->
		<%

	elseif left(line,17) = "|feedbackchapter|" then
		output = output & replace(line,"|feedbackchapter|","<p class=""feedback1"">") & "</p>"	
		
	elseif left(line,18) = "|feedbackdatetime|" then
		output = output & replace(line,"|feedbackdatetime|","<p class=""feedback2"">") & "</p>"
		
	elseif left(line,14) = "|feedbackname|" then
		output = output & replace(line,"|feedbackname|","<p class=""feedback2"">") & "</p>"

	elseif left(line,17) = "|feedbackcomment|" then
		output = output & replace(line,"|feedbackcomment|","<p class=""feedback3"">") & "</p>"
	
	elseif left(line,13) = "|feedbackend|" then
		output = output & replace(line,"|feedbackend|","<br>")
	
	elseif left(line,7) = "|list1|" then
		output = output & replace(line,"|list1|","<p class=""list1"">") & "</p>"
	
	elseif left(line,7) = "|list2|" then
		output = output & replace(line,"|list2|","<p class=""list2"">") & "</p>"
		
	elseif left(line,13) = "|inscription|" then
		output = output & replace(line,"|inscription|","<p class=""inscription"">") & "</p>"
				
	elseif left(line,14) = "|inscription2|" then
		output = output & replace(line,"|inscription2|","<p class=""inscription2"">") & "</p>"

	elseif left(line,7) = "|table|" then
		output = output & "<table style=""float:left; margin-bottom:20px;"">"

	elseif left(line,10) = "|endtable|" then
		output = output & "</table>"
		
	elseif left(line,4) = "|h1|" then
		output = output & replace(line,"|h1|","<tr><td style=""font-weight: bold"">")

	elseif left(line,4) = "|hn|" then
		output = output & replace(line,"|hn|","<td style=""font-weight: bold"">")
		
	elseif left(line,9) = "|hnright|" then
		output = output & replace(line,"|hnright|","<td style=""font-weight: bold; text-align: right;"">")

	elseif left(line,8) = "|hnwrap|" then
		output = output & replace(line,"|hnwrap|","<td style=""font-weight: bold; white-space: normal;"">")
	
	elseif left(line,4) = "|t1|" then
		output = output & replace(line,"|t1|","<tr><td>")

	elseif left(line,4) = "|tn|" then
		output = output & replace(line,"|tn|","<td>")
		
	elseif left(line,9) = "|tnright|" then
		output = output & replace(line,"|tnright|","<td style=""text-align: right;"">")

	elseif left(line,8) = "|tnwrap|" then
		output = output & replace(line,"|tnwrap|","<td style=""white-space: normal;"">")

	elseif left(line,10) = "|skipline|" then
		output = output & replace(line,"|skipline|","<br>")
								
	else
		if extractinprogress = 1 then
			output = output & "</div>"
			extractinprogress = 0
		end if
		if letterinprogress = 1 then
			output = output & "</div>"
			letterinprogress = 0
		end if
		output = output & replace(line,"|p|","<p>") & "</p>"
			
	end if	
		 
	loop

f.close
Set f=Nothing
Set fs=Nothing

end if
%>

<div id="container">

<div id="chapters">
<div id="historyhead">
<p class="head1">THE HISTORY<br>
OF ARGYLE</p>
<p class="head2" style="margin: 6 4 4 4; line-height: 150%; ">An original, comprehensive 
and thoroughly researched account of Plymouth Argyle Football Club from its earliest roots to the present day.</p>

<span xmlns:dc="http://purl.org/dc/elements/1.1/" href="http://purl.org/dc/dcmitype/Text" property="dc:title" rel="dc:type">
<p style="margin: 12 4 4 4">
<a rel="license" href="http://creativecommons.org/licenses/by-nc-nd/3.0/">
<img alt="Creative Commons Licence" style="border-width:0; margin-left:12; margin-right:12" src="http://i.creativecommons.org/l/by-nc-nd/3.0/88x31.png" vspace="4" align="right" /></a></span><span style="font-size: 11px" xmlns:dc="http://purl.org/dc/elements/1.1/" href="http://purl.org/dc/dcmitype/Text" rel="dc:type"><b>Important copyright conditions:</b>
</span><span xmlns:dc="http://purl.org/dc/elements/1.1/" href="http://purl.org/dc/dcmitype/Text" property="dc:title" rel="dc:type">
<p style="margin: 4 4 6 4; font-size: 11px">This chapter is licenced under a
<b> <a rel="license" href="http://creativecommons.org/licenses/by-nc-nd/3.0/">
<u>Creative Commons Licence</u></a>.</b>
<%
if author > "" then
	response.write("Attribution must include the Author (" & author & ") and this site's address (www.greensonscreen.co.uk)")
  else
	response.write("Attribution must include this site's address (www.greensonscreen.co.uk)")
end if 
%> 
 and must be displayed prominently, in close proximity to any associated material, and be implemented with strict 
regard to the <b><u><a href="http://creativecommons.org/licenses/by-nc-nd/3.0/">licence 
conditions</a></u></b>.</span></p>
</div>

<p style="margin-top: 18; margin-bottom: 10; margin-right:12; line-height:1.3;"><b>
Parts too small for comfortable reading?</b> Most browsers allow you to zoom. Try Ctrl+ (hold down the Ctrl 
key, then press the + 
key, without Shift). <br>
Ctrl- reduces the size.</p>
<p style="margin-right:12; line-height:1.3;"><b>Have you new material to offer?</b> 
Please get in touch by writing to Steve using the 'Contact Us' button at the 
top-right of the page.</p>
<p style="margin-right:30; line-height:1.3;">
<b>Photos used on this page:</b>
Greens on Screen is run as a 
service to fellow supporters, in all good faith, without commercial or private 
gain. I have no wish to abuse copyright regulations and apologise unreservedly 
if this occurs. If you own any of the material used on this page, and object to 
its inclusion, please get in touch using the 'Contact Us' button at the top-right of 
the page.</p>
</div>

<div id="history">
<div id="printhead">
	<p class="head1">THE HISTORY OF ARGYLE</p>
	<p class="head2">An original account of Plymouth Argyle Football Club from its earliest roots to the present day</p>
    <p class="head3">This is a printed representation of one chapter of GoS's History of Argyle (www.greensonscreen.co.uk/argylehistorymenu.asp), provided for ease of reading and personal retention. Inevitably it lacks links to associated pages, including match and player records, and its layout has been simplified to allow page breaks. Note also that Greens on Screen's online History of Argyle will be updated and new material added from time to time.</p> 
	<p class="head3"><b>COPYRIGHT:</b> the strict conditions for use of this printed version are the same for the corresponding online page, as specified on that page. </p>
</div>

<div id="chapterhead">
<%
response.write(output)
%><%'="a" %><%
%>
</div>

</div> 

<div class="clear"></div>

<br>
<!--#include file="base_code.htm"-->
</body></html>