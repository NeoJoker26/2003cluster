
<%@ Language=VBScript %>
<% Option Explicit %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
   <meta name="Author" content="Trevor Scallan">
   <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>Greens on Screen Daily Diary</title>

<link rel="stylesheet" type="text/css" href="gos2.css">


<style>
<!--
div#diary p { line-height:1.3; text-align:justify; }
-->
   </style>

   
<SCRIPT LANGUAGE="JavaScript">
<!-- Hide from old browsers
 function newWindow(tier) {
 if (document.all)
 var xMax = screen.width, yMax = screen.height;
 else
 if (document.layers)
 var xMax = window.outerWidth, yMax = window.outerHeight;
 else
 var xMax = 640, yMax=480;

 var xOffset = (xMax - 240)/2
 	tierWindow = window.open(tier, 'tierWin', 'width=246,height=472,screenX='+xOffset+',screenY=50,top=50,left='+xOffset+'')
 	tierWindow.focus()
 }
 // -->
   </SCRIPT>

</head>
<body><!--#include file="top_code.htm"-->

<%
Dim conn,sql,rs, rs1 
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

%>

<center>
<h1 style="margin-top: 12px; margin-bottom: 6px"><span style="font-weight: 400">
<font color="#40703F" face="Verdana" style="font-size: 18px">THE DAILY DIARY</font><font size=5 color="#40703F" face="Verdana"> </font></span></h1>
<p style="margin-top: 3; margin-bottom: 0">
<span style="font-size: 11px; font-weight: 700">A Round-up of Argyle News</span></p>
</center>

<div id="diary">
<center>

<table BORDER="0" WIDTH="980" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="6" >

<tr>
<td align="justify" valign="top" width="200" style="border-style: none; border-width: medium">
<p style="margin-top: 6; margin-bottom: 0">
<font style="font-size: 13px; font-weight: 700">
Argyle News Sites:</font><p style="margin-top: 6; margin-bottom: 0; margin-right:20" align="left">
<span style="font-size: 13px">Greens on Screen's Daily Diary is a compilation of 
Argyle news, with help from these and other Argyle-related sites.</font> </span>
<p style="margin-top: 4; margin-bottom: 0">
<span style="font-size: 13px">
<img border="0" src="images/link5.gif"> 
<a href="http://www.pafc.co.uk/">Plymouth Argyle FC</a> </span>

<p style="margin-top: 4; margin-bottom: 0">
<span style="font-size: 13px">
<img border="0" src="images/link5.gif">
<a href="http://www.thisisplymouth.co.uk/plymouthargyle">The Herald</a>
</span>

<p style="margin-top: 4; margin-bottom: 0">
<span style="font-size: 13px">
<img border="0" src="images/link5.gif">
<a href="http://www.thisisdevon.co.uk/plymouthargyle">Western Morning News</a>
</span>

<p style="margin-top: 4; margin-bottom: 0">
<span style="font-size: 13px">
<img border="0" src="images/link5.gif">
<a href="http://www.newsnow.co.uk/newsfeed/?name=Plymouth+Argyle+FC">News Now</a>
</span>
<p style="margin-top: 15; margin-bottom: 0">
<font style="font-size: 13px; font-weight: 700">
On This Day:</font><p style="margin-top: 6; margin-bottom: 0; margin-right:20" align="left">
<span style="font-size: 13px">Also included on the three most recent days, facts from Argyle's history.</font>
</span>

</td>
<td align="justify" valign="top" style="border-style: none; border-width: medium">
  
<%
Dim TodayDate, selectcriteria, archive, entrydate, output, lastmon, lastyear, lastdate, datecount, daydiff, diaryentry, n1, n2, n3, missingdays, lastdiff, showdate, daysuffix, monthdiff, workdate, workyear, workmon 

TodayDate = Date	'coded this way to allow testing for different times of year
if Request.QueryString("todaydate") > "" then TodayDate = Cdate(Request.QueryString("todaydate"))

archive = Request.QueryString("archive")

if archive = "" then 

	selectcriteria = "date <= '" & TodayDate & "' "
	selectcriteria = selectcriteria & "and date >= '1 "
	
	if month(TodayDate) > 1 then
		selectcriteria = selectcriteria & monthname(month(TodayDate)-1,1) & " " & year(TodayDate) & "' " 		'For the the current diary, "month(TodayDate)-1" forces the start to be the previous month
	  else
		selectcriteria = selectcriteria & "Dec" & " " & year(TodayDate)-1 & "' " 		'For the the current diary in January, start from December in the previous year 
	end if  
	
  else
  
  	selectcriteria = "left(datename(month,date),3) = '" & left(archive,3) & "' and year(date) = 20" & mid(archive,4)
  	
end if

datecount = 0

sql = "set dateformat dmy; "
sql = sql & "select date, entry_no, entry_para_no, entry_para "
sql = sql & "from daily_diary " 
sql = sql & "where " & selectcriteria 
sql = sql & "order by date desc, entry_no, entry_para_no "

rs.open sql,conn,1,2
		
Do While Not rs.EOF	
			 
	if rs.Fields("entry_no") = 1 and rs.Fields("entry_para_no") = 1 then	
	
		entrydate = rs.Fields("date")
		daydiff = Datediff("d", entrydate, TodayDate) 
 
		if datecount = 0 and archive = "" then
			missingdays = Datediff("d", entrydate, TodayDate)
			if missingdays > 0 then
				for n2 = 0 to missingdays-1
					if n2 < 3 then
						Call Displaydate(DateAdd("d", - n2, TodayDate))
						Call OnThisDay(n2)
					end if
				next
			 end if
     	
		  elseif datecount > 0 and archive = "" then 
			missingdays = Datediff("d", entrydate, lastdate)
			lastdiff = Datediff("d", lastdate, TodayDate)
			if daydiff-missingdays < 3 then 
				if missingdays > 1 then	
					Call OnThisDay(daydiff-missingdays)  'finish off last good day
					 For n2 = 1 to missingdays-1
						if lastdiff + n2 < 3 then
							Call Displaydate(DateAdd("d", - (lastdiff+n2), TodayDate))	 
							Call OnThisDay(lastdiff+n2)
						end if	
				 	next
				  else 
				  	Call OnThisDay(daydiff-1)	
			 	end if
			 end if	
		end if


		Call DisplayDate(entrydate)
		lastdate = entrydate
		datecount = datecount + 1	
		
	end if	

	if rs.Fields("entry_para_no") = 1 then
		diaryentry = "<p style=""margin: 12px 18px 6px 0;""><img border=""0"" style=""margin: 0 8px 1px 0;"" src=""images/green1.gif"">" & rs.Fields("entry_para") & "</p>"
  	  else
		diaryentry = "<p style=""margin: 3px 18px 0 0;"">"	& rs.Fields("entry_para") & "</p>"
	end if

	output = output & diaryentry

	rs.MoveNext
	
Loop
		
rs.close

conn.close

response.write(output)

%><%'="a" %><%
%>
</td>

<td align="justify" valign="top" width="200" style="border-style: none; border-width: medium">
<p style="margin: 6px 0 0 5px;">
<font style="font-size: 14px; font-weight: 700">
Diary Archive:</font></p>

<div id=diaryarchive>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <td colspan="6"><a href="dailydiary.asp">Current</a></td>
  </tr>

<% 

'create archive links for all diary pages 

monthdiff = DateDiff("m","1 Jul 2003",TodayDate)

output = ""
workdate = TodayDate

dim firstind
firstind = 1
 
for n1 = 1 to monthdiff
	workdate = DateAdd("m",-1,workdate)
	workmon = MonthName(Month(workdate),1)
	workyear = mid(Year(workdate),3) 
		
	if firstind = 1 or workmon = "Dec" then 
		output = output & "<tr><td width=""100%"" colspan=""6""><p style=""margin-top: 6px""><font style=""font-weight: 700"">"
        output = output & "20" & workyear & "</td></tr>"
    end if
    if firstind = 1 or workmon = "Dec" or workmon = "Jun" then output = output & "<tr>"    
	output = output & "<td><a href=""dailydiary.asp?archive=" & LCase(workmon) & workyear & """>" & workmon & "</a></td>"
	if workmon = "Jul" or workmon = "Jan" then output = output & "</tr>"
	firstind = 0	
next

response.write(output)
%>
  </table>
</div>
</td>
</tr>
</table></center>
</div>
<br>

<!--#include file="base_code.htm"--></body>
</html>

<% Sub DisplayDate(showdate) %><%

	select case Day(showdate)
		case 1
			daysuffix = "st"
		case 2
			daysuffix = "nd"
		case 3
			daysuffix = "rd"
		case 21
			daysuffix = "st"
		case 22
			daysuffix = "nd"
		case 23
			daysuffix = "rd"
		case 31
			daysuffix = "st"
		case else
			daysuffix = "th"
	end select
	
	output = output & "<p style=""margin: 18px 18px -3px 0; padding-top: 15px; border-top: 1px solid #c0c0c0; color:#457B44; font-size: 18px;"">"			
				
	if Year(showdate) <> lastyear then 
	 	lastyear = Year(showdate)
	 	lastmon = MonthName(Month(showdate))
	 	output = output & WeekdayName(WeekDay(showdate)) & " " & Day(showdate) & daysuffix & " " & MonthName(Month(showdate)) & " " & Year(showdate) & "</p>" 
	 elseif MonthName(Month(showdate)) <> lastmon then 
	 	lastmon = MonthName(Month(showdate))
	 	output = output & Day(showdate) & daysuffix & " " & MonthName(Month(showdate)) & "</p>" 
	 else
		output = output & Day(showdate) & daysuffix & "</p>" 
	end if
	
%><% End Sub %><% Sub OnThisDay(minusdays) %><%	

 		Dim texthold, adjusteddate
 		
 		adjusteddate = TodayDate - minusdays
 					
   		sql = "select year, fact "
		sql = sql & "from onthisday "  
		sql = sql & "where month = '" & monthname(month(adjusteddate),-1) & "' "
		sql = sql & "  and day = " & day(adjusteddate) & " " 
		sql = sql & "  and seqno < 99 " 
		sql = sql & "order by seqno "
		rs1.open sql,conn,1,2
		
   		if rs1.RecordCount > 0 then
   			
 			output = output & "<p align=""left"" style=""font-size: 13px; margin-top: 9px; margin-bottom: 6px""><b>On This Day:</b></p>"
   		
   			Do While Not rs1.EOF	
		
				if rs1.Fields("year") > "" then
					texthold = "<font color=""#457B44""><b>" & rs1.Fields("year") & ":</font></b> " & rs1.Fields("fact")
		 	 	else 
		  			texthold = rs1.Fields("fact")
				end if
			
				if instr(texthold,"^^") > 0 then texthold = replace(texthold, "^^", year(adjusteddate) - rs1.Fields("year"))
	 
				output = output & "<p align=""left"" style=""font-size: 13px; margin-right: 18; margin-top: 0; margin-bottom: 6px"">" & texthold & "</p>"			
				rs1.MoveNext
			Loop
		
		end if
		rs1.close
				
%><% End Sub %>