<%@ Language=VBScript %>
<% Option Explicit %>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta name="GsmallheadENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Greens on Screen</title>

<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
select,input {font-size: 12px; margin:0 1px; padding: 0 2px;}
#table {margin: 15px auto;}
#table ul {margin: 3px 0; padding: 0;}
#table li {display: inline-block; border: 1px solid #c0c0c0; padding: 1px 0; margin: 0 1px; color: #202020; font-size: 11px;}
#table li a {padding: 1px 6px}
.rowhlt	{
    /* highlighted row */
    background-color: #ecf4ec; 
    }
-->
</style>

</head>

<body><!--#include file="top_code.htm"-->

<%
Dim table(30,60), latest_table, frommatch, target_date, request_year, request_mon, request_day, prev_date, next_date 
Dim thisdate_message, match_ind, ordertype, division, divisionpart, divisionparts, divisionshort 
Dim winpoints, season_years, date_start, pos_promote, pos_promote_playoff, pos_relegate_playoff, pos_relegate
Dim match_date, mon, months, hometeam, awayteam, homegoals, awaygoals, attendance, adjustment_reason, reason_marker
Dim n, n0, n1, n2, n3, output, topnote, totalmatches, totalattendance, headseason, headdate1, headdate2, selected, adjustind
Dim sort, lowtable, hightable, step, line, inverseteam, inversework, trclass, tablemessage, height, formgif, temp

Dim conn,sql,rs 

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%


' Initialise array 
for n1 = 0 to UBound(table,1)
	for n2 = 1 to UBound(table,2)
		table(n1,n2) = 0   'initialise each count
	next
	table(n1,20) = ""      'don't initialise the form entry with a zero
	table(n1,16) = 999999  'set with high value
next
	
' This page produces a league table for a supplied date (any date since Argyle joined the Football League in 1920).
' The day can be supplied in three aways:
' 1. as form values for day, month and year. This comes from the selection option on this page.
' 2. as a URL parameter (date=yyyy-mm-dd). This comes from an option on the match page.
' 3. no date, or an invalid URL parameter, in which case the most recent table is shown.

if Request.Form("Y") > "" then 
	target_date = request.form("Y") & "-" & request.form("M") & "-" & request.form("D")
  elseif Request.QueryString("date") > "" then
	target_date = Request.QueryString("date")
	if Request.QueryString("source") =  "matchpage" then frommatch = "y"
  else
	target_date = year(Date) & "-" & month(Date) & "-" & day(Date)	'No date supplied so use today's date for now
	latest_table = "Y"
end if

if not IsDate(target_date) then
	target_date = year(Date) & "-" & month(Date) & "-" & day(Date)	'The supplied date is not valid so use today's date for now
	topnote = "<p class=""style1boldred"" style=""margin: 0 0 18px"">The date you have chosen is not valid - please try again</p>"
end if

if dateDiff("d", target_date, "1920-08-28") > 0 then target_date = "1920-08-28"

' If a date supplied, save the components for later
if latest_table = "" then
	temp = split(target_date,"-")
	request_year = temp(0)
	request_mon = temp(1)
	request_day = temp(2)
end if

match_ind = ""

' Check if matches were played on target_date
sql = "select max(date) as near_date "
sql = sql & "from FL_results " 
sql = sql & "where date <= '" & target_date & "' "
rs.open sql,conn,1,2
	if rs.RecordCount = 1 and rs.Fields("near_date") < target_date then 
		match_date = rs.Fields("near_date")	'target_date is not a matchday for any team in the division, so store the nearest date
		match_ind = "1"
	  else
	  	match_date = target_date			'yes, it's a match date
	end if
rs.close

'Having established the real match date, now find the previous and next match dates, along with the details of Argyle's game on this date (if played)
prev_date = ""
next_date = ""

sql = "select 'this' as ind, date, opposition, homeaway, goalsfor, goalsagainst "
sql = sql & "from v_match_FL " 
sql = sql & "where date = '" & target_date & "' "
sql = sql & "union all "
sql = sql & "select 'prev' as ind, max(date) as date, NULL, NULL, NULL, NULL "
sql = sql & "from FL_results " 
sql = sql & "where date < '" & match_date & "' "
sql = sql & "union all "
sql = sql & "select 'next' as ind, min(date) as date, NULL, NULL, NULL, NULL "
sql = sql & "from FL_results " 
sql = sql & "where date > '" & match_date & "' "
rs.open sql,conn,1,2

Do While Not rs.EOF
	if rs.Fields("ind") = "prev" then 
		prev_date = rs.Fields("date")	'save the date for the previous match played by any team
	  elseif rs.Fields("ind") = "next" then
		next_date = rs.Fields("date")	'save the date for the next match played by any team
	  else 
	  	if latest_table = "" then
	  		thisdate_message = "<p class=""style1"" style=""margin: 0 0 12px"">Our match on this day:<br>"
	  		if rs.Fields("homeaway") = "H" then 
	  			thisdate_message = thisdate_message & "Argyle " & rs.Fields("goalsfor") & " " & rs.Fields("opposition") & " " & rs.Fields("goalsagainst")
	  	  	  else
	  	  		thisdate_message = thisdate_message & rs.Fields("opposition") & " " & rs.Fields("goalsagainst") & " " & "Argyle " & rs.Fields("goalsfor") 
	  		end if
	  		thisdate_message = thisdate_message & " (<a href=""gosdb-match.asp?date=" & match_date & """>match details</a>)</p>"
	  		match_ind = "2"
	  	end if 	  	
	end if	  
	rs.Movenext
Loop
rs.close

if latest_table = "" and match_ind = "" then thisdate_message = "<p class=""style1"" style=""margin: 0 0 12px"">Argyle did not play on this day</p>"
if latest_table = "" and match_ind = "1" then thisdate_message = "<p class=""style1"" style=""margin: 0 0 12px"">No league matches were played on this day</p>"
	
winpoints = 3
if match_date < "1981-08-29" then winpoints = 2 	'This was the date that points for a win changed from 2 to 3

'Set the normal league order (points) if a sort value has not been passed with the url 
sort = Request.QueryString("sort")
if sort = "" then sort = 9

' Retrieve the key values for the associated season
sql = "select years, date_start, division, division_short, pos_promote, pos_promote_playoff, pos_relegate_playoff, pos_relegate "
sql = sql & "from season "
sql = sql & "where date_start <= '" & match_date & "' "
sql = sql & "  and date_end >= '" & match_date & "' "
rs.open sql,conn,1,2

season_years = rs.Fields("years")
date_start = rs.Fields("date_start") 
pos_promote = rs.Fields("pos_promote")
pos_promote_playoff = rs.Fields("pos_promote_playoff")
pos_relegate_playoff = rs.Fields("pos_relegate_playoff")
pos_relegate = rs.Fields("pos_relegate")

divisionshort = trim(rs.Fields("division_short"))
division = ""
divisionparts = split(lcase(rs.Fields("division"))," ")
for each divisionpart in divisionparts
	division = division & ucase(left(divisionpart,1)) & right(divisionpart,len(divisionpart)-1) & " "
next

rs.close

ordertype = 2

if latest_table = "Y" and topnote = "" then		
	topnote = "<p class=""style1"" style=""margin: 0 0 12px"">This is the current League table for our club. For another of over 7,900 tables since 1920, select a date above.</p>"	
  elseif topnote = "" then
	topnote = "<p class=""style1"" style=""margin: 0 0 12px"">For teams on equal points, positions in " & season_years & " were determined by "
	select case true
		case match_date > "" and match_date < "1976-07-01"
			ordertype = 1
			topnote = topnote & "GA (Goal Average)</p>" 
		case match_date > "" and match_date < "1992-07-01"
			ordertype = 2
			topnote = topnote & "GD (Goal Difference)</p>" 
		case match_date > "" and match_date < "1999-07-01"
			ordertype = 3
			topnote = topnote & "GF (Goals For)</p>" 
		case else
			ordertype = 2
			topnote = topnote & "GD (Goal Difference)</p>" 
	end select
end if


' Prepare an array for later reference

n0 = 0
totalmatches = 0
totalattendance = 0

sql = "select date, home_team, away_team, home_goals, away_goals, isnull(attendance,0) as attendance "
sql = sql & "from FL_results  "
sql = sql & "where date >= '" & date_start & "' and date <= '" & match_date & "' "
sql = sql & "order by date "

rs.open sql,conn,1,2

if rs.RecordCount = 0 then

	rs.close

	'No rows found - it must be just ahead of the start of the new season ...
	' ... so find the teams from the new fixtures in the season_this table
	
	sql = "select distinct opposition "
	sql = sql & "from season_this a join competition b on a.compcode = b.compcode "
	sql = sql & "where a.compcode = 'F' "
	
	rs.open sql,conn,1,2
	
	'First, put Argyle in the array (they won't be in the fixture table)
	table(0,0) = "Plymouth Argyle"
	
	'Now we fill the table with the latest season's opposition
	n1 = 1
	Do While Not rs.EOF
		table(n1,0) = rs.Fields("opposition")
		n1 = n1 + 1
		rs.Movenext
	loop

  else
  
  	'Rows have been found in the results table, so the date is after the start of the corresponding season 
  
  	Do While Not rs.EOF

		match_date = rs.Fields("date")
		hometeam = rs.Fields("home_team")
		awayteam = rs.Fields("away_team")
		homegoals = rs.Fields("home_goals")
		awaygoals = rs.Fields("away_goals")

		if rs.Fields("attendance") > 0 then	'avoid increasing total matches if the attendance is 0 (i.e. attendance missing, so average is only for all available attendances)		
			totalmatches = totalmatches + 1
			totalattendance = totalattendance + rs.Fields("attendance")
		end if
		
		'process home team
   			For n1 = 0 to UBound(table,1)
  	   			if hometeam = Replace(table(n1,0),"*","") then  	'check team name, ignoring * (deducted points indicator)
	    			exit for
	   			end if
	   			if table(n1,0) = "" then 
	   				table(n1,0) = hometeam  					'set team name
	    			exit for
	 			end if
    		next
     		if homegoals > awaygoals then 
	   			table(n1,2) = table(n1,2) + 1  					'increment home wins
				table(n1,13) = table(n1,13) + winpoints			'increment points
				table(n1,20) = right(table(n1,20) & "7",10) 	'increment form list - a string of up to 10 chars, each position containing: 7 (H win), 5 (H draw), 4 (H loss), 3 (A win), 1 (A draw), 0 (A loss) 
	   			  elseif homegoals = awaygoals then 
	   			    table(n1,3) = table(n1,3) + 1  				'increment home draws
	   			    table(n1,13) = table(n1,13) + 1 			'increment points
					table(n1,20) = right(table(n1,20) & "5",10) 'increment form list - a string of up to 10 chars, each position containing: 7 (H win), 5 (H draw), 4 (H loss), 3 (A win), 1 (A draw), 0 (A loss)	   			   
	   			  else 
	   			  table(n1,4) = table(n1,4) + 1					'increment home defeats
	   			  table(n1,20) = right(table(n1,20) & "4",10) 	'increment form list - a string of up to 10 chars, each position containing: 7 (H win), 5 (H draw), 4 (H loss), 3 (A win), 1 (A draw), 0 (A loss)
	   			end if
	   		table(n1,1) = table(n1,1) + 1		 	 			'increment home played
	   		table(n1,5) = table(n1,5) + homegoals    			'increment home goals for
	   		table(n1,6) = table(n1,6) + awaygoals    			'increment home goals against
	   		table(n1,14) = table(n1,14) + rs.Fields("attendance")  			'increment home attendance
	   		if rs.Fields("attendance") > 0 then table(n1,29) = table(n1,29) + 1 'increment count of published attendances
	   		if rs.Fields("attendance") > 0 and int(table(n1,15)) < int(rs.Fields("attendance")) then table(n1,15) = rs.Fields("attendance") 'replace highest home attendance
	   		if rs.Fields("attendance") > 0 and int(table(n1,16)) > int(rs.Fields("attendance")) then table(n1,16) = rs.Fields("attendance") 'replace lowest attendance
    		
		'now process away team
   			For n1 = 0 to UBound(table,1)
  	   			if awayteam = Replace(table(n1,0),"*","") then  	'check team name, ignoring * (deducted points indicator)	    			exit for
	   				exit for
	   			end if
	   			if table(n1,0) = "" then 
	   				table(n1,0) = awayteam  					'set team name (should no longer be necessary because of the TEAM linss)
	    			exit for
	 			end if
    		next
     		if awaygoals > homegoals then 
	   			table(n1,8) = table(n1,8) + 1  					'increment away wins
				table(n1,13) = table(n1,13) + winpoints			'increment points
				table(n1,20) = right(table(n1,20) & "3",10)		'increment form list - a string of up to 10 chars, each position containing: 7 (H win), 5 (H draw), 4 (H loss), 3 (A win), 1 (A draw), 0 (A loss)
	   			  elseif awaygoals = homegoals then 
	   			    table(n1,9) = table(n1,9) + 1  				'increment away draws
	   			    table(n1,13) = table(n1,13) + 1 			'increment points
	   			    table(n1,20) = right(table(n1,20) & "1",10) 'increment form list - a string of up to 10 chars, each position containing: 7 (H win), 5 (H draw), 4 (H loss), 3 (A win), 1 (A draw), 0 (A loss)
	   			  else 
	   			  table(n1,10) = table(n1,10) + 1				'increment away defeats
	   			  table(n1,20) = right(table(n1,20) & "0",10) 	'increment form list - a string of up to 10 chars, each position containing: 7 (H win), 5 (H draw), 4 (H loss), 3 (A win), 1 (A draw), 0 (A loss)
	   			end if
	   		table(n1,7) = table(n1,7) + 1 	 					'increment away played
	   		table(n1,11) = table(n1,11) + awaygoals    			'increment away goals for
	   		table(n1,12) = table(n1,12) + homegoals    			'increment away goals against

		rs.Movenext
	loop

end if

rs.close


' Apply any points adjustments to the array 

adjustment_reason = ""

sql = "select date, name_then, points_adjust, reason  "
sql = sql & "from FL_points_adjustment "	
sql = sql & "where season = '" & season_years & "' "

rs.open sql,conn,1,2

reason_marker = " "

Do While Not rs.EOF

	if match_date >= rs.Fields("date") then
	
		for n1 = 0 to UBound(table,1)
			temp = Replace(table(n1,0),"*","") 	'remove any any previous * (adjusted points indicators)
  			if rs.Fields("name_then") = temp then  	'check team name
  				reason_marker = reason_marker & "*" 							'create a new marker (e.g. * or **)
    			table(n1,0) = table(n1,0) & reason_marker						'add the adjusted points marker to the team name
    			table(n1,13) = table(n1,13) + rs.Fields("points_adjust")	'apply the points adjustment 
    			adjustind = "y"
				adjustment_reason = adjustment_reason & "<p class=""style1"" style=""margin:3px 0"">" & reason_marker & " " & rs.Fields("reason") & "</p>"	 'There might be more than one (very unlikely)
    			exit for
  			end if
		next
		
	end if	
	
	rs.Movenext
loop

rs.close
conn.close


' Now prepare more array values for the final form values, the combined values and the sorting needs 
For n1 = 0 to UBound(table,1)
  If table(n1,0) > "" then
	'calculate the final form values  
  	for n3 = 1 to len(table(n1,20))		'contains the (up to) 10 char 'form' string 
 		n2 = mid(table(n1,20),n3,1)		'process each byte
		if n2 = 7 or n2 = 3 then table(n1,18) = table(n1,18) + winpoints	'increment the last 10 accumulation (7 indicates a home win, 3 and away win)
		if n2 = 5 or n2 = 1 then table(n1,18) = table(n1,18) + 1			'increment the last 10 accumulation (5 indicates a home draw, 1 and away draw)
		'if the n3 loop has reach the last 5
  		if len(table(n1,20)) - n3 < 5 then 
			if n2 = 7 or n2 = 3 then table(n1,19) = table(n1,19) + winpoints	'increment the last 5 accumulation (7 indicates a home win, 3 and away win)
			if n2 = 5 or n2 = 1 then table(n1,19) = table(n1,19) + 1			'increment the last 5 accumulation (5 indicates a home draw, 1 and away draw)
		end if
  	next 

	table(n1,21) = table(n1,2) + table(n1,8)	'combined wins
	table(n1,22) = table(n1,3) + table(n1,9)	'combined draws
	table(n1,23) = table(n1,4) + table(n1,10)	'combined defeats
	table(n1,24) = table(n1,5) + table(n1,11)	'combined goals for
	table(n1,25) = table(n1,6) + table(n1,12)	'combined goals against
	if table(n1,25) > 0 then table(n1,26) = round(table(n1,24) / table(n1,25),4)	'combined goal average
	table(n1,27) = table(n1,24) - table(n1,25)	'combined goal difference
	if table(n1,14) > 0 then table(n1,28) = int(table(n1,14)/table(n1,29) + 0.5) 'average home attendance (+0.5 to allow roundup)	
	
	inverseteam = ""
 	For n2 = 1 To Len(table(n1,0))
		inversework = Mid(table(n1,0), n2, 1 )
		inverseteam = inverseteam & Chr(127-Asc(inversework))
	Next

	'prepare sort fields in the array range 30+ (e.g. array value 16,7 becomes 16,37)
	For n2 = 0 to 29		
		if IsNumeric(table(n1,n2)) and len(table(n1,n2)) < 7 then
			table(n1,n2+30) = string(6-len(table(n1,n2)),"0") & table(n1,n2) & inverseteam
		 else
			table(n1,n2+30) = table(n1,n2)   
		end if
	Next
	
	'Redo sort values for Goal Difference
	table(n1,57) = string(6-len(table(n1,27)+100),"0") & table(n1,27)+100 & inverseteam 	'Add 100 to the GD value to get around sorting negative numbers
	
	'Redo sort values for Points to concatenate with goal average, goal difference or goals scored, as appropriate for the season
	Select case true
		case match_date > "" and match_date < "1976-07-01"
			'Redo sort values for standard points order, adding on goal average
			table(n1,43) = string(6-len(table(n1,13)),"0") & table(n1,13) & string(6-len(table(n1,26)),"0") & table(n1,26) & inverseteam 'follow points by GA, then GF, then club name
		case match_date > "" and match_date < "1992-07-01"
			'Redo sort values for standard points order, adding on goal difference and goals scored
    		table(n1,43) = string(6-len(table(n1,13)),"0") & table(n1,13) & string(6-len(table(n1,27)+100),"0") & table(n1,27)+100 & string(6-len(table(n1,24)),"0") & table(n1,24) & inverseteam 'follow points by GD, then GF, then club name
		case match_date > "" and match_date < "1999-07-01"
			'Redo sort values for standard points order, adding on goals scored (used between 1992 and 1999)
    		table(n1,43) = string(6-len(table(n1,13)),"0") & table(n1,13) & string(6-len(table(n1,24)),"0") & table(n1,24) & string(6-len(table(n1,24)),"0") & table(n1,24)  & string(6-len(table(n1,25)),"0") & table(n1,25)& inverseteam 'follow points by GD, then GF, then GA, then club name
		case else
			'Redo sort values for modern points order ...
			'... this supplements the standard points with goal difference and goals for, as described in EFL regs: https://www.efl.com/-more/governance/efl-rules--regulations/efl-regulations/section-3-the-league/ (section 9.1) ...
			'... note that section 9.2 and the consequential sections 9.3-9.9 are ignored as the requirements of 9.2 are not always possible during the season; the club-name order is used instead.
    		table(n1,43) = string(6-len(table(n1,13)),"0") & table(n1,13) & string(6-len(table(n1,27)+100),"0") & table(n1,27)+100 & string(6-len(table(n1,24)),"0") & table(n1,24) & inverseteam 'follow points by GD, then GF, then club name
  	end select
  	
  end if
	
next 


' Build the table display

lowtable = UBound(table,1) 
hightable = 0
step = -1
tablemessage = "<font color=""red"">Warning, not normal order</font>"
select case sort
	case 1
		Call QuickSort(table,0,UBound(table,1),51)
	case 2
		Call QuickSort(table,0,UBound(table,1),52)
	case 3
		Call QuickSort(table,0,UBound(table,1),53)
	case 4
		Call QuickSort(table,0,UBound(table,1),54)
	case 5
		Call QuickSort(table,0,UBound(table,1),55)
	case 6
		Call QuickSort(table,0,UBound(table,1),56)	'GA
	case 7
		Call QuickSort(table,0,UBound(table,1),57)	'GD
	case 8
		Call QuickSort(table,0,UBound(table,1),54)	'GF
	case 9
		Call QuickSort(table,0,UBound(table,1),43)
		tablemessage = "Standard table view"
	case 10
		Call QuickSort(table,0,UBound(table,1),32)
	case 11
		Call QuickSort(table,0,UBound(table,1),33)
	case 12
		Call QuickSort(table,0,UBound(table,1),34)
	case 13
		Call QuickSort(table,0,UBound(table,1),35)
	case 14
		Call QuickSort(table,0,UBound(table,1),36)
	case 15
		Call QuickSort(table,0,UBound(table,1),38)
	case 16
		Call QuickSort(table,0,UBound(table,1),39)
	case 17
		Call QuickSort(table,0,UBound(table,1),40)
	case 18
		Call QuickSort(table,0,UBound(table,1),41)
	case 19
		Call QuickSort(table,0,UBound(table,1),42)
	case 20
		Call QuickSort(table,0,UBound(table,1),45)
	case 21
		Call QuickSort(table,0,UBound(table,1),46)
	case 22
		Call QuickSort(table,0,UBound(table,1),58)
	case 23
		Call QuickSort(table,0,UBound(table,1),48)
	case 24
		Call QuickSort(table,0,UBound(table,1),49)
end select

' Check if season is underway
if match_date > "" then
	headdate1 = "After all matches on" 
	headdate2 = left(WeekDayName(WeekDay(match_date)),3) & ", " & FormatDateTime(match_date,1)
  else
  	headdate1 = " "
  	headdate2 = " "
end if


output = "<div id=""table"">"

' Write out the top part of display (unless the request has come from the match page)

if frommatch = "" then

	output = output & "<h3 style=""font-size: 14px; margin-bottom: 6px"">" & division & " " & season_years & "</h3>"	
	output = output & "<p class=""style1"" style=""margin: 0 0 3px"">Select Another Date</p>"
	
	output = output & "<form style=""margin: 0 0 6px;"" action=""progresstables.asp"" method=""post"" name=""form1"">"
	output = output & "<select name=""D"">"
	output = output & "<option value="""">Day</option>"
	for n = 1 to 31
		output = output & "<option value=""" & n & """"
		if n = CInt(request_day) then output = output & " selected" 
		output = output & ">" & n & "</option>"
	next
	output = output & "</select>"
	
	output = output & "<select name=""M"">"
	n = 1
	const monthlist="Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"
	months = split(monthlist,",")
	output = output & "<option value="""">Mon</option>"
	for each mon in months
		output = output & "<option value=""" & n & """"
		if n = CInt(request_mon) then output = output & " selected"
		output = output & ">" & mon & "</option>"
		n = n + 1
	next
	output = output & "</select>"
	
	output = output & "<select name=""Y"">"
	output = output & "<option value="""">Year</option>"
	for n = 1920 to year(Date)
		output = output & "<option value=""" & n & """"
		if n = CInt(request_year) then output = output & " selected"
		output = output & ">" & n & "</option>"
	next
	output = output & "</select>"
	output = output & "<input type=""submit"" value=""Go"" name=""B1"">"
	output = output & "</form>"
	 
	output = output & thisdate_message	
	output = output & topnote
	
end if

output = output & "<table bordercolor=""#808080"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""border-collapse: collapse"" >"
output = output & "<tr class=""colhead1"">"
output = output & "<td class=""head2"" colspan=""3"" style=""border-bottom: none;"">Latest Form <img style=""height:12px"" title=""Points in the last 10 and 5 games"" src=""images/help.gif""</td>"
output = output & "<td colspan=""2"" style=""border-top: none; border-bottom: none; vertical-align: top; text-align: center;"">" & headdate1 & "</td>"
output = output & "<td class=""head3"" colspan=""8"" style=""border: 2px solid #808080; border-bottom: none;"">" & tablemessage & "</td>"
output = output & "<td class=""head2"" colspan=""5"" style=""border-bottom: none;"">Home</td>"
output = output & "<td class=""head2"" colspan=""5"" style=""border-bottom: none;"">Away</td>"
if totalattendance > 0 then output = output & "<td class=""head2"" colspan=""4"" style=""border-bottom: none;"">Home Attendance></td>"
output = output & "</tr>"

output = output & "<tr class=""colhead2"">"
output = output & "<td class=""head2"" style=""border-top: none; font-size: 10px;"">"
output = output & "<img src=""images/formhome.gif"" hspace=""1"" border=""1"" height=""10"" width=""3"">home<img src=""images/formaway.gif"" hspace=""1"" border=""1"" height=""10"" width=""3"">away"
output = output & "<br><img src=""images/formhome.gif"" hspace=""1"" border=""1"" height=""10"" width=""3"">W <img src=""images/formhome.gif"" hspace=""1"" border=""1"" height=""5"" width=""3"">D <img src=""images/formhome.gif"" hspace=""1"" border=""1"" height=""2"" width=""3"">L</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=23&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>-10</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=24&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>-5</td>"
output = output & "<td colspan=""2"" style=""border-top: none; border-bottom: 2px solid #808080; border-right: 2px solid #808080; text-align: center; font-weight: bold;"">" & headdate2
if frommatch = "" then
	output = output & "<br><ul>"  
	if IsDate(prev_date) then output = output & "<li><a href=""progresstables.asp?date=" & prev_date & """>Previous</a></li>"
	if IsDate(next_date) then output = output & "<li><a href=""progresstables.asp?date=" & next_date & """>Next</a></li></ul>"
end if
output = output & "</td>"
output = output & "<td class=""head3"" style=""border-top: none;"">P</td>"
output = output & "<td class=""head3"" style=""border-top: none;""><a href=""progresstables.asp?sort=1&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>W</td>"
output = output & "<td class=""head3"" style=""border-top: none;""><a href=""progresstables.asp?sort=2&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>D</td>"
output = output & "<td class=""head3"" style=""border-top: none;""><a href=""progresstables.asp?sort=3&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>L</td>"
output = output & "<td class=""head3"" style=""border-top: none;""><a href=""progresstables.asp?sort=4&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>F</td>"
output = output & "<td class=""head3"" style=""border-top: none;""><a href=""progresstables.asp?sort=5&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>A</td>"
select case ordertype
	case 1
		output = output & "<td class=""head3"" style=""border-top: none;""><a href=""progresstables.asp?sort=6&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>GA</td>"
	case 2
		output = output & "<td class=""head3"" style=""border-top: none;""><a href=""progresstables.asp?sort=7&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>GD</td>"
	case 3
		output = output & "<td class=""head3"" style=""border-top: none;""><a href=""progresstables.asp?sort=8&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>GF</td>"
end select
output = output & "<td class=""head3"" style=""border-top: none; border-right: 2px solid #808080;""><a href=""progresstables.asp?sort=9&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>Pts </td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=10&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>W</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=11&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>D</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=12&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>L</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=13&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>F</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=14&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>A</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=15&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>W</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=16&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>D</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=17&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>L</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=18&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>F</td>"
output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=19&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>A</td>"
if totalattendance > 0 then
	output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=20&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>Highest</td>"
	output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=21&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>Lowest</td>"
	output = output & "<td class=""head2"" style=""border-top: none;""><a href=""progresstables.asp?sort=22&date=" & match_date & """><img src=""images/sort.gif"" border=""0"" vspace=""2""></a><br>Average</td>"
	output = output & "<td class=""head2"" style=""border-top: none;"">+ or -<br>" & divisionshort & " Avg</td>"
end if
output = output & "</tr>"

n2 = 0
For n1 = lowtable to hightable Step step
  If table(n1,0) > "" then
  	n2 = n2 + 1
  	trclass = ""
  	if table(n1,0) = "Plymouth Argyle" then trclass = trclass & "argyle "
	if n2 = pos_promote then trclass = trclass & "promotion "
	if n2 = pos_promote_playoff then trclass = trclass & "playoffs "
	if n2 = pos_relegate_playoff - 1 then trclass = trclass & "playoffs_rel "
	if n2 = pos_relegate - 1 then trclass = trclass & "relegation "
   	output = output & "<tr class=""" & trclass & """ onmouseover=""this.className = '" & trclass & " rowhlt';"" onmouseout=""this.className = '" & trclass & "';"">" 
   	output = output & "<td style=""vertical-align:bottom"">"
   	for n3 = 1 to len(table(n1,20))
   		formgif = "formhome"
   		select case mid(table(n1,20),n3,1)
   			case 7 
   				height = 10
   			case 5
   				height = 5
   			case 4
   				height = 2
   			case 3 
   				height = 10
   				formgif = "formaway"
   			case 1
   				height = 5
   				formgif = "formaway"
   			case 0
   				height = 2
   				formgif = "formaway"	
   		end select		
   		output = output & "<img src=""images/" & formgif & ".gif"" hspace=""1"" border=""1"" height=""" & height & """ width=""3"">"
   	next 
   	output = output & "</td>"
	output = output & "</td><td class=""num1"">" & table(n1,18) & "</td><td class=""num1"">" & table(n1,19) & "</td>"
	
	output = output & "<td class=""num2 main"">" & n2 & "</td><td class=""main"">" & Replace(table(n1,0),"Wolverhampton","Wolve'ton")
	output = output & "</td><td class=""maindark num1"">" & table(n1,1) + table(n1,7) 
	output = output & "</td><td class=""main num1"">" & table(n1,21) & "</td><td class=""num1 main"">" & table(n1,22)
	output = output & "</td><td class=""main num1"">" & table(n1,23) & "</td><td class=""num1 main"">" & table(n1,24)
	output = output & "</td><td class=""main num1"">" & table(n1,25)
	output = output & "</td><td class=""main"
	select case ordertype
	case 1
		output = output & """>" & table(n1,26)
	case 2
		output = output & " num1"">" & table(n1,27)
	case 3
		output = output & " num1"">" & table(n1,24)
	end select
	output = output & "</td><td class=""num1 maindark"">" & table(n1,13) & "</td>"

	output = output & "<td class=""num2"">" & table(n1,2) & "</td><td class=""num1"">" & table(n1,3)
	output = output & "</td><td class=""num1"">" & table(n1,4) & "</td><td class=""num1"">" & table(n1,5)
	output = output & "</td><td class=""num1"">" & table(n1,6) & "</td><td class=""num1"">" & table(n1,8)
	output = output & "</td><td class=""num1"">" & table(n1,9) & "</td><td class=""num1"">" & table(n1,10)
	output = output & "</td><td class=""num1"">" & table(n1,11) & "</td><td class=""num1"">" & table(n1,12)
	if totalattendance > 0 then
		output = output & "</td><td class=""num1"">" & FormatNumber(table(n1,15),0)
		if table(n1,16) = 999999 then table(n1,16) = 0   'avoid initial value (happens before the start of the season)  
		output = output & "</td><td class=""num1"">" & FormatNumber(table(n1,16),0)
		output = output & "</td><td class=""num1"">" & FormatNumber(table(n1,28),0)
		if totalmatches = 0 then totalmatches = 1   'avoid a divide by zero (happens before the start of the season)
		if FormatNumber(table(n1,28)) > 0 then   
			output = output & "</td><td class=""num1"">" & FormatNumber(table(n1,28) - totalattendance/totalmatches,0)
		   else
		   	output = output & "</td><td class=""num1"">"
		end if
		output = output & "</td>"
	end if
	output = output & "</tr>"
  end if	
next 			

output = output & "<tr>"
output = output & "<td colspan=""3"" style=""border: none;""></td>"
output = output & "<td colspan=""10"" style=""border: none; border-top: 2px solid #808080;"">"
if adjustind = "y" then output = output & adjustment_reason & "</td>"
output = output & "<td colspan=""14"" style=""border: none"";></td>"
output = output & "</tr>"
output = output & "</table>"

response.write(output)

%>

</div>

<!--#include file="base_code.htm"-->

</body>

</html>

<%
Sub SwapRows(ary,row1,row2)
  '== This proc swaps two rows of an array 
  Dim x,tempvar
  For x = 0 to Ubound(ary,2)
    tempvar = ary(row1,x)    
    ary(row1,x) = ary(row2,x)
    ary(row2,x) = tempvar
  Next
End Sub  'SwapRows

Sub QuickSort(vec,loBound,hiBound,SortField)

  '==--------------------------------------------------------==
  '== Sort a 2 dimensional array on SortField                ==
  '==                                                        ==
  '== This procedure is adapted from the algorithm given in: ==
  '==    ~ Data Abstractions & Structures using C++ by ~     ==
  '==    ~ Mark num1 headington and David Riley, pg. 586    ~     ==
  '== Quicksort is the fastest array sorting routine for     ==
  '== unordered arrays.  Its big O is  n log n               ==
  '==                                                        ==
  '== Parameters:                                            ==
  '== vec       - array to be sorted                         ==
  '== SortField - The field to sort on (2nd dimension value) ==
  '== loBound and hiBound are simply the upper and lower     ==
  '==   bounds of the array's 1st dimension.  It's probably  ==
  '==   easiest to use the LBound and UBound functions to    ==
  '==   set these.                                           ==
  '==--------------------------------------------------------==

  Dim pivot(),loSwap,hiSwap,counter
  Redim pivot (Ubound(vec,2))

  '== Two items to sort
  if hiBound - loBound = 1 then
    if vec(loBound,SortField) > vec(hiBound,SortField) then Call SwapRows(vec,hiBound,loBound)
  End If

  '== Three or more items to sort
  
  For counter = 0 to Ubound(vec,2)
    pivot(counter) = vec(int((loBound + hiBound) / 2),counter)
    vec(int((loBound + hiBound) / 2),counter) = vec(loBound,counter)
    vec(loBound,counter) = pivot(counter)
  Next

  loSwap = loBound + 1
  hiSwap = hiBound
  
  do
    '== Find the right loSwap
    while loSwap < hiSwap and vec(loSwap,SortField) <= pivot(SortField)
      loSwap = loSwap + 1
    wend
    '== Find the right hiSwap
    while vec(hiSwap,SortField) > pivot(SortField)
      hiSwap = hiSwap - 1
    wend
    '== Swap values if loSwap is less then hiSwap
    if loSwap < hiSwap then Call SwapRows(vec,loSwap,hiSwap)


  loop while loSwap < hiSwap
  
  For counter = 0 to Ubound(vec,2)
    vec(loBound,counter) = vec(hiSwap,counter)
    vec(hiSwap,counter) = pivot(counter)
  Next
    
  '== Recursively call function .. the beauty of Quicksort
    '== 2 or more items in first section
    if loBound < (hiSwap - 1) then Call QuickSort(vec,loBound,hiSwap-1,SortField)
    '== 2 or more items in second section
    if hiSwap + 1 < hibound then Call QuickSort(vec,hiSwap+1,hiBound,SortField)

End Sub  'QuickSort
%>