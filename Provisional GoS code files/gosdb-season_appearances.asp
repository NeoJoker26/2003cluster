
<%
Dim ColOffset, matchList,objXML, objXMLmatch, objXML1, matXML, id1, output, output1, output2, outputhold, players, match, player, numnamePlayer, namePlayer, substitute, substitute2, substitute3, squad(99,17), arraybit, num, n0, n1, n2, result(80,15), n3, n4, n5, n6, n7, pos1, pos2, square1, square2
Dim subhold(5), subind, scorers, scorer, scorename, goaltime, numgoals, graphic, title1, title2, cardhold, matchno, homeno, season(99,80,2), contentlink(80,2), position(24,50), lpos, thhlt, ptsbar, sort, lowsquad, highsquad, step, hometotattend, homecount, maint
Dim season_start, season_end, years, season_no, pointsforawin, currentseason, division, divisionpart, divisionparts, D2Tflag

Dim conn, sql, rs, rs1, sqlhold
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

sort = Request.QueryString("sort")
years = Request.QueryString("years")

output = "" 

sql = "select max(years) as years "
sql = sql & "from season " 


rs.open sql,conn,1,2

if years = rs.Fields("years") then currentseason = "Y"

if years = ""  then 
	years = rs.Fields("years")
	currentseason = "Y"
end if

rs.close

sql = "select season_no, years, division, date_start, date_end "
sql = sql & "from season " 
sql = sql & "where years = '" & years & "' "

rs.open sql,conn,1,2

season_no = rs.Fields("season_no")

division = ""
divisionparts = split(lcase(rs.Fields("division"))," ")
for each divisionpart in divisionparts
	division = division & ucase(left(divisionpart,1)) & right(divisionpart,len(divisionpart)-1) & " "
next


season_start = rs.Fields("date_start")
season_end = rs.Fields("date_end")
rs.close

output = output & "<p style=""font-size: 14px; font-weight:700; color:green; margin: 0 0 6px; text-align: center;"">PLAYER APPEARANCES</p>"
output = output & "<p class=""style1"" style=""margin: 0; text-align: center;"">Click on "																	
output = output & "<img border=""0"" src=""images/more.png"">  along the top for the match page, and down the side for player details."  																	'Delete this line when using for the season summary page 
output = output & "</p>"
output = output & "<p style=""margin: 9px 0; text-align: center;""><span style=""font-weight:bold; font-size:12px"">#</span> "
if left(years,4) < "2011" then
	output = output & "Usual position&nbsp;&nbsp;"
  else
	output = output & "Squad Number&nbsp;&nbsp;"  
end if
output = output & "<img style=""border:0; vertical-align:text-top;"" src=""images/spot.gif""> starting appearance&nbsp;&nbsp;"
if left(years,4) >= "1965" then output = output & "<img style=""border:0; vertical-align:text-top;"" src=""images/spotsub.gif""> appearing substitute&nbsp;&nbsp;"
output = output & "<img style=""border:0; vertical-align:text-top;"" src=""images/spot_1.gif""> goals&nbsp;&nbsp;"
if left(years,4) >= "2006" then 
	output = output & "<img style=""border:0; vertical-align:text-top;"" src=""images/spot_y.gif""> cautioned&nbsp;&nbsp;"
	output = output & "<img style=""border:0; vertical-align:text-top;"" src=""images/spot_r.gif""> sent off"
end if
output = output & "</p>"					

output = output & "<table class=""apptable"">"
output = output & "<tr><th id=""sort0"" class=""sort""><span style=""font-family:verdana; font-weight:bold; font-size:12px"">#</span><br>"
output = output & "<img src=""images/sort.gif"" border=""0"" hspace=""1"" vspace=""2""></th>"
output = output & "<th id=""sort1"" class=""sort""><span style=""font-weight:bold; font-size:12px"">&nbsp</span><br>"
output = output & "<img src=""images/sort.gif"" border=""0"" hspace=""1"" vspace=""2""></th>"


' Initialise squad array 

For n1 = 0 to UBound(squad,1)
	For n2 = 2 to UBound(squad,2)
		squad(n1,n2) = 0
	next
next
	
sql = "select distinct a.player_id, player_id_spell1, c.surname, forename, left(initials,1) as initial, isnull(squad_no,0) as squad_no "
sql = sql & "from match_player a join season b on date_start <= a.date and date_end >= a.date "
sql = sql & " join player c on a.player_id = c.player_id "
sql = sql & " left outer join player_squad d on c.player_id = d.player_id and b.season_no = d.season_no " 
sql = sql & "where b.season_no = " & season_no
'The following union picks up players who were given squad numbers but did not play that season
sql = sql & "union "
sql = sql & "select distinct a.player_id, player_id_spell1, b.surname, forename, left(initials,1) as initial, squad_no "
sql = sql & "from player_squad a "
sql = sql & "join player b on a.player_id = b.player_id "
sql = sql & "where season_no = " & season_no
sql = sql & "order by squad_no, surname, forename "

rs.open sql,conn,1,2

n1 = 0

Do While Not rs.EOF
	squad(n1,0) = rs.Fields("squad_no")
	squad(n1,1) = trim(rs.Fields("initial") & " " & rs.Fields("surname")) 
	squad(n1,10) = n1											'unsorted player number
	squad(n1,11) = rs.Fields("surname")							'surname only (for sorting)										
	squad(n1,15) = rs.Fields("player_id")						'GoS-DB playerid
	squad(n1,16) = right("0" & squad(n1,0),2) & squad(n1,11)	'squad no + surname (for sorting on squad no)
	squad(n1,17) = rs.Fields("player_id_spell1")				'GoS-DB first spell playerid for player page										
	n1 = n1 + 1				
	rs.MoveNext
Loop
rs.close

if currentseason = "Y" then
	sqlhold = "union all " 
	sqlhold = sqlhold & "select a.date, a.opposition, opposition_qual, name_abbrev, lfc, homeaway, 'Home Park', NULL, NULL, cupinitial, subcomp, competition, NULL, "
	sqlhold = sqlhold & " case when a.compcode = 'D2T' then 2 when a.compcode = 'SWRL' then 3 when lfc = 'C' then 4 else 1 end "
	sqlhold = sqlhold & "from season_this a join competition b on a.compcode = b.compcode " 
	sqlhold = sqlhold & "join opposition c on a.opposition = c.name_then "
	sqlhold = sqlhold & "where not exists (select * from match c where c.date = a.date) "
	sqlhold = sqlhold & "  and homeaway = 'H' "
	sqlhold = sqlhold & "  and date between '" & season_start & "' and '" & season_end & "'"
	sqlhold = sqlhold & "union all " 
	sqlhold = sqlhold & "select a.date, a.opposition, opposition_qual, name_abbrev, lfc, homeaway, ground_name, NULL, NULL, cupinitial, subcomp, competition, NULL, "
	sqlhold = sqlhold & " case when a.compcode = 'D2T' then 2 when a.compcode = 'SWRL' then 3 when lfc = 'C' then 4 else 1 end "
	sqlhold = sqlhold & "from season_this a join competition b on a.compcode = b.compcode " 
	sqlhold = sqlhold & "join opposition c on a.opposition = c.name_then "
	sqlhold = sqlhold & "join venue d on a.opposition = d.club_name_then and a.date between d.first_game and d.last_game "   
	sqlhold = sqlhold & "where not exists (select * from match c where c.date = a.date) "
	sqlhold = sqlhold & "  and homeaway in ('A', 'N') "
	sqlhold = sqlhold & "  and date between '" & season_start & "' and '" & season_end & "'"
  else
	sqlhold = ""
end if

sql = "select date, opposition, opposition_qual, NULL as name_abbrev, lfc, homeaway, 'Home Park' as ground_name, goalsfor, goalsagainst, cupinitial, subcomp, competition, attendance, "
sql = sql & " case when a.compcode = 'D2T' then 2 when a.compcode = 'SWRL' then 3 when lfc = 'C' then 4 else 1 end as comporder "
sql = sql & "from v_match_season a " 
sql = sql & "where season_no = " & season_no
sql = sql & "  and homeaway = 'H' "
sql = sql & "union all "
sql = sql & "select date, opposition, opposition_qual, NULL, lfc, homeaway, ground_name, goalsfor, goalsagainst, cupinitial, subcomp, competition, attendance, "
sql = sql & " case when a.compcode = 'D2T' then 2 when a.compcode = 'SWRL' then 3 when lfc = 'C' then 4 else 1 end as comporder "
sql = sql & "from v_match_season a "
sql = sql & "join venue b on a.opposition = b.club_name_then and a.date between b.first_game and b.last_game "  
sql = sql & "where season_no = " & season_no
sql = sql & "  and homeaway in ('A', 'N') "
sql = sql & sqlhold
sql = sql & "order by comporder, date"
rs.open sql,conn,1,2

matchno = 0
homeno = 0

Do While Not rs.EOF

		erase subhold
		subind = 0
		if rs.Fields("lfc") <> "C" and rs.Fields("homeaway") = "H" then homeno = homeno + 1
		matchno = matchno + 1

			if rs.Fields("lfc") = "C" and matchno < 61 then matchno = 61	'move on to cup games 
			if rs.Fields("lfc") = "C" then
				ColOffset = 3
			  else
				ColOffset = 0
			end if  
			 						
		 	result(matchno,0) = matchno
		 	result(matchno,1) = rs.Fields("homeaway")
			result(matchno,9) = rs.Fields("attendance")	
		 	result(matchno,10) = rs.Fields("competition")
		 	result(matchno,10) = Replace(result(matchno,10),"'","\\\'")	'deal with an apostrophe in the competition name 
		 	result(matchno,11) = rs.Fields("date") 
		 	result(matchno,13) = rs.Fields("ground_name") 
		 	result(matchno,13) = Replace(result(matchno,13),"'","")
			result(matchno,14) = rs.Fields("name_abbrev")
			
			sql = "with cte as (select row_number() over(order by date) as match, date, attendance "
			sql = sql & "from v_match_season "    
			sql = sql & "where season_no = " & season_no - 1
			sql = sql & "and lfc <> 'C' "
			sql = sql & "and homeaway = 'H' "
			sql = sql & ") "
			sql = sql & "select sum(attendance)/count(*) as aveatt "
			sql = sql & "from cte "
			sql = sql & "where match <= " & homeno
			rs1.open sql,conn,1,2
			
			result(matchno,15) = rs1.Fields("aveatt")   'average home attendance last season
			
			rs1.close

			if matchno > 60 then result(matchno,2) = rs.Fields("cupinitial") & left(rs.Fields("subcomp"),1)

		 	if rs.Fields("homeaway") = "A" then
		 		result(matchno,12) = rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " v Argyle"
				if not isnull(rs.Fields("goalsfor")) then result(matchno,8) = rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " " & rs.Fields("goalsagainst") & " Argyle " & rs.Fields("goalsfor")
		 	  else
		 		result(matchno,12) = "Argyle v " & rs.Fields("opposition") & " " & rs.Fields("opposition_qual")
		 		if not isnull(rs.Fields("goalsfor")) then result(matchno,8) = "Argyle " & rs.Fields("goalsfor") & " " & rs.Fields("opposition") & " " & rs.Fields("opposition_qual") & " " &rs.Fields("goalsagainst") 
		 	end if
 	 	 		
 		 	if rs.Fields("goalsfor") > rs.Fields("goalsagainst") then
		 		result(matchno,3) = 3 
		 	  elseif rs.Fields("goalsfor") = rs.Fields("goalsagainst") then
				result(matchno,3) = 1
			  else 
			  	result(matchno,3) = 0
		 	end if
						
			' Initialise content link
			contentlink(matchno,0) = ""
			contentlink(matchno,1) = "more_grey.png"
	
			if result(matchno,8) > "" then 
			  contentlink(matchno,0) = "<div class=""nohover""><a href=""gosdb-match.asp?date=" & result(matchno,11) & """>"
			  contentlink(matchno,1) = "more.png"
			  elseif matchno < 60 then 
			  	contentlink(matchno,0) = "<div class=""nohover""><a href=""#" & result(matchno,14) & """>"
			end if
		
		 		 			
' Now count all players involved

			sql = "select player_id, startpos, replaced_by, card "
			sql = sql & "from match_player a "
			sql = sql & "where date = '" & rs.Fields("date") & "' "
			rs1.open sql,conn,1,2

			Do While Not rs1.EOF
					
				For n1 = 0 to UBound(squad,1)
				
					if squad(n1,15) = rs1.Fields("player_id") then
						
						if rtrim(rs1.Fields("card")) = "y" then 
							squad(n1,8) = squad(n1,8) + 1
							season(n1,matchno,1) = "y"
						end if
						
						if rtrim(rs1.Fields("card")) = "r" or rs1.Fields("card") = "yr" then 
							squad(n1,9) = squad(n1,9) + 1
							season(n1,matchno,1) = "r"
						end if
				
						if rs1.Fields("startpos") > 0 then
					
							squad(n1,2+ColOffset) = squad(n1,2+ColOffset) + 1
							squad(n1,12) = squad(n1,12) + 1
							
							if not isnull(rs1.Fields("player_id")) then
								season(n1,matchno,0) = 2
						  		else season(n1,matchno,0) = 1
							end if
					
						  else
					
							squad(n1,3+ColOffset) = squad(n1,3+ColOffset) + 1
							squad(n1,13) = squad(n1,13) + 1
							season(n1,matchno,0) = 3			
						
						end if
						
						exit For
					
					end if
					
				next				
				
				rs1.MoveNext
			Loop
			rs1.close
		
' Now count all goalscorers

			sql = "select player_id, count(*) as numgoals "
			sql = sql & "from match_goal a "
			sql = sql & "where date = '" & rs.Fields("date") & "' "
			sql = sql & "group by player_id "
			rs1.open sql,conn,1,2

			Do While Not rs1.EOF
										
				For n1 = 0 to UBound(squad,1)
				
					if squad(n1,15) = rs1.Fields("player_id") then	   

						squad(n1,4+ColOffset) = squad(n1,4+ColOffset) + rs1.Fields("numgoals")
						squad(n1,14) = squad(n1,14) + rs1.Fields("numgoals")
						season(n1,matchno,2) = rs1.Fields("numgoals")
						exit For
					end if

				next
			
				rs1.MoveNext
			Loop
			rs1.close
	 			
		rs.MoveNext
	Loop
	rs.close

' Build squad display

lowsquad = UBound(squad,1) 
highsquad = 0
step = -1
 
select case sort
	case 0
		Call QuickSort(squad,0,UBound(squad,1),16)
		lowsquad = 0 
		highsquad = UBound(squad,1)
		step = 1
	case 1
		Call QuickSort(squad,0,UBound(squad,1),11)
		lowsquad = 0 
		highsquad = UBound(squad,1)
		step = 1
	case 2
		Call QuickSort(squad,0,UBound(squad,1),12)
	case 3
		Call QuickSort(squad,0,UBound(squad,1),13)
	case 4
		Call QuickSort(squad,0,UBound(squad,1),14)
	case 5
		Call QuickSort(squad,0,UBound(squad,1),9)
	case 6
		Call QuickSort(squad,0,UBound(squad,1),8)
end select
 
output = output & "<th id=""sort2"" class=""sort"" onmouseover=""showtip('Starting appearances<br>(League and Cup)')"" onmouseout=""hidetip()""><img src=""images/spot.gif""><br><img src=""images/sort.gif"" border=""0"" hspace=""1"" vspace=""2""></th>"
if left(years,4) >= "1965" then output = output & "<th id=""sort3"" class=""sort"" onmouseover=""showtip('Substitute appearances<br>(League and Cup)')"" onmouseout=""hidetip()""><img src=""images/spotsub.gif""><br><img src=""images/sort.gif"" border=""0"" hspace=""1"" vspace=""2""></th>"
output = output & "<th id=""sort4"" class=""sort"" onmouseover=""showtip('Goals scored')"" onmouseout=""hidetip()""><img src=""images/spot_1.gif""><br><img src=""images/sort.gif"" border=""0"" hspace=""1"" vspace=""2""></th>"
if left(years,4) >= "2006" then 
	output = output & "<th id=""sort5"" class=""sort"" onmouseover=""showtip('Red cards')"" onmouseout=""hidetip()""><img src=""images/spot_r.gif""><br><img src=""images/sort.gif"" border=""0"" hspace=""1"" vspace=""2""></th>"
	output = output & "<th id=""sort6"" class=""sort"" onmouseover=""showtip('Yellow Cards')"" onmouseout=""hidetip()""><img src=""images/spot_y.gif""><br><img src=""images/sort.gif"" border=""0"" hspace=""1"" vspace=""2""></th>"
end if
	
homecount = 0
hometotattend = 0	
outputhold = ""
For matchno = 1 to 60
	if result(matchno,1) = "H" then
		ptsbar = "homebar_"
	   else 
	    ptsbar = "awaybar_"
	end if
	if result(matchno,8) > "" and contentlink(matchno,1) = "more.png" then
		output1 = "<th class=""match"" onmouseover=""showtip('" & result(matchno,10) & "<br>" & FormatDateTime(result(matchno,11),1) & "<br><b>" & result(matchno,8) & "</b>"
		output1 = output1 & "<br>" & result(matchno,13)
		if result(matchno,9) > 0 then output1 = output1 & ", att: " & FormatNumber(result(matchno,9),0) 
		if result(matchno,1) = "H" and result(matchno,9) > 0 and result(matchno,15) > 0 then
			output1 = output1 & "<br>At this point, average home league attendance<br>this season: " 
			hometotattend = hometotattend + result(matchno,9)
			homecount = homecount + 1
			output1 = output1 & FormatNumber(Round(hometotattend/homecount),0) & "; last season: " & FormatNumber(result(matchno,15),0) 
		end if	
		output1 = output1 & "')"" onmouseout=""hidetip()"">"
		
		if result(matchno,10) = "South West Regional League" then
			output1 = output1 & matchno-3
 	  	  else
			output1 = output1 & matchno
		end if
		outputhold = outputhold + output1 + "</th>" 
		output2 = "<br>" & contentlink(matchno,0) & "<img border=""0"" src=""images/" & contentlink(matchno,1) & """>"
		output2 = output2 & "<br><img border=""0"" vspace=""2"" src=""images/" & ptsbar & result(matchno,3) & ".gif""></a></th>"
		if result(matchno,10) = "Division Two [season abandoned]" and matchno = 3 then
			output2 = output2 & "<th style=""width:12px; border-top:0; border-bottom:0""></th>" 	'add a separation column between the D2T matches and the SWRL
			outputhold = outputhold & "<th style=""width:8px; border-top:0; border-bottom:0""></th>" 	'add a separation column between the D2T matches and the SWRL
			D2Tflag = "Y"
		end if	
		output = output + output1 + output2
	 else
	 if result(matchno,12) > "" then
		output1 = "<th class=""match"" onmouseover=""showtip('" & result(matchno,10) & "<br>" & FormatDateTime(result(matchno,11),1) & "<br><b>" & result(matchno,12) & "</b>"
		if result(matchno,1) = "H" and result(matchno,9) > 0 and result(matchno,15) > 0 then
			output1 = output1 & "<br>Average league attendance at home<br>last season: " & FormatNumber(result(matchno,15),0) 
		end if
		output1 = output1 & "')"" onmouseout=""hidetip()"">" 
		output1 = output1 & matchno
		outputhold = outputhold + output1 + "</th>"
		output2 = "<br>" & contentlink(matchno,0) & "<img border=""0"" src=""images/" & contentlink(matchno,1) & """>"	
		output2 = output2 & "<img border=""0"" vspace=""2"" src=""images/dummbar_0.gif""></a></div></th>"		
		output = output + output1 + output2
	 end if
	end if
next
  
For matchno = 61 to 80

	if result(matchno,8) > "" and contentlink(matchno,1) = "more.png" then
		output1 = "<th class=""match"" onmouseover=""showtip('" & result(matchno,10) & "<br>" & FormatDateTime(result(matchno,11),1) & "<br><b>" & result(matchno,8) & "</b>" 
		output1 = output1 & "<br>" & result(matchno,13) & ", att: " & result(matchno,9) & "')"" onmouseout=""hidetip()"">" 
		output1 = output1 & right(result(matchno,2),2) 
		outputhold = outputhold + output1 + "</th>"
		output2 = "<br>" & contentlink(matchno,0) & "<img border=""0"" src=""images/" & contentlink(matchno,1) & """></a></th>"
		output = output + output1 + output2
	else
	 if result(matchno,12) > "" then
		output1 = "<th class=""match"" onmouseover=""showtip('" & result(matchno,10) & "<br>" & FormatDateTime(result(matchno,11),1) & "<br><b>" & result(matchno,12) & "</b>" 
		output1 = output1 & "')"" onmouseout=""hidetip()"">" 
		output1 = output1 & right(result(matchno,2),2) 
		outputhold = outputhold + output1 + "</th>"
	 	output2 = "<br>" & contentlink(matchno,0) & "<img border=""0"" src=""images/" & contentlink(matchno,1) & """></a></th>"
	 	output = output + output1 + output2
	 end if
	end if
next

For n0 = lowsquad to highsquad Step step
  If squad(n0,1) > "" Then 
	
	n1 = squad(n0,10) 'unsorted player number for use in season array

	output = output & "<tr>"

	For n2 = 0 to 1
		select case n2
		case 0
			output = output & "<td class=""num"" style=""padding:2 2 1 0"">" 
		case 1
			output = output & "<td class=""name"">"
			if squad(n0,17) <> "xxx" then output = output & "<a href=""gosdb-players2.asp?pid=" & squad(n0,17) & "&from=appear""><img style=""margin-right: 6px; border:none;"" src=""images/more.png"">" 
		end select
		output = output & squad(n0,n2)
		select case n2
		case 0
			output = output & "</td>"
		case 1
			if squad(n0,17) <> "xxx" then output = output & "</a>"			
		end select		
	next
	
	output = output & "<td class=""num"">" & squad(n0,2)+squad(n0,5) & "</td>"
	if left(years,4) >= "1965" then output = output & "<td class=""num"">" & squad(n0,3)+squad(n0,6) & "</td>"
	output = output & "<td class=""num"">" & squad(n0,4)+squad(n0,7) & "</td>"
	if left(years,4) >= "2006" then 	
		output = output & "<td class=""num"">" & squad(n0,9) & "</td>"
		output = output & "<td class=""num"">" & squad(n0,8) & "</td>"
	end if	
	
	For matchno = 1 to 60
	 if result(matchno,0) > "" then
	   if season(n1,matchno,0) > 0 then
		output = output & "<td><img src=""images/spot"
			if season(n1,matchno,0) = 3 then output = output & "sub"
			if season(n1,matchno,2) > 0 then output = output & "_" & season(n1,matchno,2)
			if season(n1,matchno,1) = "y" then output = output & "_y"
			if season(n1,matchno,1) = "r" then output = output & "_r"
		output = output & ".gif""></td>"
		else output = output & "<td></td>"
	   end if
	   if D2Tflag = "Y" and matchno = 3 then
	   	output = output & "<td style=""border-top:0; border-bottom:0""></td>"	'add a separation column between the D2T matches and the SWRL
	   end if	 
	 end if
	next
	
	For matchno = 61 to 80
	 if result(matchno,0) > "" then
	   if season(n1,matchno,0) > 0 then
		output = output & "<td><img src=""images/spot"
			if season(n1,matchno,0) = 3 then output = output & "sub"
			if season(n1,matchno,2) > 0 then output = output & "_" & season(n1,matchno,2)
			if season(n1,matchno,1) = "y" then output = output & "_y"
			if season(n1,matchno,1) = "r" then output = output & "_r"
		output = output & ".gif""></td>"
		else output = output & "<td></td>"
	   end if	 
	 end if
	next
 end if
output = output & "</tr>" 
next

' Finish off with the final line

output = output & "<td style=""border-left:0; border-bottom:0;"" colspan="""
if left(years,4) >= "2006" then 
	output = output & "7"
  elseif left(years,4) >= "1965" then
  	output = output & "5"	
  else
  	output = output & "4"
  	end if
output = output & """></td>"
output = output & outputhold
output = output & "</tr>"
output = output & "</table>"

response.write(output)
%><%'="a" %>


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
  '==    ~ Mark Headington and David Riley, pg. 586    ~     ==
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

  Dim pivot(),loSwap,hiSwap,temp,counter
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