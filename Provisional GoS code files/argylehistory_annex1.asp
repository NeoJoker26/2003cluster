<%

output = "<p class=""head1"">AN ARGYLE TIMELINE</p>"
output = output & "<p>Key moments in Argyle's history, season by season.</p>"
output = output & "<p><b>Version:</b> 1.0</p>"
output = output & "<p><b>Date:</b> 23 Jun 2011</p>"

output = output & "<p class=""noprint"" style=""margin:0 0 6 0;""><a href=""argylehistorymenu.asp"">Return to History Contents</a></p>"

output = output & "</div>"

 
' build timeline

output = output & "<div id=""chart1"" align=""left"">"

output = output & "<img src=""images/transparentdot.gif"" style=""height:17;width:1""><br>"
output = output & "<span style=""position:absolute;top:2;left:29;font-size:8px;"">1890-00</span>"
output = output & "<span style=""position:absolute;top:2;left:81;font-size:8px;"">1900-10</span>"
output = output & "<span style=""position:absolute;top:2;left:133;font-size:8px;"">1910-20</span>"
output = output & "<span style=""position:absolute;top:2;left:185;font-size:8px;"">1920-30</span>"
output = output & "<span style=""position:absolute;top:2;left:237;font-size:8px;"">1930-40</span>"
output = output & "<span style=""position:absolute;top:2;left:289;font-size:8px;"">1940-50</span>"
output = output & "<span style=""position:absolute;top:2;left:341;font-size:8px;"">1950-60</span>"
output = output & "<span style=""position:absolute;top:2;left:393;font-size:8px;"">1960-70</span>"
output = output & "<span style=""position:absolute;top:2;left:445;font-size:8px;"">1970-80</span>"
output = output & "<span style=""position:absolute;top:2;left:497;font-size:8px;"">1980-90</span>"
output = output & "<span style=""position:absolute;top:2;left:549;font-size:8px;"">1990-00</span>"
output = output & "<span style=""position:absolute;top:2;left:601;font-size:8px;"">2000-10</span>"
output = output & "<span style=""position:absolute;top:2;left:653;font-size:8px;"">2010-20</span>"
output = output & "<img style=""position:absolute;top:0;left:0"" src=""images/blackpixel.gif"" height=""1"" width=""697"">"
output = output & "<img style=""position:absolute;top:1;left:0"" src=""images/blackpixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:20"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:72"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:15;left:84"" src=""images/blackpixel.gif"" height=""13"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:124"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:176"" src=""images/blackpixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:228"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:280"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:332"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:384"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:436"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:488"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:540"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:592"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:644"" src=""images/greypixel.gif"" height=""26"" width=""1"">"
output = output & "<img style=""position:absolute;top:1;left:696"" src=""images/blackpixel.gif"" height=""26"" width=""1"">"

output = output & "<p style=""position:absolute;left:5;top:7;font-size:9px;z-index:9"">Local Football</p>"
output = output & "<p style=""position:absolute;left:88;top:7;font-size:9px;z-index:9"">Southern League</p>"
output = output & "<p style=""position:absolute;left:181;top:7;font-size:9px;z-index:9"">Football League</p>"

output = output & "</div>"

output = output & "<div id=""chart2"" align=""left"">"

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

sql = "select years, promrel, endpos, tier, (teams_in_div-endpos) as height, (92-teams_above_div-teams_in_div) as bottompos, division_short, teams_in_div "
sql = sql & "from season "
sql = sql & "union all "
sql = sql & "select years, null, null, null, null, null, null, null "
sql = sql & "from fill_in_years "
sql = sql & "order by years "
  
rs.open sql,conn,1,2
  

output = output & "<img style=""position:absolute;top:0"" & src=""images/blackpixel.gif"" height=""1"" width=""697"">"
output = output & "<img style=""position:absolute;top:66;left:177"" src=""images/greypixel.gif"" height=""1"" width=""349"">"
output = output & "<img style=""position:absolute;top:63;left:525"" src=""images/greypixel.gif"" height=""1"" width=""5"">"
output = output & "<img style=""position:absolute;top:63;left:525"" src=""images/greypixel.gif"" height=""4"" width=""1"">"
output = output & "<img style=""position:absolute;top:60;left:530"" src=""images/greypixel.gif"" height=""4"" width=""1"">"
output = output & "<img style=""position:absolute;top:60;left:530"" src=""images/greypixel.gif"" height=""1"" width=""16"">"
output = output & "<img style=""position:absolute;top:60;left:546"" src=""images/greypixel.gif"" height=""6"" width=""1"">"
output = output & "<img style=""position:absolute;top:66;left:546"" src=""images/greypixel.gif"" height=""1"" width=""21"">"
output = output & "<img style=""position:absolute;top:60;left:566"" src=""images/greypixel.gif"" height=""6"" width=""1"">"
output = output & "<img style=""position:absolute;top:60;left:566"" src=""images/greypixel.gif"" height=""1"" width=""131"">"
output = output & "<img style=""position:absolute;top:132;left:87"" src=""images/blackpixel.gif"" height=""54"" width=""1"">"
output = output & "<img style=""position:absolute;top:132;left:88"" src=""images/blackpixel.gif"" height=""1"" width=""89"">"
output = output & "<img style=""position:absolute;top:132;left:177"" src=""images/greypixel.gif"" height=""1"" width=""369"">"
output = output & "<img style=""position:absolute;top:132;left:546"" src=""images/greypixel.gif"" height=""6"" width=""1"">"
output = output & "<img style=""position:absolute;top:138;left:546"" src=""images/greypixel.gif"" height=""1"" width=""21"">"
output = output & "<img style=""position:absolute;top:132;left:566"" src=""images/greypixel.gif"" height=""6"" width=""1"">"
output = output & "<img style=""position:absolute;top:132;left:566"" src=""images/greypixel.gif"" height=""1"" width=""131"">"
output = output & "<img style=""position:absolute;top:186;left:87"" src=""images/blackpixel.gif"" height=""1"" width=""20"">"
output = output & "<img style=""position:absolute;top:186;left:103"" src=""images/blackpixel.gif"" height=""6"" width=""1"">"
output = output & "<img style=""position:absolute;top:192;left:103"" src=""images/blackpixel.gif"" height=""1"" width=""10"">"
output = output & "<img style=""position:absolute;top:192;left:113"" src=""images/blackpixel.gif"" height=""3"" width=""1"">"
output = output & "<img style=""position:absolute;top:195;left:113"" src=""images/blackpixel.gif"" height=""1"" width=""5"">"
output = output & "<img style=""position:absolute;top:195;left:118"" src=""images/blackpixel.gif"" height=""3"" width=""1"">"
output = output & "<img style=""position:absolute;top:198;left:118"" src=""images/blackpixel.gif"" height=""1"" width=""6"">"
output = output & "<img style=""position:absolute;top:193;left:124"" src=""images/blackpixel.gif"" height=""6"" width=""1"">"
output = output & "<img style=""position:absolute;top:192;left:124"" src=""images/blackpixel.gif"" height=""1"" width=""46"">"
output = output & "<img style=""position:absolute;top:198;left:170"" src=""images/blackpixel.gif"" height=""1"" width=""162"">"
output = output & "<img style=""position:absolute;top:193;left:169"" src=""images/blackpixel.gif"" height=""6"" width=""1"">"
output = output & "<img style=""position:absolute;top:204;left:332"" src=""images/blackpixel.gif"" height=""1"" width=""41"">"
output = output & "<img style=""position:absolute;top:204;left:375"" src=""images/greypixel.gif"" height=""1"" width=""171"">"
output = output & "<img style=""position:absolute;top:204;left:546"" src=""images/greypixel.gif"" height=""6"" width=""1"">"
output = output & "<img style=""position:absolute;top:210;left:546"" src=""images/greypixel.gif"" height=""1"" width=""21"">"
output = output & "<img style=""position:absolute;top:204;left:566"" src=""images/greypixel.gif"" height=""6"" width=""1"">"
output = output & "<img style=""position:absolute;top:204;left:566"" src=""images/greypixel.gif"" height=""1"" width=""125"">"
output = output & "<img style=""position:absolute;top:201;left:690"" src=""images/greypixel.gif"" height=""1"" width=""7"">"
output = output & "<img style=""position:absolute;top:201;left:690"" src=""images/greypixel.gif"" height=""3"" width=""1"">"
output = output & "<img style=""position:absolute;top:276;left:374"" src=""images/blackpixel.gif"" height=""1"" width=""318"">"
output = output & "<img style=""position:absolute;top:273;left:691"" src=""images/blackpixel.gif"" height=""1"" width=""6"">"
output = output & "<img style=""position:absolute;top:0;left:176"" src=""images/blackpixel.gif"" height=""132"" width=""1"">"
output = output & "<img style=""position:absolute;top:197;left:332"" src=""images/blackpixel.gif"" height=""7"" width=""1"">"
output = output & "<img style=""position:absolute;top:204;left:373"" src=""images/blackpixel.gif"" height=""73"" width=""1"">"
output = output & "<img style=""position:absolute;top:272;left:691"" src=""images/blackpixel.gif"" height=""4"" width=""1"">"

    
Do While Not rs.EOF

    if trim(rs.Fields("division_short")) = "SL" then
 	  	output = output & "<img style=""position:relative;bottom:" & 78+3*(22-rs.Fields("teams_in_div")) & "; width:4; height:" & 1+3*rs.Fields("height") & """ src=""images/lightgreenpixel.gif"">"
 	  elseif isnull(rs.Fields("tier")) or isnull(rs.Fields("endpos")) then
 		output = output & "<img src=""images/transparentdot.gif"" style=""height:0;width:4"">"
 	  else
 	  	output = output & "<img style=""position:relative;bottom:" & 3*rs.Fields("bottompos") & "; width:4; height:" & 2+3*rs.Fields("height") & """ src=""images/greenpixel.gif"">"
	end if  
	
	if right(rs.Fields("years"),1) = "0" then 
			  if rs.Fields("years") < "1909-1910" then 
				output = output & "<img src=""images/transparentdot.gif"" height=""1"" width=""1"" style=""position:relative;bottom:12"">"
			  elseif rs.Fields("years") < "1919-1920" then 
				output = output & "<img src=""images/greypixel.gif"" height=""59"" width=""1"" style=""position:relative;bottom:84"">"
			  elseif rs.Fields("years") < "1929-1930" then 
				output = output & "<img src=""images/blackpixel.gif"" height=""65"" width=""1"" style=""position:relative;bottom:78"">"
			  elseif rs.Fields("years") < "1959-1960" then 
				output = output & "<img src=""images/greypixel.gif"" height=""197"" width=""1"" style=""position:relative;bottom:78"">"
			  elseif rs.Fields("years") = "2019-2020" then 
				output = output & "<img src=""images/greypixel.gif"" height=""272"" width=""1"" style=""position:relative;bottom:3"">"
			  else
			  	output = output & "<img src=""images/greypixel.gif"" height=""276"" width=""1"">"
			end if  
	end if
	
rs.MoveNext
Loop

rs.close

output = output & "<img style=""position:absolute;top:0;left:696"" src=""images/blackpixel.gif"" height=""273"" width=""1"">"	'final upright

output = output & "<p style=""top:5; left:183"">Tier 1</p>"
output = output & "<p style=""top:71; left:183"">Tier 2</p>"
output = output & "<p style=""top:137; left:183; color:white"">Tier 3</p>"
output = output & "<p style=""top:209; left:380"">Tier 4</p>"
output = output & "<p style=""top:137; left:93; color:gray;"">Southern</p>"
output = output & "<p style=""top:147; left:93; color:gray;"">League</p>"
output = output & "<p style=""top:137; left:4; color:gray;"">Friendlies &</p>"
output = output & "<p style=""top:147; left:4; color:gray;"">Devon</p>"
output = output & "<p style=""top:157; left:4; color:gray;"">Leagues</p>"

output = output & "<span class=""letred"" style=""left:0;top:0;"">A</span>"
output = output & "<span class=""letred"" style=""left:85;top:0;"">B</span>"
output = output & "<span class=""letred"" style=""left:254;top:0;"">D</span>"
output = output & "<span class=""letred"" style=""left:403;top:0;"">E</span>"
output = output & "<span class=""letred"" style=""left:450;top:0;"">E</span>"
output = output & "<span class=""letred"" style=""left:502;top:0;"">F</span>"
output = output & "<span class=""letred"" style=""left:566;top:0;"">G</span>"
output = output & "<span class=""letred"" style=""left:645;top:0;"">H</span>"
output = output & "<span class=""letgreen"" style=""left:133;top:120"">C</span>"
output = output & "<span class=""letgreen"" style=""left:220;top:120;"">C</span>"
output = output & "<span class=""letgreen"" style=""left:335;top:120;"">C</span>"
output = output & "<span class=""letgreen"" style=""left:370;top:120;"">C</span>"
output = output & "<span class=""letgreen"" style=""left:595;top:192;"">C</span>"
output = output & "<span class=""letgreen"" style=""left:605;top:120;"">C</span>"
output = output & "<span class=""letgreen"" style=""left:221;top:112;"">P</span>"
output = output & "<span class=""letgreen"" style=""left:336;top:112;"">P</span>"
output = output & "<span class=""letgreen"" style=""left:371;top:112;"">P</span>"
output = output & "<span class=""letgreen"" style=""left:513;top:125;"">P</span>"
output = output & "<span class=""letgreen"" style=""left:456;top:125;"">P</span>"
output = output & "<span class=""letgreen"" style=""left:566;top:203;"">P</span>"
output = output & "<span class=""letgreen"" style=""left:596;top:184;"">P</span>"
output = output & "<span class=""letgreen"" style=""left:606;top:112;"">P</span>"
output = output & "<span class=""letgreen"" style=""left:675;top:196;"">P</span>"
output = output & "<span class=""letgreen"" style=""left:691;top:196;"">P</span>"
output = output & "<span class=""letblack"" style=""left:326;top:116;"">R</span>"
output = output & "<span class=""letblack"" style=""left:358;top:116;"">R</span>"
output = output & "<span class=""letblack"" style=""left:420;top:119;"">R</span>"
output = output & "<span class=""letblack"" style=""left:467;top:116;"">R</span>"
output = output & "<span class=""letblack"" style=""left:546;top:119;"">R</span>"
output = output & "<span class=""letblack"" style=""left:561;top:188;"">R</span>"
output = output & "<span class=""letblack"" style=""left:576;top:185;"">R</span>"
output = output & "<span class=""letblack"" style=""left:638;top:116;"">R</span>"
output = output & "<span class=""letblack"" style=""left:645;top:188;"">R</span>"
output = output & "<span class=""letblack"" style=""left:685;top:182;"">R</span>"

output = output & "</div>"

output = output & "<div id=""chart3"">"

output = output & "<p style=""font-size:11px; margin:0 3px 4px 2px;""><b>Key to major events</b></p>"

output = output & "<table style=""float:left; border:0; border-collapse:collapse; width:250px"">"
output = output & "<tr>"
output = output & "<td width=""13"" valign=""top"" style=""color:red"">A:</td>"
output = output & "<td>Argyle Football Club formed</td>"
output = output & "</tr><tr>"
output = output & "<td  valign=""top"" style=""color:red"">B:</td>"
output = output & "<td>Club turned professional; AFC became PAFC</td>"
output = output & "</tr><tr>"
output = output & "<td  valign=""top"" style=""color:red"">D:</td>"
output = output & "<td>Record Home Park attendance</td>"
output = output & "</tr><tr>"
output = output & "<td  valign=""top"" style=""color:red"">E:</td>"
output = output & "<td>League Cup semi-finals</td>"
output = output & "</tr><tr>"
output = output & "<td  valign=""top"" style=""color:red"">F:</td>"
output = output & "<td>FA Cup semi-finals</td>"
output = output & "</tr><tr>"
output = output & "<td  valign=""top"" style=""color:red"">G:</td>"
output = output & "<td>Wembley appearance</td>"
output = output & "</tr><tr>"
output = output & "<td  valign=""top"" style=""color:red"">H:</td>"
output = output & "<td>Points deduction and administration</td>"
output = output & "</tr><tr>"
output = output & "<td  valign=""top"" style=""color:green"">C:</td>"
output = output & "<td>Division champions</td>"
output = output & "</tr><tr>"
output = output & "<td  valign=""top"" style=""color:green"">P:</td>"
output = output & "<td>Promotion</td>"
output = output & "</tr><tr>"
output = output & "<td  valign=""top"" style=""color:black"">R:</td>"
output = output & "<td>Relegation</td>"
output = output & "</tr>"
output = output & "</table>"
 
sql = "select avg(endpos+teams_above_div) as avgpos, max(endpos+teams_above_div) as maxpos, min(endpos+teams_above_div) as minpos "
sql = sql & "from season "
sql = sql & "where tier is not null and endpos is not null "
  
rs.open sql,conn,1,2
avgpos = rs.Fields("avgpos")
maxpos = rs.Fields("maxpos")
minpos = rs.Fields("minpos")
rs.close

for n = 0 to Ubound(tiers)
 tiers(n) = 0
next

sql = "select tier, count(*) as count "
sql = sql & "from season "
sql = sql & "where tier is not null and endpos is not null "
sql = sql & "group by tier "
sql = sql & "order by tier "
  
rs.open sql,conn,1,2
Do While Not rs.EOF
 tiers(rs.Fields("tier")) = rs.Fields("count")
 rs.MoveNext
Loop
rs.close

conn.Close

output = output & "<div id=""chart4"">"
output = output & "<p style=""margin:0""><b>Football League Record</b></p>"
output = output & "<p style=""margin:0 0 4px; font-size:10px;"">(excluding the abandoned 1939-40 season, the<br>1945-46 FL South and the current campaign)</p>"

output = output & "<table style=""float:left; border:0; border-collapse:collapse; width:30px"">"
output = output & "<tr>"
output = output & "<td style=""width:14px""><u>Tier</u></td>"
output = output & "<td><u>#Seasons</u></td>"
output = output & "</tr><tr>"
output = output & "<td>1:</td>"
output = output & "<td class=""right"">" & tiers(1) & "</td>"
output = output & "</tr><tr>"
output = output & "<td>2:</td>"
output = output & "<td class=""right"">" & tiers(2) & "</td>"
output = output & "</tr><tr>"
output = output & "<td>3:</td>"
output = output & "<td class=""right"">" & tiers(3) & "</td>"
output = output & "</tr><tr>"
output = output & "<td>4:</td>"
output = output & "<td class=""right"">" & tiers(4) & "</td>"
output = output & "</tr><tr>"
output = output & "<td>All:</td>"
output = output & "<td class=""right"">" & int(tiers(1)) + int(tiers(2)) + int(tiers(3)) + int(tiers(4)) & "</td>"
output = output & "</tr>"
output = output & "</table>"

output = output & "<table style=""float:right; border:0; border-collapse:collapse; width:110px"">"
output = output & "<tr>"
output = output & "<td colspan=""2""><u>End Season Position</u></td>"
output = output & "</tr><tr>"
output = output & "<td style=""width:70px"">Best:</td>"
output = output & "<td class=""right"" style=""width:40px"" >" & minpos & "</td>"
output = output & "</tr><tr>"
output = output & "<td>Worst:</td>"
output = output & "<td class=""right"">" & maxpos & "</td>"
output = output & "</tr><tr>"
output = output & "<td>Average:</td>"
output = output & "<td class=""right"">" & avgpos & "</td>"
output = output & "</tr>"
output = output & "</table>"

output = output & "</div>"

output = output & "</div>"


'end of timeline

'formheading = "<p style=""margin: 0 0 6 0;"">Feel free to leave your thoughts about this annex by completing the boxes below. If you have a general comment to make about the history project as a whole, please use the feedback facility on the <a href=""argylehistorymenu.asp""><u>contents page</u></a>.</p>"
%>
<!--#include file="argylehistory_form.asp"-->
<%

%>