<%@ Language=VBScript %> 
<% Option Explicit %>

<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Author" content="Trevor Scallan">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>GoS-DB Miscellaneous Report</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
#table1 td {border: 1px solid #c0c0c0; text-align:right; white-space:nowrap; margin: 0; padding-left:4; padding-right:4; padding-top:3; padding-bottom:3}
-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs,rsappears,rsgoals, n, outline, startyear, endyear, playerappears, playergoals, playerappears10, playergoals10, weightfactor, greenwidth, yellowwidth, greywidth, weightedpointspergamewidth, percwins, percdraws, percdefeats, bgcolour

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rsappears = Server.CreateObject("ADODB.Recordset")
Set rsgoals = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

weightfactor = Request.Form("weightfactor")
if weightfactor = "" then weightfactor = 1.2
%>
  
  <center>
  <table border="0" cellspacing="0" style="border-collapse: collapse" 
  cellpadding="0" width="980">
    <tr>
    <td width="260" valign="top" style="text-align:center;">

	<p style="text-align: center; margin-top:0; margin-bottom:3">
	<a href="gosdb.asp"><font color="#404040"><img border="0" src="images/gosdb-small.jpg" align="left"></font></a><font color="#404040"> 
	<b><font style="font-size: 15px">Search by<br>
	</font></b><span style="font-size: 15px"><b>Player</b></span></font><p style="text-align: center; margin-top:0; margin-bottom:6">
	<b>
	<a href="gosdb.asp">Back to<br>GoS-DB Hub</a></b></p>

	</td>
    
  	<td width="460" align="center" style="text-align: center" valign="top">	
	<p style="margin-top:12; margin-bottom:0; text-align:center; font-size:18px; color:#006E32">
    MISCELLANEOUS REPORTS</p>  
    
	<p style="margin-top:6; margin-bottom:0; text-align:center; font-size:13px">
    <b>Report 7: Football League Record by Decade</b></p> 
       
    </td>
        
	<td width="260" valign="top"  align="justify">
	'<span style="font-size: 10px">Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
     
    </td>
    </tr>   
	</table>
    
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; margin:3 0 12 0" bordercolor="#111111" width="980">
        <tr>
          <td width="50%" style="border-style: none; border-width: medium" valign="top">
          <p style="margin-top: 0; margin-bottom: 2"><b>Notes:</b></td>
          <td width="50%" style="border-style: none; border-width: medium" valign="top">
          &nbsp;</td>
        </tr>
        <tr>
          <td width="50%" style="border-style: none; border-width: medium" valign="top">
          <p style="text-align: left; margin-top: 0; margin-bottom: 6">1. This 
          report compares Argyle's playing record by decade. Note that the 
          numbers are based on match dates, so some 
          seasons appear across two decades.</p>
          <p style="text-align: left; margin-top: 0; margin-bottom: 6">2. 
          The average home attendance gives a reasonable indication of a 
          successful team, although oddly, the best attended period coincides 
          with some of the worst results. These were the post-WW2 years, when football 
          was hugely popular after many years of austerity.</p>
          <p style="text-align: left; margin-top: 0; margin-bottom: 6">3. A 
          better indication of success is the proportion of games won. The green, yellow and grey bars illustrate the percentage of 
          games in that decade that were won, drawn and lost (detailed figures 
          appear when you hover on a bar).&nbsp; </p>
          <p style="text-align: left; margin-top: 0; margin-bottom: 6">4. The 
          blue bars need some explanation. A simple total of points is not much 
          use when comparing seasons because the number of games  varies. 
          The average points per game (using 3 points for a win in all cases) is 
          a better indicator, i.e. total points divided by total games. However, 
          most would agree that a win in a higher division is more impressive 
          than in a lower one, so without some form of a weighting, we 
          still don't get a true view. The blue bars are the result of applying 
          a  weighting factor of 1.2 in favour of each higher division.
          <i>(continued in next column)</i></td>
          <td width="50%" style="border-style: none; border-width: medium" valign="top">
          <p style="text-align: left; margin-top: 0; margin-bottom: 0; margin-left:6">
          So how does this work? A win in tier-4 gets 3 points, a win in tier-3 gets 1.2 x 3 points 
          and a win in tier-2 gets 1.2 x 1.2 x 3 points. This method is also 
          used for draws. 
          Thus a&nbsp;record in a higher tier is given more weight than a 
          similar one lower down. Of course the big question is what should 
          that factor be? How much better is each tier compared with the next 
          lower one? The report 
          starts with a factor of 1.2 but to be honest, it's a guess. And by the 
          way, reducing the factor to 1 is the same as no weighting, so the bars 
          reduce to the actual points gained.<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.weightfactor.value == "")
  {
    alert("Please enter a value for the \"Weight Factor\" field.");
    theForm.weightfactor.focus();
    return (false);
  }

  if (theForm.weightfactor.value.length > 3)
  {
    alert("Please enter at most 3 characters in the \"Weight Factor\" field.");
    theForm.weightfactor.focus();
    return (false);
  }

  var checkOK = "0123456789-.";
  var checkStr = theForm.weightfactor.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Weight Factor\" field.");
    theForm.weightfactor.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Please enter a valid number in the \"weightfactor\" field.");
    theForm.weightfactor.focus();
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseFloat(allNum);
  if (chkVal != "" && !(prsVal <= 2 && prsVal >= 1))
  {
    alert("Please enter a value less than or equal to \"2\" and greater than or equal to \"1\" in the \"Weight Factor\" field.");
    theForm.weightfactor.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form style="padding: 0; margin: 0;" method="POST" action="gosdb-misc7.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1" webbot-action="--WEBBOT-SELF--">
            <!--webbot bot="SaveResults" u-file="_private/form_results.csv" s-format="TEXT/CSV" s-label-fields="TRUE" startspan --><input TYPE="hidden" NAME="VTI-GROUP" VALUE="0"><!--webbot bot="SaveResults" endspan i-checksum="43374" -->
            <p style="margin-left: 6; margin-top:3; margin-bottom:0">
            You can change the weighting factor here:
            <!--webbot bot="Validation" s-display-name="Weight Factor" s-data-type="Number" s-number-separators="x." b-value-required="TRUE" i-maximum-length="3" s-validation-constraint="Less than or equal to" s-validation-value="2" s-validation-constraint="Greater than or equal to" s-validation-value="1" --><input type="text" name="weightfactor" size="1" value="<%response.write(weightfactor)%>" maxlength="3" style="font-size: 10px"><input type="submit" value="Change" name="B1" style="font-size: 10px"> 
            (between 1 and 2). </p>
            <p style="margin-left: 6"></p>
          </form>
          <p style="margin-left: 6">
          </p>
          <p style="text-align: left; margin-top: 0; margin-bottom: 0; margin-left:6">
          Like the bars in the previous column, detailed figures appear when you hover on a 
          a blue bar.<p style="text-align: left; margin-top: 6; margin-bottom: 0; margin-left:6">
          5. So 
          what are the best and worst decades? From the number of games won, 
          the 1920s is the best, but this was in tier-3. With a small rise in 
          the weighting factor, the 1930's soon comes into its own. The worst 
          season is a close-run thing; at least using the default weighting. Despite the high attendances, the four seasons in 
          the 1940s were not the best, nor were the 1990s, but it's the latter 
          case that takes the wooden spoon if you raise the weighting factor. </td>
        </tr>
    </table>	    
	
    <table id="table1" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
    <tr>
      <td style="text-align:left;" rowspan="2" valign="bottom"><b>Decade</b></td>
      <td colspan="4" style="border-bottom-style: none; border-bottom-width: medium" valign="bottom">
      <b>Seasons in Tiers</b></td>
      <td valign="bottom" rowspan="2"><b>Average<br>
&nbsp;Home<br>
	  Attend.</b></td>
      <td valign="bottom" rowspan="2" style="text-align: left"><b>Top 
      Player This<br>
      Decade <br>
      </b>Hover in cell for top 10</td>
      <td style="text-align:left;" rowspan="2" valign="bottom">
      <font color="#FFFFFF"><span style="background-color: #008000">&nbsp;Won
      </span></font><span style="background-color: #FFFF99"> 
      <font color="#FFCC66">&nbsp;</font>Drawn<font color="#FFCC66"> </font>
      </span><span style="background-color: #808080">
      <font color="#FFFFFF">&nbsp;Lost </font></span><b>&nbsp;Ratios<br>
      </b>Hover in cell for values</td>
      <td style="text-align:left; border-bottom-style:none; border-bottom-width:medium" valign="bottom" rowspan="2">
      <font color="#333399"><b>
      Weighted average
      points<br>
      per game </b>(see note 4)<b><br>
      </b></font>Hover in cell for values</td>
    </tr>
    <tr>
      <td style="border-top-style: none; border-top-width: medium" valign="bottom">
      <b>1</b></td>
      <td style="border-top-style: none; border-top-width: medium" valign="bottom">
      <b>2</b></td>
      <td style="border-top-style: none; border-top-width: medium" valign="bottom">
      <b>3</b></td>
      <td style="border-top-style: none; border-top-width: medium" valign="bottom">
      <b>4</b></td>
    </tr>
	<%
	outline = ""
	sql = "WITH CTE1 as "
	sql = sql & "(select date, 1 as matches, "
	sql = sql & "row_number() over (partition by years order by date) as game, "
	sql = sql & "case when goalsfor > goalsagainst then 1 else 0 end as wins, "
	sql = sql & "case when goalsfor = goalsagainst then 1 else 0 end as draws, "
	sql = sql & "case when goalsfor < goalsagainst then 1 else 0 end as defeats, "
	sql = sql & "case when goalsfor > goalsagainst then 3 "
	sql = sql & "	 when goalsfor = goalsagainst then 1 "
	sql = sql & "	 else 0  "
	sql = sql & "	 end as modernpoints, "
	sql = sql & "case when homeaway = 'H' then attendance else NULL end as attendance "
	sql = sql & "from [v_match_FL-39]  a join season b "
	sql = sql & "on a.date >= b.date_start and a.date <= b.date_end "  
	sql = sql & "), "
	sql = sql & "CTE2 AS "
	sql = sql & "(select 1 as matches, decade, "
	sql = sql & "case when game = 1 and tier = 1 then 1 else 0 end as tier1, "
	sql = sql & "case when game = 1 and tier = 2 then 1 else 0 end as tier2, "
	sql = sql & "case when game = 1 and tier = 3 then 1 else 0 end as tier3, "
	sql = sql & "case when game = 1 and tier = 4 then 1 else 0 end as tier4, "
	sql = sql & "wins, draws, defeats, modernpoints, power(" & weightfactor & ",4-tier)*modernpoints as weightedpoints, "
	sql = sql & "attendance "
	sql = sql & "from CTE1 a join season b "
	sql = sql & "on a.date >= b.date_start and a.date <= b.date_end "  
	sql = sql & ") "
	sql = sql & "select case when grouping(decade) = 1 then 'All' else decade end as decade, "
	sql = sql & "sum(tier1) as tier1, sum(tier2) as tier2, sum(tier3) as tier3, sum(tier4) as tier4, "
	sql = sql & "sum (matches) as matches, sum(wins) as wins, sum(draws) as draws, sum(defeats) as defeats, "
	sql = sql & "sum(modernpoints)*100/sum(matches) as modernpoints, sum(weightedpoints)*100/sum(matches) as weightedpoints, "
	sql = sql & "avg(attendance) as attendance "
	sql = sql & "from CTE2 "
	sql = sql & "group by decade with rollup "
	sql = sql & "order by decade "
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
	
		percwins = round(100*rs.Fields("wins")/rs.Fields("matches"),1)
		percdraws = round(100*rs.Fields("draws")/rs.Fields("matches"),1)
		percdefeats = round(100*rs.Fields("defeats")/rs.Fields("matches"),1)
		
		greenwidth = int(200*rs.Fields("wins")/rs.Fields("matches"))
		yellowwidth = int(200*rs.Fields("draws")/rs.Fields("matches"))
		greywidth = int(200*rs.Fields("defeats")/rs.Fields("matches"))

		weightedpointspergamewidth = int(CStr(rs.Fields("weightedpoints")))
		
		if rs.Fields("decade") <> "All" then
			startyear = left(rs.Fields("decade"),4)
			endyear = left(rs.Fields("decade"),2) & right(rs.Fields("decade"),2)
		  else
		    startyear = "1920"
		    endyear = year(Date)
		end if
		
		sql = "select top 10 player_id_spell1, surname, initials, count(distinct b.date) as appears "
		sql = sql & "from [v_match_FL-39] a join match_player b on a.date = b.date " 
		sql = sql & "join player d on b.player_id = d.player_id "
		sql = sql & "where year(a.date) between '" & startyear & "' and '" & endyear & "' "
		sql = sql & "  and startpos > 0 "
		sql = sql & "group by player_id_spell1, surname, initials "
		sql = sql & "order by appears desc "
		rsappears.open sql,conn,1,2

		playerappears = ""
		playerappears10 = "Starts: "
		n = 0
		Do While Not rsappears.EOF
	  	playerappears10 = playerappears10 & left(trim(rsappears.Fields("initials")),1) & "." & trim(rsappears.Fields("surname")) & " (" & rsappears.Fields("appears") & "), "
  	  	if playerappears = "" then playerappears = left(playerappears10,len(playerappears10)-2)
  	  	n = n + 1
  	  	if n = 5 then playerappears10 = playerappears10 & "<br>"  	  	
  	  	rsappears.MoveNext
		Loop
		rsappears.close
		playerappears10 = left(playerappears10,len(playerappears10)-2)	'drop final comma
		
		sql = "select top 10 player_id_spell1, surname, initials, count(distinct b.date) as goals "
		sql = sql & "from [v_match_FL-39] a join match_goal b on a.date = b.date " 
		sql = sql & "join player d on b.player_id = d.player_id "
		sql = sql & "where year(a.date) between '" & startyear & "' and '" & endyear & "' "
		sql = sql & " and player_id_spell1 < 9000 "
		sql = sql & "group by player_id_spell1, surname, initials "
		sql = sql & "order by goals desc "
		rsgoals.open sql,conn,1,2
	
		playergoals = ""
		playergoals10 = "<br>Goals: "
		n = 0
		Do While Not rsgoals.EOF
	  	playergoals10 = playergoals10 & left(trim(rsgoals.Fields("initials")),1) & "." & trim(rsgoals.Fields("surname")) & " (" & rsgoals.Fields("goals") & "), "
  	  	if playergoals = "" then playergoals = left(playergoals10,len(playergoals10)-2)
  	  	n = n + 1
  	  	if n = 5 then playergoals10 = playergoals10 & "<br>"
  	  	rsgoals.MoveNext
		Loop
		rsgoals.close
		playergoals10 = left(playergoals10,len(playergoals10)-2)	'drop final comma
		
		bgcolour = "#ffffff"
		if left(rs.Fields("decade"),3) = "All" then bgcolour = "#e0e0e0"
				
		outline  = outline & "<tr style=""background-color:" & bgcolour & """>"
		outline  = outline & "<td style=""text-align:left;""><b>" & rs.Fields("decade") & "</b></td>"
		outline  = outline & "<td>" & rs.Fields("tier1") & "</td>"
		outline  = outline & "<td>" & rs.Fields("tier2") & "</td>"
		outline  = outline & "<td>" & rs.Fields("tier3") & "</td>"
		outline  = outline & "<td>" & rs.Fields("tier4") & "</td>"
		outline  = outline & "<td>" & FormatNumber(rs.Fields("attendance"),0) & "</td>"
		outline  = outline & "<td style=""text-align:left;"" onmouseover=""showtip('" & rs.Fields("decade") & ": <br>" & playerappears10 & playergoals10 & "')"" onmouseout=""hidetip()"">"
		outline  = outline & playerappears & playergoals & "</td>"
		outline  = outline & "<td onmouseover=""showtip('" & rs.Fields("decade") & ": <br>"
		outline  = outline & rs.Fields("matches") & " matches<br>"
		outline  = outline & rs.Fields("wins") & " wins (" & percwins & "%)<br>" 
		outline  = outline & rs.Fields("draws") & " draws (" & percdraws & "%)<br>"
		outline  = outline & rs.Fields("defeats") & " defeats (" & percdefeats & "%)')"" onmouseout=""hidetip()"">" 
		outline  = outline & "<img border=""0"" src=""images/green_1x1.gif"" height=""10"" width=""" & greenwidth & """>"
		outline  = outline & "<img border=""0"" src=""images/yellow_1x1.gif"" height=""10"" width=""" & yellowwidth & """>"
		outline  = outline & "<img border=""0"" src=""images/darkgrey_1x1.gif"" height=""10"" width=""" & greywidth & """>"
		outline  = outline & "</td>"
		outline  = outline & "<td style=""text-align:left;"" onmouseover=""showtip('" & rs.Fields("decade") & ": <br>"
		outline  = outline & "Average points per game: " & round(rs.Fields("modernpoints")/100,2) & "<br>"
		outline  = outline & "Applying the weighting factor (" & weightfactor & ") for higher tiers,<br>average points per game: " & round(Cstr(rs.Fields("weightedpoints"))/100,2) & "')"" "  		
		outline  = outline & "onmouseout=""hidetip()"">"
		outline  = outline & "<img style=""margin-top:1"" border=""0"" src=""images/blue2_1x1.gif"" height=""10"" width=""" & weightedpointspergamewidth/1.5 & """>"
		outline  = outline & "</td>"
		outline  = outline & "</tr>"
  		rs.MoveNext
	Loop
		
	rs.close
	response.write(outline)

conn.close
%>	
	
</table>
</center><br>

<!--#include file="base_code.htm"-->
</body>

</html>