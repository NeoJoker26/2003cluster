<%@ Language=VBScript %> <% Option Explicit %>
<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>
<script language="javascript">

function Waiting(item) {
   obj=document.getElementById(item);
   obj.style.display="block";
   }
</script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Author" content="Trevor Scallan">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<title>GoS-DB Players</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--

#ajaxplayerlist p {margin: 0 0 0 12; padding: 0;  font-size: 11px; font-weight:normal; text-align: left; color:#202020; }

#ajaxplayerdetails p {margin: 0 0 3 0; padding: 0;  font-size: 11px; font-weight:normal; text-align: left; }
#ajaxplayerdetails .name {width: 100%; margin: 0 0 9 0; padding: 2 4 2 4; color: #ffffff; background-color: #404040; font-size: 14px; font-weight:bold; text-align: left; }

#gottable1 td {text-align:left; margin: 0; padding: 0 2 0 2;  font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 12px; }
#gottable1 .right {text-align: right; } 
#gottable1 .bold {font-weight: bold; }
#gottable1 .head1 {font-family: verdana,arial,sans-serif; font-size: 11px; padding: 4 4 4 4; } 
#gottable1 .head2 {font-family: verdana,arial,sans-serif; font-size: 11px; } 

#gottable2 td {text-align:left; margin: 0; padding: 0 2 0 2;  font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 12px; } 
#gottable2 .right {text-align: right; }
#gottable2 .bold {font-weight: bold; }
#gottable2 .head1 {font-family: verdana,arial,sans-serif; font-size: 11px; padding: 4 4 4 4; }
#gottable2 .head2 {font-family: verdana,arial,sans-serif; font-size: 11px; } 
-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->

  <center>
  	<table border="0" cellspacing="0" style="border-collapse: collapse" 
  	cellpadding="0" width="980">
    <tr>
    <td width="260" valign="top" align="center">
    <p style="text-align: center; margin-top:0; margin-bottom:3">
	<a href="gosdb.asp"><font color="#404040"><img border="0" src="images/gosdb-small.jpg" align="left"></font></a><font color="#404040"> 
	<b><font style="font-size: 15px">Search by<br>
	</font></b><span style="font-size: 15px"><b>Player</b></span></font><p style="text-align: center; margin-top:0; margin-bottom:0">
	<b>
	<a href="gosdb.asp">Back to<br>GoS-DB Hub</a></b></p>

    <%
    Dim objFSO, photoname, photofound, photocaption, player_no, conn, rs, sql, wenttoyear
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    
    Set conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	%><!--#include file="conn_read.inc"--><%
  
    Do Until photofound = "True"
    	Randomize
    	player_no = right("000" & int(rnd*1000)+1,3)
    	photoname = "gosdb/photos/players/" & player_no & ".jpg" 
    	photofound = objFSO.FileExists(Server.MapPath(photoname))
    	
    	if photofound = "True" then
    		sql = "select player_id_spell1, initials, forename, surname, min(first_game_year) as first_game_year, max(last_game_year) as last_game_year "
			sql = sql & "from player "
			sql = sql & "where player_id = " & player_no
			sql = sql & "group by player_id_spell1, initials, forename, surname "
			rs.open sql,conn,1,2
			if rs.Fields("first_game_year") < 1960 then
				photocaption = "At Random: <a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & "&scp=1,2,3,4,5,6,7"">" 
				if IsNull(rs.Fields("forename")) then
					photocaption = photocaption & rs.Fields("initials") & " " & trim(rs.Fields("surname"))
	  			  else
	  				photocaption = photocaption & rs.Fields("forename") & " " & trim(rs.Fields("surname"))
	  			end if
				photocaption = photocaption &  " " & rs.Fields("first_game_year") & "-" & rs.Fields("last_game_year") & "</a>"
			  else
			  	photofound = "False"
			end if
			rs.close
		end if	
    Loop
    
    response.write("<p style=""margin:12 0 4 0""><img border=""0"" width=""251px"" height=""351px"" src=""" & photoname & """></p>")
    response.write("<p style=""margin:4 0 6 0"">" & photocaption & "</p>")
    photofound = ""
    
    %>

    </td>
    
    <td width="460" valign="top" align="center">
    
    <form style="padding:0; margin:0;" action="gosdb-players1.asp" method="post" name='form1'>
    
    <table border="0" style="border-collapse: collapse" 
  	cellpadding="0" cellspacing="0" width="360" cellspacing="0">
    <tr>
    <td colspan="2" valign="top" style="text-align: left">
    <p style="margin-top:12; margin-bottom:18; text-align:center; font-size:18px; color:#006E32">
    THE PLAYERS</p>
    </td>
    <tr>
    <td colspan="2" valign="top" style="text-align: left">
    <p style="margin-top: 0; margin-bottom: 6"><b>1. Remove </b>ticks, if 
    appropriate, to affect 
    the counts, calculations and matches revealed on the players' pages. Leave 
    all checked for the complete view.</td>
    <tr>
    <td width="20" valign="top"><input type="checkbox" name="scope" value="1" checked></td>
    <td width="380" valign="top">
    <p style="margin-top: 3; margin-bottom: 0"><b>Southern and Western Leagues: </b>The regular league competitions before entering the Football League in 1920</td>
    </tr><tr>
    <td valign="top"><input type="checkbox" name="scope" value="2" checked></td>
    <td valign="top">
    <p style="margin-top: 3; margin-bottom: 0"><b>Football League: </b>All full seasons from 1920, plus 
    the current season</td>
    </tr><tr>
    <td valign="top"><input type="checkbox" name="scope" value="3" checked></td>
    <td valign="top">
    <p style="margin-top: 3; margin-bottom: 0"><b>Football League 1939: </b>The abandoned 1939-40 season</td>
    </tr><tr>
    <td valign="top"><input type="checkbox" name="scope" value="4" checked></td>
    <td valign="top">
    <p style="margin-top: 3; margin-bottom: 0"><b>War Leagues: </b>The South West Regional League 
    (1939-40) and Football League South (1945-46)</td>
    </tr><tr>
    <td valign="top"><input type="checkbox" name="scope" value="5" checked></td>
    <td valign="top">
    <p style="margin-top: 3; margin-bottom: 0"><b>FA Cup: </b>The Football Association Challenge Cup from 1903 to the present day.</td>
    </tr><tr>
    <td valign="top"><input type="checkbox" name="scope" value="6" checked></td>
    <td valign="top">
    <p style="margin-top: 3; margin-bottom: 0"><b>League Cup: </b>The Football League Cup from its inception in 1960; also known by various 
    sponsors' names - currently the Carling Cup.</td>
    </tr><tr>
    <td valign="top"><input type="checkbox" name="scope" value="7" checked></td>
    <td valign="top">
    <p style="margin-top: 3; margin-bottom: 0"><b>Minor Cups and Trophies: </b>Every other knock-out competition. See foot of page for more details.</td>
    </tr>
       
	<tr>
    <td valign="top" colspan="2">
    <p style="margin-top: 18; margin-bottom: 9"><b>2. Choose </b>one of the 
    following:</p>
	
	<p align="center" style="margin-top: 0; margin-bottom: 0">
	
	<input type="submit" onclick="javascript:Waiting('waiting')" style="width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 10pt; margin: 0; ; padding-left:5; padding-right:5; padding-top:1; padding-bottom:1" 
    name="button" value="Player Rankings">
	<input type="submit" onclick="javascript:Waiting('waiting')" style="width: auto; overflow: visible; color: #000000; background-color: #e0f0e0; font-size: 10pt; margin: 0; ; padding-left:5; padding-right:5; padding-top:1; padding-bottom:1" 
	name="button" value="Player Search">
    
    </form>
        
    </p>

    <p id="waiting" align="center" style="margin-top: 12; margin-bottom: 0; display: none;">
    <b>Please wait ...</b></p>
        
    </td>
    </tr>
    
	</table>
	</td>
	    
	<td width="260" valign="top"  align="center">
    <p style="margin-bottom:0; margin-right:0; margin-left:10; margin-top:6" 
    align="justify">Top 100 lists for Careers, Appearances and Goals Scored, and individual 
    playing records 
    for every first-team player in the club's professional 
    history, with an ever- growing stock of player photos.</p>
    
    <%
    
    Do Until photofound = "True"
    	Randomize
    	player_no = right("000" & int(rnd*1000)+1,3)
    	photoname = "gosdb/photos/players/" & player_no & ".jpg" 
    	photofound = objFSO.FileExists(Server.MapPath(photoname))

    	
    	if photofound = "True" then
    		sql = "select player_id_spell1, initials, forename, surname, min(first_game_year) as first_game_year, max(last_game_year) as last_game_year "
			sql = sql & "from player "
			sql = sql & "where player_id = " & player_no
			sql = sql & "group by player_id_spell1, initials, forename, surname " 
			rs.open sql,conn,1,2
			if rs.Fields("first_game_year") >= 1960 then
				wenttoyear = rs.Fields("last_game_year")
				if wenttoyear = "9999" then wenttoyear = "present"
				photocaption = "At Random: <a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & "&scp=1,2,3,4,5,6,7"">" 
				if IsNull(rs.Fields("forename")) then
					photocaption = photocaption & rs.Fields("initials") & " " & trim(rs.Fields("surname"))
	  			  else
	  				photocaption = photocaption & rs.Fields("forename") & " " & trim(rs.Fields("surname"))
	  			end if
	  			photocaption = photocaption &  " " & rs.Fields("first_game_year") & "-" & wenttoyear & "</a>"			  
			  else
			  	photofound = "False"
			end if
			rs.close
		end if	
    Loop

    conn.close
    
    response.write("<p style=""margin:6 0 4 0""><img border=""0"" width=""251px"" height=""351px"" src=""" & photoname & """></p>")
    response.write("<p style=""margin:4 0 6 0"">" & photocaption & "</p>")
    
    %>
    
    </td>
    </tr>
    <tr>
    <td width="980" valign="top" style="text-align: center" colspan="3">
    <p style="text-align: justify; margin-left: 110; margin-right: 110; margin-top: 18; margin-bottom: 12">
    <b>Minor Cups and Trophies:</b> the Football League War Cup [1940]; the Full 
    Members Cup [1986] (a competition for tiers 1 and 2, also known as the Simod 
    Cup [1987-88] and Zenith Data Systems Cup [1989-91]); the Football League 
    Trophy (a generic name for a competition for tiers 3 and 4, including the 
    Associate Members Cup [1984], the Freight Rover Trophy [1985-86], the Autoglass Trophy [1993], the Auto Windshields Shield [1994-2000], the LDV Vans Trophy [2000-03] 
    and the Johnstone's Paint Trophy [from 2010]); and official pre-season competitions (the Watney Cup 
    [1973], the Anglo Scottish Cup [1977-79] and the Football League Group Cup 
    [1981]).</td>
    
    </tr>
    </table>
    
<!--#include file="base_code.htm"--></body></html>