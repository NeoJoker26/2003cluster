<%@ Language=VBScript %> 
<% Option Explicit %>

<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>Greens on Screen Database</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--

#message {width:500px; margin: 0px auto; text-align:center; border: 1px solid #808080; background-color: #e7f1ec;}
#usermsg {width:600px; margin: 12px auto; text-align:left; padding: 6px; border: 1px solid #808080; background-color: #e7f1ec;}

#lengthmessage {text-align:left; border: 0px none; background-color: #ffffff; font-family: verdana,arial,helvetica,sans-serif; font-size: 11px; color: green}

#ajaxplayerlist {margin: 0 0 20px; width: 240px;} 
#ajaxplayerlist p {margin: 0; padding: 0;  font-size: 11px; font-weight:normal; text-align: left; color:#202020;}

textarea {border: 1px solid #808080;}
.currentpenpic:focus {outline: none !important;}

textarea {font-family: verdana,arial,helvetica,sans-serif; font-size: 11px; line-height:1.3; padding:4px}
td p {font-family: verdana,arial,helvetica,sans-serif; font-size: 11px; margin: 0 0 6px}

-->
</style>

<script language="javascript">

//window.location.reload(true); // force page refresh 

function GetPlayerList(initial,username) { 

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
// triggered1() function handles the events 
xmlhttp.onreadystatechange = triggered1; 

// open takes in the HTTP method and url.
document.body.style.cursor='wait';   
var url="gosdb-getplayerspageplayerlist.asp";     
url=url+"?initial="+initial;
url=url+"&camefrom=gosdb-playerupdate.asp";
url=url+"&username="+username;
url=url+"&sid="+Math.random();

xmlhttp.open("GET", url, true);
document.getElementById('ajaxplayerlist').innerHTML = '<img style="margin: 0 0 0 12;" border="0" src="images/ajax-loader.gif"><br><img border="0"" src="images/dummbar_0.gif" height="20px" width="1">' 
 
// send the request. if this is a POST request we would have 
// sent post variables: send("name=aleem&gender=male) 
// Moz is fine with just send(); but 
// IE expects a value here, hence we do send(null); 
xmlhttp.send(null);
document.body.style.cursor='auto';  
} 
 
function triggered1() { 
// if the readyState code is 4 (Completed) 
// and http status is 200 (OK) we go ahead and get the responseText 
// other readyState codes: 
// 0=Uninitialised 1=Loading 2=Loaded 3=Interactive 
if (xmlhttp.readyState == 4) { 
        // xmlhttp.responseText object contains the response.
		 document.getElementById('ajaxplayerlist').innerHTML = xmlhttp.responseText ;
} 
} 

function Validate() {
    var radios = document.getElementsByName('phase')

    for (var i = 0; i < radios.length; i++) {
        if (radios[i].checked) {
        return true; // checked
    }
    };
    // not checked, show error
    document.getElementById('ValidationError').innerHTML = '<p class="style1boldred">You must select an option before submitting</p>';
    return false;
}

function countChar(txtBox,messageDiv)
    {
        try
        {
           count = txtBox.value.length;
            if (count < 3500)
            {
                txt = "Current length: " + count + " characters";
                messageDiv.innerHTML=txt 
                document.getElementById('actionbox').style.display="block";
            }
            else if (count < 4095)
            {
                txt = "WARNING: pen-pictures have a limit of 4095 characters (including spaces). This one now has " + count + " characters.";
                messageDiv.innerHTML='<font style="color:blue;">' + txt + '</font>';
                document.getElementById('actionbox').style.display="block";
            }

            else
            {
                txt = "ERROR: This pen-picure is too large and cannot be stored. The current length is " + count + ". Please reduce it to less that 4096.";
                messageDiv.innerHTML='<font style="color:red;font-weight:bold;">' + txt + '</font>';
                document.getElementById('actionbox').style.display="none";
            }
            }
            catch ( e )
            {
            }
    }  

</script>

</head>

<body>
<!--#include file="top_code.htm"-->

<div style="margin: 0 auto">
<p class="style1boldgreen" style="margin: 12px auto 6px; font-size: 15px; font-style: Arial; ">PLAYER PROFILE EDIT</p>


<%
dim conn, sql, rs, phase, playerid, forename, surname, penpic, penpic_pending, penpic_pending_approval, penpic_pending_author, penpic_pending_approver, penpic_pending_notes
dim outline, playername, username, favplayer, contributor_type, approver, fromapprovelist, linktext, boxwidth, work
dim x, y, message, paras_old, paras_new, para_old, para_new, paras_old_array(30,20), paras_new_array(30,20)
dim sentences_old, sentences_new, sentence_old, sentence_new

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

playerid = Request.QueryString("pid")
fromapprovelist = Request.QueryString("approvelist")

phase = Request.Form("phase")
if phase = "" then 
	phase = session("phase")
	session("phase") = ""
end if
	
Select Case phase 

	Case ""
		
		username = Request.Cookies("volunteer")("username")
		favplayer = Request.Cookies("volunteer")("favplayer")
		approver = Request.Cookies("volunteer")("approver")
	
		if username = "" then		
			%>
			<div style="margin: 12px 0; width: 200px; text-align: left; ">
			<%
			if Request.QueryString("signin") = "unknown" then response.write("<p class=""style1boldred"">Unregistered details - try again</p>") 
			%>
			<form style="margin: 0 0 24px" method=post action="gosdb-playerupdate.asp">  
  			<p style="font-size: 11px; margin:9px 0 6px;">Your name:</p>
  			<input type="text" style=""margin: 0" name="username" size="25">
  			<p style="font-size: 11px; margin:9px 0 6px;">Favourite player (surname only):</p>
			<input type="text" style=""margin: 0" name="favplayer" size="25">
			<p class="style1bold"Note that if you allow cookies on your device, your details will be remembered for 30 days after each use.</p>
  			<input type="hidden" name="phase" value="signin"> 
			<input style="font-size: 12px; margin: 0 auto; " type=submit value="Sign in">
  			</form>
  			</div>
			<%
	  	  else 'previously authenticated
	  	  
	  	  	Response.Cookies ("volunteer").Expires = Date + 30		'extend volunteer cookie
	  
	  		if playerid > "" then 
	  			Call Edit_Display
	  		  else 
	  			Call Selection_Display 
	  		end if
	  	
	  	end if
	  	
	Case "signin"
	
		username = Request.Form("username")
		favplayer = Request.Form("favplayer")
		
		sql = "select contributor_type, name "
		sql = sql & "from contributor "
		sql = sql & "where lower(name) = '" & LCase(username) & "' "
		sql = sql & "  and rtrim(lower(favourite)) = '" & LCase(favplayer) & "' "
	
		rs.open sql,conn,1,2
		if rs.RecordCount > 0  then	  
			contributor_type = rs.Fields("contributor_type")
			username = rs.Fields("name")					'save user name again to ensure exactly as registered
		  else 
		  	response.redirect "gosdb-playerupdate.asp?signin=unknown"
		end if
		rs.close
		
		if isnull(contributor_type) or instr(contributor_type,"P") = 0 then
			response.write("<p class=""style1bold"" style=""margin:18px 0 96px;"">Necessary authorisation for player updates not found</p>")
		  else
		  	response.Cookies ("volunteer")("username") = username
		  	response.Cookies ("volunteer")("favplayer") = favplayer
		  	if instr(contributor_type,"Z") > 0 then response.Cookies ("volunteer")("approver") = "Y"

			Response.Cookies ("volunteer").Expires = Date + 30
			
			approver = request.Cookies ("volunteer")("approver")
			
			Call Email
		  	Call Selection_Display
		end if
		
	Case "signout"
	
		username = Request.Cookies("volunteer")("username")
		Call Email
	
		Response.Cookies ("volunteer").Expires = Date - 1
		response.redirect "gosdb-playerupdate.asp" 
		
	Case "confirm"
	
		playerid = session("playerid")
		playername = session("playername")
		username = session("username")
		%>
		<div style="margin: 12px 0; width: 500px; text-align: left; ">
		<form style="margin: 0 0 24px" method=post action="gosdb-playerupdate.asp">  
  		<p style="font-size: 11px; margin:9px 0 6px;">Please confirm that you are ready to pass 
        the profile for <%response.write playername%> to the sign-off stage. If 
        you think it might be useful, please add a message for the approver 
        here. For example, you might want to explain an aspect of your work, or perhaps you've uncovered something in your 
        researches that would help GoS-DB in the wider sense, e.g. a missing 
        date of birth or middle name.</p>
  		<textarea style="margin:0 0 12px; width:500px; height:100px;" name="notes" rows="1" cols="20"></textarea>
 		<input type="hidden" name="phase" value="signal_approve">
 		<input type="hidden" name="playerid" value="<%response.write(playerid)%>">
  		<input type="hidden" name="playername" value="<%response.write(playername)%>">
  		<input type="hidden" name="username" value="<%response.write(username)%>">
		<input style="font-size: 12px; margin: 0 auto; " type=submit value="Confirm ready for sign-off"> or 
        <a href="gosdb-playerupdate.asp">RETURN TO RECONSIDER</a>
  		</form>
  		</div>
  	<%
	Case "signal_approve"
	
		session("username") = username
		session("playerid") = playerid
		session("playername") = playername
		
		Server.Execute("gosdb-playerupdate_action.asp") 			'signal that this player is ready for approval
 
 		response.redirect "gosdb-playerupdate.asp"

End Select
%>		  

</div>	
<!--#include file="base_code.htm"-->

</body>

</html>

<% Sub Selection_Display
%>
	
	<p class="style1boldgreen" style="margin: 0; font-size:14px"><% response.write(username) %></p>
	
	<form method=post action="gosdb-playerupdate.asp">
	<input type="hidden" name="phase" value="signout">
	<p class="style1">Public PC?</span>  	
	<input style="display: inline; margin:0 0 18px 6px;font-size: 11px; margin: 0 auto;" type=submit value="Sign out">
	</form></p>
        
    <% 
	message = session("message")
	if message > "" then
		response.write("<div id=""message"">" & message & "</div>")
		session("message") = ""
	end if 
	%>    
	
    <table style="margin: 18px 0;" border="0" cellspacing="0" style="border-collapse: collapse" cellpadding="0" width="960">
    <tr>
    
    <td width="240" valign="top">
	    
	<p style="margin:0 18px 9px 0; "> 
    <b>To find a player ...</b><br>type the first few letters of the surname, then select a name revealed below.</p>

	<formxxxx style="padding:0; margin:0;" name='form1'>
	<p style="margin-left: 0; margin-top:0; margin-bottom:6px">
	<input type='text' name='surname' onKeyUp="GetPlayerList(this.value,'<% response.write(username) %>')" size="12" ></p>
	</form>
	
    <center>
    <div id="ajaxplayerlist" style="padding-top:6px; padding-right: 18px;">
    
    <p><img border="0" src="images/dummbar_0.gif" height="20" width="1"></p>
    
    </div>   
	</td>
	  
	<td width="240" valign="top">
	<p style="margin: 0; " class="style1bold">Suggestions ... <br>
    <span style="font-weight: 400">Short profile, long careers</span></p>
	<p style="margin: 0 0 9px" class="style3">
    <span style="font-weight: 400">(years here are first spells only) </span></p>
	<%
	sql = "select top 20 a.player_id, surname, b.initials, first_game_year, "
	sql = sql & "case last_game_year when '9999' then '' else '-' + cast(right(last_game_year,2) as varchar) end as lastgameyear, "
	sql = sql & "c.name, count(*), 1000*count(*)/ len(penpic) as ratio "
	sql = sql & "from match_player a join player b on a.player_id = b.player_id left outer join contributor c on b.penpic_pending_author = c.name "
	sql = sql & "where spell = 1 "
	sql = sql & "  and (penpic_defer_until < GETDATE() or penpic_defer_until is null) "
	sql = sql & "  and penpic_pending_date is null "
	sql = sql & "group by a.player_id, surname, b.initials, first_game_year, last_game_year, c.name, len(penpic) "
	sql = sql & "order by ratio desc "
	
	rs.open sql,conn,1,2	  

	Do While Not rs.EOF
		if trim(rs.Fields("name")) = username or isnull(rs.Fields("name")) then
			linktext = "<a href=""gosdb-playerupdate.asp?pid=" & rs.Fields("player_id") & """>" & trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials")) & "</a>"
		  else
		  	linktext = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if	  	
		response.write("<p style=""margin: 1px 0"">" & linktext & "&nbsp;&nbsp;" & rs.Fields("first_game_year") & rs.Fields("lastgameyear") & "</p>")
		rs.MoveNext	
	Loop
	rs.close
	
	%>
    </td>
    
    <td width="240" valign="top">
	<p class="style1bold" style="margin: 0; text-align: left">Suggestions 
    ... <br>
    <span style="font-weight: 400">Players in the last 10 years</span></p>
	<p class="style3" style="margin: 0 0 9px; text-align: left;">
    <span style="font-weight: 400">(years here are first spells only)</span></p>
	<%
	sql = "select top 20 newid() as random, player_id_spell1, surname, a.initials, first_game_year, "
	sql = sql & "case last_game_year when '9999' then '' else '-' + cast(right(last_game_year,2) as varchar) end as lastgameyear, b.name "
	sql = sql & "from player a left outer join contributor b on a.penpic_pending_author = b.name "
	sql = sql & "where spell = 1 "
	sql = sql & "  and (penpic_defer_until < GETDATE() or penpic_defer_until is null) "
	sql = sql & "  and penpic_pending_date is null "
	sql = sql & "  and exists (select * from player c "
	sql = sql & "  				where c.player_id_spell1 = a.player_id_spell1 "
	sql = sql & "  				and c.last_game_year between year(getdate())-10 and year(getdate()) "
	sql = sql & "  			  ) "
	sql = sql & "order by random "
	
	rs.open sql,conn,1,2	  

	Do While Not rs.EOF
		if trim(rs.Fields("name")) = username or isnull(rs.Fields("name")) then
			linktext = "<a href=""gosdb-playerupdate.asp?pid=" & rs.Fields("player_id_spell1") & """>" & trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials")) & "</a>"
		  else
		  	linktext = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
		end if	  	
		response.write("<p style=""margin: 1px 0"">" & linktext & "&nbsp;&nbsp;" & rs.Fields("first_game_year") & rs.Fields("lastgameyear") & "</p>")
		rs.MoveNext	
	Loop
	rs.close
	
	%>
    </td>

    <td width="240" valign="top">
    
    <p class="style1bold" style="margin: 0 0 6px; text-align: left;">Under construction ...</p>
	<%
	sql = "select surname, player_id, a.initials, first_game_year, "
	sql = sql & "case last_game_year when '9999' then '' else '-' + cast(right(last_game_year,2) as varchar) end as lastgameyear, shortname, b.name "
	sql = sql & "from player a join contributor b on penpic_pending_author = b.name "
	sql = sql & "where a.penpic_pending_author is not null "
	sql = sql & "and penpic_pending_approval is null "
	sql = sql & "order by penpic_pending_date "
	
	rs.open sql,conn,1,2	  

	if rs.RecordCount > 0  then	 
		Do While Not rs.EOF
			if trim(rs.Fields("name")) = username then
				linktext = "<a href=""gosdb-playerupdate.asp?pid=" & rs.Fields("player_id") & """>" & trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials")) & "</a>"
		  	  else
		  		linktext = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
			end if	  	
			response.write("<p style=""margin: 0"">" & trim(rs.Fields("shortname")) & ": " & linktext & "&nbsp;&nbsp;" & rs.Fields("first_game_year") & rs.Fields("lastgameyear") & "</p>")
			rs.MoveNext	
		Loop
	  else
	  	response.write("<p style=""margin: 0"">None</p>")
	end if	  	
	rs.close
	
	%>
    
    <p class="style1bold" style="margin: 9px 0 6px; text-align: left;">Ready for sign-off ...</p>
    <%
	sql = "select surname, player_id, a.initials, first_game_year, "
	sql = sql & "case last_game_year when '9999' then '' else '-' + cast(right(last_game_year,2) as varchar) end as lastgameyear, "
	sql = sql & "shortname, penpic_pending_approver "
	sql = sql & "from player a join contributor b on penpic_pending_author = b.name "
	sql = sql & "where a.penpic_pending_author is not null "
	sql = sql & "and penpic_pending_approval = 'Y' "
	sql = sql & "order by penpic_pending_date "
	
	rs.open sql,conn,1,2
	
	if rs.RecordCount > 0  then	  
		Do While Not rs.EOF
			if approver = "Y" then
				linktext = "<a href=""gosdb-playerupdate.asp?pid=" & rs.Fields("player_id") & "&approvelist=Y"">" & trim(rs.Fields("surname")) & "</a>, " 
				linktext = linktext & "<a href=""gosdb-players2.asp?pid="  & rs.Fields("player_id") & "&status=preview"" target=""_blank"">" & trim(rs.Fields("initials")) & "</a>"
		  	  else
		  		linktext = trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials"))
			end if	
			linktext = linktext & "&nbsp;&nbsp;" & rs.Fields("first_game_year") & rs.Fields("lastgameyear")   	
			if not isnull(rs.Fields("penpic_pending_approver")) then 
		  		work = split(rs.Fields("penpic_pending_approver"))
		  		linktext = linktext & " - " & left(work(0),1) & left(work(1),1) 
		  	end if
			response.write("<p style=""margin: 0"">" & trim(rs.Fields("shortname")) & ": " & linktext & "</p>")
			rs.MoveNext	
		Loop
	  else
	  	response.write("<p style=""margin: 0"">None</p>")
	end if	  	
	rs.close
	
	%>
    </td>
    
    </tr>
 	 	
    <tr>
    
    <td valign="top" colspan="4">
	<p style="margin: 18px 0; line-height: 1.5; text-align: left;"><span class="style1bold">Recently completed ... </span>
    <%
	sql = "select surname, a.player_id, initials, first_game_year, "
	sql = sql & "case last_game_year when '9999' then '' else '-' + cast(right(last_game_year,2) as varchar) end as lastgameyear, "
	sql = sql & "penpic_pending_author, archive_timestamp, " 
	sql = sql & "cast(day(archive_timestamp) as varchar) + '/' + cast(month(archive_timestamp) as varchar) as date_archived "
	sql = sql & "from player a join player_penpic_archive b on a.player_id = b.player_id "
	sql = sql & "where archive_timestamp >= DATEADD(month,-1,GETDATE()) "
	sql = sql & "  and archive_timestamp = (select max(archive_timestamp) from player_penpic_archive c where c.player_id = b.player_id "
	sql = sql & "    and c.archive_reason = 'PP') "
	sql = sql & "order by archive_timestamp desc "
	
	rs.open sql,conn,1,2	  

	if rs.RecordCount > 0  then
		Do While Not rs.EOF
			response.write("<span style=""white-space: nowrap;"">" & rs.Fields("date_archived") & ": <a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id") & """>" & trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials")) & " " & rs.Fields("first_game_year") & rs.Fields("lastgameyear") & "</a> ...</span> ")
			rs.MoveNext	
		Loop
	  else
	  	response.write("None")
	end if
	rs.close
	
	%>
	    
	</p></td>
	
	</tr>
 	 	
    <tr>
	  
	<td valign="top" colspan="4">
	<p class="style1bold" style="margin-top: 12px">IMPORTANT NOTES</p>
	<p class="style1" align="justify">1. You can either search for a specific player or 
    choose from two suggestion lists, the first being those players who have had 
    significant careers with us but whose current profile is short, and the 
    second is a list of players who have left in the last 10 years. Note that 
    the players in the first list will probably require serious research, so it 
    might not be a good place to start. The second list is a random selection of 
    leavers, which will change every time you refresh the page.</p>
    <p class="style1" align="justify">2. Every name here is a link to the next step, unless 
    someone is already working on him, in which case the name will appear in 
    black rather than green and will also be shown in the 'Under construction' 
    list in the right-hand column.</p>
    <p class="style1" align="justify">3. Very important: when you click on a name, the player 
    will be allocated to you and made unavailable to anyone else. If you decide 
    not to proceed with him, you must remember to select and submit 'Cancel and 
    release' on the next page.</td>
    
    </tr>
 	 	
	</table>
<%
End Sub

Sub Edit_Display

	sql = "select player_id_spell1, surname, forename, penpic, penpic_pending, penpic_pending_approval, penpic_pending_author, penpic_pending_approver, penpic_pending_notes "
	sql = sql & "from player  "
	sql = sql & "where player_id = " & playerid
	
	rs.open sql,conn,1,2	  
		
	if not IsNull(rs.Fields("forename")) then forename = trim(rs.Fields("forename"))
	if not IsNull(rs.Fields("surname")) then surname = trim(rs.Fields("surname"))
	
	playername = forename & " " & surname
	
	if not IsNull(rs.Fields("penpic")) then 
		penpic = replace(rs.Fields("penpic"),"|p|","&#13;&#10;&#13;&#10;")
	  else
	  	penpic = "None"
	end if

	if not IsNull(rs.Fields("penpic_pending")) then penpic_pending = replace(rs.Fields("penpic_pending"),"|p|","&#13;&#10;&#13;&#10;")
	
	penpic_pending_approval = rs.Fields("penpic_pending_approval")
	penpic_pending_author = rs.Fields("penpic_pending_author")
	penpic_pending_approver = rs.Fields("penpic_pending_approver")
	penpic_pending_notes = rs.Fields("penpic_pending_notes")

	rs.close
	
	boxwidth = "485px"
	
	'reserve or approve_reserve this player for this username (session variables are the only way to pass data for these phase options)
	if penpic_pending_approval = "Y" and approver = "Y" and fromapprovelist = "Y" then
		phase = "approve_reserve"
		session("phase") = phase
		boxwidth = "400px"
	  else
		session("phase") = "reserve"
	end if	
	
	session("username") = username
	session("playerid") = playerid
	session("playername") = playername
		
	Server.Execute("gosdb-playerupdate_action.asp") 			'reserve this player
	
	%>
	
	<p class="style1boldgreen" style="margin: 6px auto 0; font-size: 15px;"><%response.write(playername)%></p>
	
	<%
	if penpic_pending_notes > "" then 
		response.write("<div id=""usermsg""><p class=""style1"" style=""margin: 0 0 6px""><b>A message from " & penpic_pending_author & ": </b></p>")
		response.write("<p class=""style1"" style=""margin: 0"">" & penpic_pending_notes & "</p></div>")
	end if
	%>  
	
	<form method=post action="gosdb-playerupdate_action.asp">

  	<table border="0" cellspacing="2" style="border-collapse: collapse" cellpadding="0" width="980">
    <tr>
    
	<td valign="top">
	<p style="margin: 0 0 9px" class="style1bold">Current Version</p>
	<textarea class="currentpenpic" style="background-color:f0f0f0; width:<%response.write(boxwidth)%>; height:600px;" name="penpic1" readonly rows="1" cols="20"><%response.write(penpic)%></textarea>
	</td>
    
    
    <% 
    if phase = "approve_reserve" then 
    	response.write("<td valign=""top"" style=""width:170px; padding: 3px 0 6px 12px"">")
    	outline = "<p class=""style1bold"" style=""margin: 26px 0 6px;"">Difference Summary</p>"
    	
    	paras_old = split(penpic,"&#13;&#10;")
    	x = 0
    	for each para_old in paras_old
    		if trim(para_old) > "" then 
    			paras_old_array(x,0) = trim(para_old)
    			x = x + 1
    		end if	 
    	next
    	
    	paras_new = split(penpic_pending,"&#13;&#10;")
    	x = 0
    	for each para_new in paras_new
    		if trim(para_new) > "" then 
    			paras_new_array(x,0) = trim(para_new)
    			x = x + 1
    		end if	 
    	next
    	
    	x = 0
    	do until paras_old_array(x,0) = "" and paras_new_array(x,0) = ""
    		if paras_old_array(x,0) = paras_new_array(x,0) then
    			outline = outline & "<p>Para " & x+1 & ": <span class=""style1boldgreen"">Same</span></p>"
    		  else
    		    outline = outline & "<p>Para " & x+1 & ": <span class=""style1boldred"">Different</span></p>"
    		    sentences_old = split(paras_old_array(x,0),". ")
    		    y = 1
    		   	for each sentence_old in sentences_old
    				paras_old_array(x,y) = trim(sentence_old)
    				y = y + 1	 
    			next
    		    sentences_new = split(paras_new_array(x,0),". ")
    		    y = 1
    		   	for each sentence_new in sentences_new
    				paras_new_array(x,y) = trim(sentence_new)
    				y = y + 1	 
    			next 
    			y = 1
    			do until paras_old_array(x,y) = "" and paras_new_array(x,y) = ""
    				if paras_old_array(x,y) = paras_new_array(x,y) then
    					outline = outline & "<p style=""margin-left: 6px"">Sentence " & y & ": <span class=""style1boldgreen"">Same</span></p>"
    		  		  else
    		    		outline = outline & "<p style=""margin-left: 6px"">Sentence " & y & ": <span class=""style1boldred"">Different</span></p>"
		    		end if
		    		y = y + 1
		    	loop	
    		end if
    		x = x + 1
    	loop
    	
    	response.write(outline)
    	response.write("</td>")
    end if
    %>
    
    <td valign="top">
	<p class="style1bold" style="margin: 0 0 9px; text-align: right;">New/Amended Version</p>
	
	<%
	response.write("<textarea name=""penpic2"" style=""width:" & boxwidth & "; height:600px;"" onkeyup=""countChar(penpic2,lengthmessage);"" onkeydown=""if(event.keyCode == 13){document.getElementById('btnSendTextMessage').click();}""")
		
	if penpic_pending_approval = "Y" and penpic_pending_approver > "" and penpic_pending_approver <> username then
		response.write(" readonly class=""currentpenpic"">")		
	  else
		response.write(">")
	end if

	if penpic_pending = "" then 
		response.write(penpic & "</textarea>")
	  else
		response.write(penpic_pending & "</textarea>")
	end if
	
	response.write("<div id=""lengthmessage"" style=""margin: 9px 0; width:" & boxwidth & """></div>")

	%>
	
    </td>

 	</tr>
 	
 	<tr>
 	
 	<td valign="top" style="padding:0 18px 0 6px 0">
 	<%
 	if phase <> "approve_reserve" then
 	%>
 		<p class="style1bold">Notes:</p>
 		<p>1. The current version is for reference only.</p>
 		<p>2. The new or amended version is initially primed with the current 
    	version if there is one. Changes here will be lost unless you choose the 
    	appropriate action.</p>
 		<p>3. Start a new paragraph at an appropriate point every few sentences by 
    	skipping to a new line. Skip twice - i.e. insert a blank line - if it helps you visually, but the number of skips does not matter - the 
    	correct spacing for the new paragraph will automatically be inserted on the player's page, regardless of 
    	the number of skips. </p>
 		<p>4. Amendments to recent players will most likely involve adding sentences 
    	(and perhaps paragraphs) at the end of the existing text, but the way the 
    	old and new blend together might need some adjustment, so feel free to 
    	change the text at any point within the existing profile. </p>
    <% else %>
    	</td><td>
    <% end if %>
 	</td>
 	
 	<td id="actionbox" valign="top"  style="padding:0 0 0 6px">
 	<p class="style1bold">Choose an action:</p>
 	<% if phase <> "approve_reserve" then %>
		<p><input type="radio" value="cancel-release" name="phase">Cancel action, drop changes and release name</p>
		<p><input type="radio" value="review-in6" name="phase">No action required but display again in 6 months time</p>
		<p><input type="radio" value="review-end" name="phase">No action required and football career at an end, so no further review</p>
		<p><input type="radio" value="incomplete" name="phase">Incomplete, save changes, I'll continue later</p>
		<p><input type="radio" value="ready" name="phase">Job done, send for sign-off</p>
		<% if username = "Steve Dean" then %>
			<p><input type="radio" value="approve_fasttrack" name="phase">No approval needed, fast-track sign-off</p>
			<p><input type="radio" value="approve_correct" name="phase">Correction to latest version</p>
		<% end if %>
	<% end if %>
	<% if phase = "approve_reserve" then %>
		<p><input type="radio" value="approve_release" name="phase">Release from my approval</p>
		<p><input type="radio" value="approve_incomplete" name="phase">Incomplete, save changes, I'll finish approving later</p>
		<p><input type="radio" value="approve" name="phase">Approve and promote to current version</p>
	<% end if %>
  	<% response.write("<input type=""hidden"" name=""playerid"" value=""" & playerid & """>") %>
  	<% response.write("<input type=""hidden"" name=""playername"" value=""" & playername & """>") %>
  	<% response.write("<input type=""hidden"" name=""username"" value=""" & username & """>") %>
	<div id="ValidationError" name="ValidationError">
	</div>
    <p style="font-size: 11px; margin:12px 0"><input type=submit value="Submit" onclick="return Validate();" ></p>
    </form>

 	</td>
 	
	</table>
<%
End Sub

Sub Email

		Dim strTo,strFrom,strCc,strBcc,message,subject
	   								
		strTo = "GoSDBprofiles@greensonscreen.co.uk"
		strFrom = "GoSDBprofiles@greensonscreen.co.uk" 
		strBcc = "player_contribution@greensonscreen.co.uk"			
		subject = "GoS-DB Profiles : " & username & " - " & phase
		message = "" 
		
 	    %><!--#include file="emailcode.asp"--><%

End Sub
%>