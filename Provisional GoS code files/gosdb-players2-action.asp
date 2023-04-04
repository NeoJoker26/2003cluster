<%@ Language=VBScript %>
<% Option Explicit %>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Greens on Screen</title>

<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
.heading {font-size: 12px; font-weight: bold}
p {font-size: 12px; text-align: left; line-height:1.4;}

-->
</style>
  
</head>
  
<body><!--#include file="top_code.htm"-->

<% Dim output, error, parm, modind, key, playerid, playername, forename, mod97num, contributor, location, oldtext, text, textparts, textpart, action, urlparm, moderator, modmessage, buttontext, rejected %>

<%
output = ""
error = 0
parm = Request.QueryString("parm")
modind = Request.QueryString("action")
key = left(parm,32)
playerid = mid(parm,33,4)
mod97num = right(parm,10)

if (left(mod97num,8) - right(mod97num,2)) Mod 97 = 0 then

  	Dim conn, sql, rs
	Set conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")

	%><!--#include file="conn_update.inc"--><%
	
	sql = "select forename, surname "
	sql = sql & "from player " 
	sql = sql & "where player_id = " & playerid

	rs.open sql,conn,1,2
	forename = rtrim(rs.Fields("forename"))
	playername = rtrim(rs.Fields("forename")) & " " & rtrim(rs.Fields("surname"))
	rs.close
	
	moderator = rtrim(Request.Form("moderator"))
	modmessage = Request.Form("modmessage")
	contributor = rtrim(Request.Form("contributor"))
	location = rtrim(Request.Form("location"))
	text = rtrim(Request.Form("text"))
	textparts = split(text,Chr(13)&Chr(10))
	rejected = Ucase(Request.Form("rejected"))
	
	text = ""
	for each textpart in textparts
		if textpart > "" then text = text & textpart & "|p|"
	next

	if right(text,3) = "|p|" then text = left(text,len(text)-3)		'drop last paragraph marker
		
	if contributor = "" then
	
		sql = "select contributor, location, text "
		sql = sql & "from player_contribution " 
		sql = sql & "where uniqueid = '" & key & "'"

		rs.open sql,conn,1,2
		if rs.RecordCount = 1 then
			contributor = rtrim(rs.Fields("contributor"))
			location = rtrim(rs.Fields("location"))
			text = rtrim(rs.Fields("text"))
			text = replace(text,"|p|",Chr(13)&Chr(10))		'new paragraph indicator back to CR+LF 
		end if
		rs.close


%>

		<div style="margin:24px auto; width:840px; text-align:left">
		<% 
		urlparm = "parm=" & parm
		if modind = "mod" then urlparm = urlparm & "&action=mod"		
		%>
		<form method="post" action="gosdb-players2-action.asp?<%response.write(urlparm)%>">
		<p style="margin: 6px 0 0"><span class="heading">Thanks for your request to write about <%response.write(forename)%>.</span> You have agreed that you will respect 
		the spirit of Greens on Screen and that you will not abuse this facility. In particular, your contribution will 
   		be relevant, truthful and accurate, will not be insulting, obscene or illegal, and does not reveal matters that are private and personal.</p>
    	<p style="margin: 12px 0 0" class="heading">Your name</p> 
    	<p style="margin: 4px 0">Please add your real name, in the form: John Smith, J. Smith or John S. Other forms will be rejected when moderated. If you 
        contribute to more than one player, please use the same name.</p>
    	<input type="text" name="contributor" value="<%response.write(contributor)%>" size="25">
    	<p style="margin: 12px 0 0" class="heading">Where you live</p>
    	<p style="margin: 4px 0">Your location, based on these examples: 
        'Nottingham' ... 'Newbridge, near Penzance' ... 'Kidbrooke, London' ... 
        'Melbourne, Australia'</p>
    	<input type="text" name="location" value="<%response.write(location)%>" size="25">
    	<p style="margin: 12px 0 0" class="heading">Your text</p>
        <p style="margin: 4px 0 0">
        Write your contribution here, or paste from your preferred word 
        processing software. New paragraphs should be indicated by a skip to a 
        new line (there is no need to add a blank line by skipping twice).</p>
        <p style="margin: 4px 0; font-size: 11px;">
        Note that if you include HTML tags, they will be ignored. To indicated bold, 
        italic or underlined words, surround the characters as in the following 
        examples: [b]these words will be bold[/b] ... [i]these words will be in 
        italics[/i] ... [u]these words will be underlined[/u] ... [in]these 
        words will form an indented new paragraph[/in]</p>
        <textarea style="margin: 6px 0" rows="20" name="text" cols="100"><%response.write(text)%></textarea>
        <% if modind = "mod" then %>
       		<p style="margin: 4px 0 0">Moderator: <input type="text" name="moderator" size="10" style="vertical-align:top">
       		<span style="margin-left:15px">Rejected?</span> <input type="text" name="rejected" size="1" value="N" style="vertical-align:top"></p>
       		<p style="margin: 4px 0 0">Mod Message: <textarea rows="3" name="modmessage" cols="88" style="vertical-align:top"></textarea>	
       		</p>
       		<% buttontext="Moderate this entry" %>
       	<% else %>
        	<p style="margin: 6px 0 0">Your words will be added to Greens on Screen's 
        	player page immediately, so please take a moment to check your text, including 
        	its spelling and grammar.</p>
        	<p style="margin: 4px 0 0">Note that we reserve the right to edit as we feel necessary and to reject a contribution when 
        	it's considered inappropriate.</p>
        	<% buttontext="Add or amend your contribution" %>
        <% end if %>	
       	<p style="margin: 15px 0; text-align:center"><input style="padding:3px 0;" type="submit" value="<%response.write(buttontext)%>"></p>
    	</form>
		</div>

<%				
	
	  else 
	  	
	  	Dim RegEx
		Set RegEx = New RegExp
		RegEx.Pattern = "<[^>]*>"		'detect html
		RegEx.Global = True
	  
	  	text = RegEx.Replace(text,"")	'remove any submitted html
	  	
	  	contributor = replace(contributor,"'","''")		'double apostrophe for sql insert or replace
	  	location = replace(location,"'","''")			'double apostrophe for sql insert or replace
	  	text = replace(text,"'","''")					'double apostrophe for sql insert or replace
	  	text = replace(text,Chr(13)&Chr(10),"|p|")		'new paragraph indicator  
	  		  	
	  	'Check if a player_contribution row already exists for this set: if so, update; if not, insert

		sql = "select text "
		sql = sql & "from player_contribution " 
		sql = sql & "where uniqueid = '" & key & "' "

		rs.open sql,conn,1,2
		if rs.RecordCount > 0 then oldtext = rs.Fields("text")
		rs.close

		if oldtext > "" then

			sql = "update player_contribution set "
			sql = sql & "contributor = '" & contributor & "',"	
			sql = sql & "location = '" & location & "',"
			sql = sql & "text = '" & text & "'"
			if modind = "mod" then sql = sql & ", rejected = '" & rejected & "' " 
			sql = sql & " where uniqueid = '" & key & "' "

			on error resume next
			conn.Execute sql
			if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
			On Error GoTo 0
	
  		  else

			sql = "insert into player_contribution (uniqueid, player_id, contributor, location, datetime_added, text) "
			sql = sql & "values ("
			sql = sql & "'" & key & "',"
			sql = sql & playerid & ","
			sql = sql & "'" & contributor & "',"
			sql = sql & "'" & location & "',"
			sql = sql & "getdate(),"
			sql = sql & "'" & text & "'"
			sql = sql & ")"	
			
			on error resume next
			conn.Execute sql
			if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
			On Error GoTo 0

		end if
		
		Dim strTo,strFrom,strCc,strBcc,message,subject
	   								
		strTo = "gos_mod@emaildodo.com"
		strFrom = "player_contribution@greensonscreen.co.uk"
		strCc = "steve@greensonscreen.co.uk"
		
		if oldtext = "" then
			action = "added"
		 elseif modind = "mod" then
		 	if rejected = "Y" then
		 		action = "ENTRY REJECTED by " & moderator
		 	  elseif replace(oldtext,"'","''") = text then 
		 		action = "moderated (no changes needed) by " & moderator
		 	  else
		 		action = "moderated (changes made) by " & moderator
		 	end if
		 else
		 	action = "amended"		
		end if
		
		subject = "GoS Entry " & action & " for " & playername & ", written by " & contributor

		message = "Text has been " & action & " for " & playername & ", written by " & contributor & " from " & location & ":" 
		if modmessage > "" then message = message & "<br><br><span style=""color:blue"">Moderator Message: " & modmessage & "</span>"		 
		if oldtext > "" then message = message & "<br><br>Old Text:<br><br>" & oldtext & "<br><br>New Text:"		 
		message = message & "<br><br>" & replace(text,"''","'")	'undo the double apostrophe for the email
		message = message & "<br><br>Link to player's page: <a href=""http://www.greensonscreen.co.uk/gosdb-players2.asp?pid=" & playerid & """>http://www.greensonscreen.co.uk/gosdb-players2.asp?pid=" & playerid & "</a>" 
		message = message & "<br><br>Link to moderate (be careful!): <a href=""http://www.greensonscreen.co.uk/gosdb-players2-action.asp?parm=" & parm & "&action=mod"">http://www.greensonscreen.co.uk/gosdb-players2-action.asp?parm=" & parm & "&action=mod</a>"
	   				
		%><!--#include file="emailcode.asp"--><%		
%>
		<div style="margin:36px auto; width:420px; height:300px">
		<p style="font-size:12px; text-align: justify; line-height:1.5; margin:12px 0;">Thank you for your contribution to Greens on Screen. Your words can 
        be viewed <a href="gosdb-players2.asp?pid=<%response.write(playerid)%>"><u>here</u></a>, and will be included in the home page's What's New 
        section as soon as it has been checked and passed, or within 12 hours, 
        whichever comes first, assuming that it hasn't been rejected. </p>
		<p style="font-size:12px; text-align: justify; line-height:1.5; margin:12px 0;">If you want to change any of the content, use the link provided in the email
		or if that's not possible, get in touch using 'Contact Us' (top-right).</p>
		</div>
<%	  
	end if 
		  	 
  else error = 1

end if  

select case error
	case 1
		response.write("<p>Incorrect format for the code</p>")
	case 2
		response.write("<p>Unknown author</p>")
	case 3
		response.write("<p>The date does not match the most recent fixture</p>")
	case 4
		response.write("<p>No valid material found from you - check the code</p>")
	case else
		response.write(output)
end select

%>

<!--#include file="base_code.htm"-->
</body>

</html>