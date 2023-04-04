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
.heading {font-weight: bold}
p {font-size: 12px; text-align: left; line-height:1.5; margin: 0 0 12px; }
ol, ul {font-size: 12px; text-align: left; line-height:1.7;  }
-->
</style>
  
</head>
  
<body><!--#include file="top_code.htm"-->

<% Dim fs,f,Folder, file, output, error, parm, key, playerid, playername, forename, mod97num, emailaddr %>

<%
output = ""
error = 0
parm = Request.QueryString("parm")
key = left(parm,32)
playerid = mid(parm,33,4)
mod97num = right(parm,10)

if (left(mod97num,8) - right(mod97num,2)) Mod 97 = 0 then

  	Dim conn, sql, rs
	Set conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")

	%><!--#include file="conn_read.inc"--><%
	
	sql = "select forename, surname "
	sql = sql & "from player " 
	sql = sql & "where player_id = " & playerid

	rs.open sql,conn,1,2
	forename = rtrim(rs.Fields("forename"))
	playername = rtrim(rs.Fields("forename")) & " " & rtrim(rs.Fields("surname"))
	rs.close

	emailaddr = Request.Form("emailaddr")
	if emailaddr = "" then

%>

		<div style="margin:36px auto; width:640px">
		<p class="heading">How to contribute to the page for <%response.write("<span style=""color: #0f6e3c"">" & playername & "</span>")%></p> 
		<p>From a brief recollection of a recent player to your in-depth researches of a pioneer, from a few sentences to many paragraphs, you can submit your contribution using these simple steps:</p> 
		<ol>
      		<li>This page allows you to write about <%response.write(playername)%>.  If you wish to write about another player, go back to the player pages, find that player and select his 'contribute 
            here' link.</li>
      		<li>Please read the guidelines below, and only proceed if you are 
            prepared to abide by them.</li>
      		<li>Enter your email address. Note that your address will only be used to 
      		trigger the next step, and to allow us to get in touch should the need 
      		arise. It will not be stored or shared for any other purpose. </li>
      		<li>You will be sent an email - click on the link within to add your contribution.</li>
      		<li>Once submitted, refresh the player page to view your work.</li>
      		<li>If you want, use the same email link to make an amendment or add more, 
            but only for this player. If you'd like to write about more than one player, select 
            the 'contribute here' link for each one. </li>
    	</ol>
		<p class=heading>Guidelines</p>
		<p>You contribution will appear on <%response.write(forename)%>'s page as soon as you submit it. 
        We will review it within a few hours to ensure that it meets these 
        guidelines:</p>
		<ul>
      		<li>Please respect the spirit of Greens on Screen by not abusing this facility 
      		and in particular, by not abusing players. If you didn't like a 
            player, this is not the place to say so.&nbsp; </li>
      		<li>Feel free to write whatever you want, as long as it is relevant, 
            truthful and accurate; is not insulting, obscene or 
            illegal; and does not reveal matters that would be commonly 
            considered private or personal.</li>
      		<li>Note that we reserve the right to edit as we feel necessary and 
      		reject contributions at our discretion.&nbsp; </li>
    	</ul>
    	<form method="post" action="gosdb-players2-contribute1.asp?parm=<%response.write(parm)%>">
    	<p>To continue, please enter your email address:
    	<input type="text" name="emailaddr" size="25"> 
      	<input style="padding:2px; margin: 0" type="submit" value="Continue">
    	</form></p>
		</div>

<%				
	
	  else 

		Dim strTo,strFrom,strCc,strBcc,message,subject
	   								
		strTo = emailaddr
		strFrom = "player_contribution@greensonscreen.co.uk"
		strCc = "player_contribution@greensonscreen.co.uk"
		subject = "Your contribution to Greens on Screen"
				 
		message = "Please use the following link to add your contribution for " & playername & "."  
		message = message & "<br><br><a href=""http://www.greensonscreen.co.uk/gosdb-players2-action.asp?parm=" & parm & """>Click here to proceed" & "</a>, "
		message = message & "or copy and paste the following string into your browser:<br>www.greensonscreen.co.uk/gosdb-players2-action.asp?parm=" & parm 
		message = message & "<br><br>Note that you can use this link to amend your words, so it's a good idea to keep this email for a while."
		message = message & "<br><br>Thanks for joining in," 
		message = message & "<br>Steve" 	   			
		
 		%><!--#include file="emailcode.asp"--><%
%>
		<div style="margin:36px auto; width:420px; height:300px">
		<p>Thanks, an email has been sent to <%response.write(emailaddr)%>.<p>To write your contribution, please follow the link that it provides.</p> 
		<p>If you do not receive an email within a few minutes, check to see if it has 
        been misdirected to your junk-mail folder, or try again (you might have miskeyed your email address).</p>
		<p style="text-align: center"><a href="gosdb-players2.asp?pid=<%response.write(playerid)%>">Return to GoS-DB</a></p>
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