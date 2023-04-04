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
td, p {font-size: 11px;}
-->
</style>
  
</head>
  
<body><!--#include file="top_code.htm"-->

<% Dim output, content, code, codepart, error, latestdate, datetime, eventdate, existingcount  %>

<p style="text-align: center; margin: 36 0 36;">
<font color="#47784D"><span style="font-size: 18px"><b>YouTube Links</b></span></font></p>

<p class="style1" style="margin-bottom:18px"><a target="_blank" href="https://www.youtube.com/user/argylemedia">Copy link parameter from here</a>, then ...</p>

<%
output = ""
error = 0
code = Request.Form("code")

if code = "" then

	output = output & "<div style=""width:300; margin:0 auto;"">"
	output = output & "<form action=""youtubelinks.asp"" method=""post"" name=""form1"">"
	output = output & "<p style=""margin:0 auto;"">YouTube code: <input type=""text"" name=""code"" size=""14""></p>"
	output = output & "<p style=""margin:48 auto 300;""><input type=""submit"" name=""b1"" value=""Continue"" style=""width: 100; font-size: 12px; margin-left:0; margin-right:0; padding:0;""></p>"
	output = output & "</form>"
	output = output & "</div>"

  else
  
  	codepart = split(code,":")
  	if ubound(codepart) = 1 then
		eventdate = codepart(0)
				
		Select Case codepart(1)
			Case 1
				content = "Action Highlights"
			Case 2
				content = "Matchday Moments"
			Case 3
				content = "Found on YouTube"
			Case Else
				content = "Error"
		End Select

		if content <> "Error" then
		
		  	Dim conn, sql, rs
			Set conn = Server.CreateObject("ADODB.Connection")
			%><!--#include file="conn_read.inc"--><%
			Set rs = Server.CreateObject("ADODB.Recordset")

			' Find latest match and also get current date/time to use as timestamp
			sql = "select date, convert(varchar,getdate(),120) as datetime "
			sql = sql & "from match " 
			sql = sql & "where date = (select max(date) from match) "
		
			rs.open sql,conn,1,2
				latestdate = rs.Fields("date") 
				datetime = rs.Fields("datetime")
			rs.close
			
			' Look for existing YouTube content for this date and type
			sql = "select count(*) as linkcount "
			sql = sql & "from event_control " 
			sql = sql & "where event_date = '" & eventdate & "' "
			sql = sql & "  and material_details2 = '" & content & "' "
		
			rs.open sql,conn,1,2
				existingcount = rs.Fields("linkcount") 
			rs.close

			if latestdate <> eventdate then output = output & "<p style=""font-size:14px; color:red; margin: 0 auto 10px"">WARNING: The specified date is not for the latest match</p>" 

			output = output & "<div style=""text-align:left; width:400px"">"
			output = output & "<form action=""youtubelinks_action.asp"" method=""post"" name=""form2"">"
			output = output & "<p>Match date: " & eventdate & "</p>"
			output = output & "<input type=""hidden"" name=""eventdate"" value=""" & eventdate & """>"
			output = output & "<p>Content type: " & content & "</p>"
			output = output & "<input type=""hidden"" name=""content"" value=""" & content & """>"			
			output = output & "<p>Update date/time: <input type=""text"" name=""datetime"" value=""" & datetime & """ size=""18""></p>"
			output = output & "<p>YouTube code: <input type=""text"" name=""youtubecode"" size=""12""  maxlength=""12""></p>"
			if existingcount > 0 then 
				output = output & "<p style=""color:red;"">" & existingcount
				if existingcount = 1 then  
					output = output & " link exists"
				  else
				  	output = output & " links exist"
				end if 
				output = output & " for this date and type, and will be removed unless unchecked here: " 
				output = output & "<input type=""checkbox"" name=""delete"" value=""Y"" checked></p>"
			end if
			output = output & "<p style=""margin:18 auto 24;""><input type=""submit"" name=""b2"" value=""Add YouTube link"" style=""width: 100; font-size: 12px; margin-left:0; margin-right:0; padding:0;""></p>"
			output = output & "</form>"
			output = output & "</div>"
			
 		   else error = 2
 		   
		end if
				
	  else error = 1
	end if 	  	 

end if  

select case error
	case 1
		response.write("<p>Incorrect format for the code</p>")
	case 2
		response.write("<p>Unknown category</p>")
	case else
		response.write(output)
end select

%>

<!--#include file="base_code.htm"-->
</body>

</html>