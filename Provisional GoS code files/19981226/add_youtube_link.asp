<%@ Language=VBScript %> 
<% Option Explicit %>
<!DOCTYPE html>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GoS Admin</title>
<link rel="stylesheet" type="text/css" href="../gos2.css">
<style>
<!--
#container {
	font-size:11px; 
	text-align:left; 
	width:fit-content; 
	margin:24px auto;
	}
-->
</style>
</head>

<body>

<% 
Dim output, phase, administrator
Dim match_date, video_type, old_code, new_code, timestamp, result_message
Dim conn,sql,rs

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_admin.inc"--><%
%>

<div id="container">
<!--#include file="admin_head.inc"-->

<h3 style="margin:6px 0 15px;">ADD YOUTUBE LINKS</h3>

<% 

phase = request.form("phase")
if request.form("admin") > "" then administrator = request.form("admin")

select case phase
	case 1
		Call Ask_for_code
	case 2
		Call Add_code
	case else
		Call Choose_type
end select

Response.write(output)

%>
</body>
</html>


<%
Sub Choose_type

	output = "<p class=""style1bold"" style=""margin-bottom:18px"">1. <a target=""_blank"" href=""https://www.youtube.com/user/argylemedia"">Copy link parameter from here</a></p>"

	output = output & "<form action=""add_youtube_link.asp"" method=""post"">"
	output = output & "<input type=""hidden"" name=""phase"" value=""1"">"
	output = output & "<input type=""hidden"" name=""admin"" value=""" & administrator & """>"	
	output = output & "<p class=""style1bold"">2. Select video type:</p>"
	output = output & "<input type=""radio"" name=""video"" id=""action"" value=""action"">"
	output = output & "<label for=""action"">Action Highlights</label><br>"
	output = output & "<input type=""radio"" name=""video"" id=""moments"" value=""moments"">"
	output = output & "<label for=""moments"">Matchday Moments</label><br>"
	output = output & "<input type=""radio"" name=""video"" id=""found"" value=""found"">"
	output = output & "<label for=""found"">Found on YouTube</label><br>"

	'Get latest match date
	
	sql = "select max(date) as maxdate " 
	sql = sql & "from match "
	rs.open sql,conn,1,2
					
	match_date = rs.Fields("maxdate")

	rs.close

	output = output & "<p class=""style1bold"">3. Check match date:</p>"
	output = output & "<input type=""text"" name=""matchdate"" size=""6"" value=""" & match_date & """>"		
	output = output & "<br><input style=""margin:20px 5px 0;"" type=""submit"" value=""Continue"">"
	output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
	output = output & "</form>"

End Sub

Sub Ask_for_code

	if request.form("video") = "" then 
		output = output & "<p class=""style1boldred"">Select a video type</p>"
		Call Backbutton
	  else
		match_date = request.form("matchdate")
	
		select case request.form("video")
			case "action"		
				video_type = "Action Highlights"
			case "moments"		
				video_type = "Matchday Moments"
			case "found"		
				video_type = "Found on YouTube"
		end select			
	
		output = "<form action=""add_youtube_link.asp"" method=""post"">"
		output = output & "<input type=""hidden"" name=""phase"" value=""2"">"
		output = output & "<input type=""hidden"" name=""admin"" value=""" & administrator & """>"	
		output = output & "<input type=""hidden"" name=""matchdate"" value=""" & match_date & """>"		
		output = output & "<input type=""hidden"" name=""video"" value=""" & video_type & """>"	
	
		'Get YouTube code if it already exists
		
		sql = "select material_details1 " 
		sql = sql & "from event_control "
		sql = sql & "where event_date = '" & match_date & "' "
		sql = sql & "  and event_type = 'M' "		 		 		'indicates a match event
		sql = sql & "  and material_type = 'Y' "		 	 		'indicates a YouTube link
		sql = sql & "  and material_details2 = '" & video_type & "' "		 	 		
		rs.open sql,conn,1,2							
			if rs.RecordCount > 0 then old_code = rs.Fields("material_details1")	
		rs.close
	
		output = output & video_type & ":<input style=""margin-left:10px;"" type=""text"" name=""newcode"" size=""8"""
		if old_code > "" then output = output & " value=""" & old_code & """"
		output = output & ">"	
		output = output & "<input type=""hidden"" name=""oldcode"" value=""" & old_code & """>"
		output = output & "<br><input style=""margin:20px 5px 0;"" type=""submit"" value=""Add or correct the code"">"
		output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
		output = output & "</form>"
		if old_code > "" then  
			output = output & "<p><span class=""style1boldred"">WARNING:</span> A code already exists for this date and video type.<br>"
			output = output & "If you remove this code without a new one, the existing YouTube link will be removed.<br>"
			output = output & "If you replace it with a new code, the old link will be removed and a new one added.</p>"
		end if
	end if

End Sub

Sub Add_code

	match_date = request.form("matchdate")
	video_type = request.form("video")
	old_code = request.form("oldcode")
	old_code= rtrim(old_code)
	new_code = request.form("newcode")
	new_code = trim(new_code)
	
	if old_code = new_code then
		output = output & "<p class=""style1boldblue"">The old and new codes are the same. No action taken.</p>"
	  else	
		sql = "select convert(varchar,getdate(),120) as datetime "
		rs.open sql,conn,1,2							
		timestamp = rs.Fields("datetime")	
		rs.close
	
		if old_code > "" then
		
			'a row exists for this date and video type, so delete it before adding the new one
	
			sql = "delete from event_control "
			sql = sql & "where event_date = '" & match_date & "' "
			sql = sql & "  and material_details1 = '" & old_code & "' "			
			on error resume next
			conn.Execute sql
			if err <> 0 then output = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
			On Error GoTo 0	
		
		end if
		
		'Now insert the new row for this video type 
		
		sql = "set dateformat ymd; "
		sql = sql & "insert into event_control (event_date, event_published, event_type, material_type, material_seq, publish_timestamp, updateno, material_details1, material_details2)"
		sql = sql & "values ("
		sql = sql & "'" & match_date & "',"
		sql = sql & "'Y',"
		sql = sql & "'M',"
		sql = sql & "'Y',"
		sql = sql & "1,"
		sql = sql & "'" & timestamp & "',"
		sql = sql & "99,"
		sql = sql & "'" & new_code & "',"
		sql = sql & "'" & video_type & "'"
		sql = sql & ")"	
			
		on error resume next
		conn.Execute sql
		if err <> 0 then 
			result_message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
			output = output & result_message
		  else
			result_message = "<p class=""style1boldgreen"">" & video_type & " code successfully added</p>"
			output = output & result_message	  
		end if
		On Error GoTo 0

		Dim strTo,strFrom,strCc,strBcc,subject,message
	   								
		strTo = "youtube_added@greensonscreen.co.uk"
		strFrom = "youtube_added@greensonscreen.co.uk"
		strCc = ""
		subject = "GoS YouTube link added"
		message = result_message & " for " & match_date
		%><!--#include virtual="/emailcode.asp"--><% 
	end if
	
End Sub 

Sub Backbutton
	output = output & "<form>"
	output = output & "<input type=""button"" value=""Back"" onclick=""history.back()"">"
	output = output & "</form"
End sub

%>