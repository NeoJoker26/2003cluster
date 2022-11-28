<%@ Language=VBScript %>
<% Option Explicit %>
<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">

<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
</head>
<body>
<%
Dim penpic1, penpic2, playerid, phase, playername, username, recaffected, message, text, textpart, textparts, emailbody, notes
playername = Request.Form("playername")
if session("playername") > "" then 								'used in the case of phase = reserve (and a few others)
	playername = session("playername")		
	session("playername") = ""
end if

'penpic1 is only accessed to provide information in the notification email 
penpic1 = Request.Form("penpic1")
penpic2 = replace(Request.Form("penpic2"),"'","''")				'convert to double apostrophe for SQL string
penpic2 = replace(penpic2,"£","&pound;")						'convert £ to &pound; for SQL string

phase = Request.Form("phase")
if session("phase") > "" then 									'used in the case of phase = reserve (and a few others)
	phase = session("phase")		
	session("phase") = ""
end if

username = Request.Form("username")
if session("username") > "" then 								'used in the case of phase = reserve (and a few others)
	username = session("username")		
	session("username") = ""
end if

playerid = Request.Form("playerid")
if session("playerid") > "" then 								'used in the case of phase = reserve (and a few others)
	playerid = session("playerid")		
	session("playerid") = ""
end if

message = ""

Dim conn, sql, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_update.inc"--><%

	
Select Case phase

	Case "reserve"

		sql = "update player set " 
		sql = sql & "penpic_pending_author = '" & username & "', "
		sql = sql & "penpic_pending_date = getdate() "
		sql = sql & "where player_id = '" & playerid & "' "
		sql = sql & "  and spell = 1 "
		sql = sql & "  and penpic_pending_author is null " 
		
		on error resume next
		conn.Execute sql,recaffected
		if err <> 0 then 
			message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
		  else
		  	message = "<p class=""style1"">The profile for " & playername & " has been allocated to you</p>"
		end if
		On Error GoTo 0	
		
		Call Respond
		
	
	Case "cancel-release"

		sql = "update player set " 
		sql = sql & "penpic_pending_author = NULL, "
		sql = sql & "penpic_pending_date = NULL, "
		sql = sql & "penpic_pending = NULL, "
		sql = sql & "penpic_pending_notes = NULL "	
		sql = sql & "where player_id = '" & playerid & "' "
		sql = sql & "  and penpic_pending_author = '" & username & "' "
		
		on error resume next
		conn.Execute sql,recaffected
		if err <> 0 then 
			message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
		  elseif recaffected <> 1 then 
			message = "<p class=""style1boldred"">Cancel failed: this is not your player. Check with Steve.</p>"
		  else
		  	message = "<p class=""style1"">" & playername & " has been released and any incomplete changes have been removed</p>"
		end if
		On Error GoTo 0	
		  	
		Call Respond	
	
	Case "review-in6"

		sql = "update player set " 
		sql = sql & "penpic_pending_author = NULL, "
		sql = sql & "penpic_pending_date = NULL, "
		sql = sql & "penpic_pending = NULL, "
		sql = sql & "penpic_defer_until = DATEADD(month,6,GETDATE()) "
		sql = sql & "where player_id = '" & playerid & "' "
		sql = sql & "  and penpic_pending_author = '" & username & "' "
		
		on error resume next
		conn.Execute sql,recaffected
		if err <> 0 then 
			message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
		  elseif recaffected <> 1 then 
			message = "<p class=""style1boldred"">Update failed: this is not your player. Check with Steve.</p>"
		end if
		On Error GoTo 0		
				
		if message = "" then			'Update successful, so insert to archive
			sql = "insert into player_penpic_archive (player_id, archive_reason, author, archive_penpic) "
			sql = sql & "values ("
			sql = sql & playerid & ","
			sql = sql & "'6',"
			sql = sql & "'" & username & "',"
			sql = sql & "NULL"
			sql = sql & ")"	
			
			on error resume next
			conn.Execute sql
			if err <> 0 then 
				message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
			  else
			  	message = "<p class=""style1"">" & playername & " will not appear on the suggestion lists for the next six months</p>"
			end if
			On Error GoTo 0
		end if
		
		Call Respond	
	
	Case "review-end"

		sql = "update player set " 
		sql = sql & "penpic_pending_author = NULL, "
		sql = sql & "penpic_pending_date = NULL, "
		sql = sql & "penpic_pending = NULL, "
		sql = sql & "penpic_defer_until = '9999-12-31' "	
		sql = sql & "where player_id = '" & playerid & "' "
		sql = sql & "  and penpic_pending_author = '" & username & "' "
		
		on error resume next
		conn.Execute sql,recaffected
		if err <> 0 then 
			message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
		  elseif recaffected <> 1 then 
			message = "<p class=""style1boldred"">Update failed: this is not your player. Check with Steve.</p>"
		end if
		On Error GoTo 0		
				
		if message = "" then			'Update successful, so insert to archive
			sql = "insert into player_penpic_archive (player_id, archive_reason, author, archive_penpic) "
			sql = sql & "values ("
			sql = sql & playerid & ","
			sql = sql & "'99',"
			sql = sql & "'" & username & "',"
			sql = sql & "NULL"
			sql = sql & ")"	
			
			on error resume next
			conn.Execute sql
			if err <> 0 then 
				message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
			  else
			  	message = "<p class=""style1"">" & playername & " will no appear on the suggestion lists</p>"	
			end if  			
			On Error GoTo 0
			  	
		end if
				  	
		Call Respond
			
		
	Case "incomplete","approve_incomplete","ready"
	
		textparts = split(penpic2,Chr(13)&Chr(10))
		penpic2 = ""
	
		for each textpart in textparts
			if trim(textpart) > "" then penpic2 = penpic2 & trim(textpart) & "|p|"
		next
		
		if right(penpic2,3) = "|p|" then penpic2 = left(penpic2,len(penpic2)-3)		'remove final paragraph marker

		sql = "update player set " 
		sql = sql & "penpic_pending = '" & penpic2 & "' "
		sql = sql & "where player_id = '" & playerid & "' "
		if phase = "approve_incomplete" then
			sql = sql & "  and penpic_pending_approver = '" & username & "' "
		  else
			sql = sql & "  and penpic_pending_author = '" & username & "' "
		end if
		
		on error resume next
		conn.Execute sql,recaffected
		if err <> 0 then 
			message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
		  elseif recaffected <> 1 then 
			message = "<p class=""style1boldred"">Update failed: this is not your player. Check with Steve.</p>"
		  else
		  	message = "<p class=""style1"">Any changes for " & playername & " have been saved for you to work on later</p>"
		end if
		On Error GoTo 0	
		  	
		Call Respond	
		
		
	Case "approve_reserve"

		sql = "update player set " 
		sql = sql & "penpic_pending_approver = '" & username & "' "
		sql = sql & "where player_id = '" & playerid & "' "
		sql = sql & "  and spell = 1 "
		sql = sql & "  and penpic_pending_approval = 'Y' " 
		sql = sql & "  and penpic_pending_approver is null "
		
		on error resume next
		conn.Execute sql,recaffected
		if err <> 0 then 
			message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
		  else
		  	message = "<p class=""style1"">The profile for " & playername & " has been allocated for you to sign off</p>"
		end if
		On Error GoTo 0	
		
		Call Respond


	Case "approve_release"

		sql = "update player set " 
		sql = sql & "penpic_pending_approver = NULL "
		sql = sql & "where player_id = '" & playerid & "' "
		sql = sql & "  and penpic_pending_approver = '" & username & "' "
		
		on error resume next
		conn.Execute sql,recaffected
		if err <> 0 then 
			message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
		  elseif recaffected <> 1 then 
			message = "<p class=""style1boldred"">Cancel failed: this is not your player. Check with Steve.</p>"
		  else
		  	message = "<p class=""style1"">" & playername & " has been released from your approval</p>"
		end if
		On Error GoTo 0	
		  	
		Call Respond
	
	
	Case "signal_approve"
	
		notes = replace(Request.Form("notes"),"'","''")				'convert to double apostrophe for SQL string
	
		sql = "update player set " 
		sql = sql & "penpic_pending_approval = 'Y', "
		sql = sql & "penpic_pending_notes = '" & trim(notes) & "' "
		sql = sql & "where player_id = '" & playerid & "' "
		sql = sql & "  and penpic_pending_author = '" & username & "' "
		
		on error resume next
		conn.Execute sql,recaffected
		if err <> 0 then 
			message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
		  elseif recaffected <> 1 then 
			message = "<p class=""style1boldred"">" & playerid & username & recaffected & "Update failed: this is not your player. Check with Steve.</p>"
		  else
		  	message = "<p class=""style1"">Thanks - your work for " & playername & " is waiting for sign-off</p>"
		end if
		On Error GoTo 0	
		  	
		Call Respond	
		
		
	Case "approve","approve_fasttrack","approve_correct"

	
		sql = "select penpic_version, penpic_pending_author, penpic "
		sql = sql & "from player  "
		sql = sql & "where player_id = " & playerid 
		if phase = "approve" then sql = sql & "  and penpic_pending_approver = '" & username & "' " 	'not appropriate for 'approve_fasttrack' or 'approve_correct'

		rs.open sql,conn,1,2	  

		if rs.RecordCount > 0  then	 

			sql = "insert into player_penpic_archive (player_id, archive_reason, author, archive_penpic) "
			sql = sql & "values ("
			sql = sql & playerid & ","
			if phase = "approve_fasttrack" then
				sql = sql & "'PX',"
			  else
				sql = sql & "'PP',"
			end if
			sql = sql & "'" & rs.Fields("penpic_pending_author") & "',"
			if not IsNull(rs.Fields("penpic")) then 
				sql = sql & "'" & replace(rs.Fields("penpic"),"'","''") & "'"	'convert any apostrophes to double apostrophe for SQL string
			  else
			  	sql = sql & "NULL"
			end if  	
			sql = sql & ")"	
			
			on error resume next
			conn.Execute sql
			if err <> 0 then message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
			On Error GoTo 0
	
			if message = ""	then 	
				textparts = split(penpic2,Chr(13)&Chr(10))
				penpic2 = ""
	
				for each textpart in textparts
					if trim(textpart) > "" then penpic2 = penpic2 & trim(textpart) & "|p|"
				next
		
				if right(penpic2,3) = "|p|" then penpic2 = left(penpic2,len(penpic2)-3)		'remove final paragraph marker

				sql = "update player set " 
				sql = sql & "penpic = '" & penpic2 & "', "
				if phase = "approve" or phase = "approve_fasttrack" then					'do not update penpic date or version for a correction
					sql = sql & "penpic_date = getdate(), "
					if isnull(rs.Fields("penpic_version"))  then
						sql = sql & "penpic_version = 1, "
				  	  else
						sql = sql & "penpic_version = penpic_version + 1, "
					end if
				end if					
				sql = sql & "penpic_defer_until = DATEADD(month,6,GETDATE()), "
				sql = sql & "penpic_pending_author = NULL, "
				sql = sql & "penpic_pending_date = NULL, "
				sql = sql & "penpic_pending_approval = NULL, "
				sql = sql & "penpic_pending_approver = NULL, "
				sql = sql & "penpic_pending = NULL, "
				sql = sql & "penpic_pending_notes = NULL " 
				sql = sql & "where player_id = '" & playerid & "' "

				on error resume next
				conn.Execute sql,recaffected
				if err <> 0 then 
					message = "<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>"
				  else
				  	message = "<p class=""style1"">The profile for " & playername & " has been updated successfully</p>"
				end if
				On Error GoTo 0
			end if
			
	  	  else
	  	  
	  		message = "<p class=""style1boldred"">Approval failed: this is not your player. Check with Steve.</p>"
	  		phase = phase & " failed"	'change for the email subject
		
		end if	  	
		
		rs.close
		
		Call Respond			

End Select
	   
%>
</body>
</html>
<% 
Sub Respond

session("message") = message

if (phase = "reserve" or phase = "approve_reserve") and recaffected = 0 then

  else
	
	emailbody = "From " & username & "<br><br>"
	emailbody = emailbody & "<a href=""http://www.greensonscreen.co.uk/gosdb-playerupdate.asp"">Go to GoS Profile Update page</a><br><br>"
	
	if penpic2 > ""	then	'i.e. for the phases when penpics are passed through
		if StrComp(penpic1,replace(penpic2,"''","'")) = 0 then 
			emailbody = emailbody & "No changes were made to the text"
  	  	  else
  			emailbody = emailbody & "Old Pen-pic:<br><br>" & penpic1 & "<br><br>Proposed Pen-pic:<br><br>" & replace(penpic2,"''","'")
		end if
	end if
	
	Dim strTo,strFrom,strCc,strBcc,subject
	   								
	strFrom = "GoSDBprofiles@greensonscreen.co.uk"
	strCc = ""
	strBcc = ""

	if phase = "signal_approve" then
		strTo = "gosprofilesignoff@emaildodo.com"
		subject = "GoS-DB Profiles - " & playername & " ready for sign-off"
		emailbody = emailbody & "Notes from author: " & notes
	  else
	  	strTo = "GoSDBprofiles@greensonscreen.co.uk"
	  	subject = "GoS-DB Profiles - " & username & " - " & playername & " (" & playerid & ") - " & phase
	end if
	message = emailbody 
	
 	%><!--#include file="emailcode.asp"--><%

end if

select case phase

	case "reserve","approve_reserve","signal_approval"		
		'These option come from a server.execute, so no need to redirect (server.execute returns anyway)
	
	case "ready"
		session("playerid") = playerid
		session("playername") = playername
		session("username") = username	
		session("phase") = "confirm"	 
		Response.Redirect "gosdb-playerupdate.asp"
	
	case else 
		Response.Redirect "gosdb-playerupdate.asp"

end select

End Sub
		 
%>