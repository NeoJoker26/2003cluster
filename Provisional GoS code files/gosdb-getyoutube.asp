<%
response.expires = -1

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_read.inc"--><%

	' Get match details
	
	sql = "select date, publish_timestamp, opposition, goalsfor, goalsagainst, homeaway, material_details1 "
	sql = sql & "from match join event_control on date = event_date "
	sql = sql & "where event_published = 'Y' and event_type = 'M' and material_type = 'Y' "
	sql = sql & "and datediff(""dd"", publish_timestamp, getdate()) < 14 "
	sql = sql & "order by publish_timestamp desc "
	
	rs.open sql,conn,1,2
	
	if rs.RecordCount > 0 then 
	      	
       	output = "<table style=""border: 1px solid #d0d0d0; padding: 2px 1px;"">"
       	output = output & "<tr><td colspan=""5"" class=""style4boldgreen"" style=""margin-top: 12px;"">Video Clips Found on YouTube<img class=""close"" style=""margin: 0; float:right; border: 0"" src=""images/close.png""></td></tr>"
        output = output & "<tr>"
					
		Do While Not rs.EOF
		
			output = output & "<td style=""width:20px""><span class=""WNtag"">" & DateDiff("d", Now(), rs.Fields("publish_timestamp")) & "</span></td>"
			output = output & "<td>" & right("0" & day(rs.Fields("date")),2) & " " & monthname(month(rs.Fields("date")),True) & " " & year(rs.Fields("date")) & "</td>"
			
			output = output & "<td>"
			if rs.Fields("homeaway") = "H" then
				output = output & "Argyle " & rs.Fields("goalsfor") & "-" & rs.Fields("goalsagainst") & " " & rs.Fields("opposition") & "</td>"
		 	 else 
				output = output & rs.Fields("opposition") & " " & rs.Fields("goalsagainst") & "-" & rs.Fields("goalsfor") & " Argyle" & "</td>"   		  
			end if	  
			
			output = output & "<td style=""white-space: nowrap;""><img class=""video"" src=""images/video16.png"">"
			output = output & "<a href=""https://www.youtube.com/embed/" & rs.Fields("material_details1") 
			output = output & "?rel=0&amp;wmode=transparent"" onclick=""return hs.htmlExpand(this, {objectType: 'iframe'})"" class=""highslide"">"
			output = output & "on YouTube</a></td>"

			output = output & "<td><a href=""gosdb-match.asp?date=" & rs.Fields("date") & """>Match page</a></td>"
		
		    output = output & "</tr>"
			rs.MoveNext
		
		Loop
		
       	output = output & "</table>"		
						     	
      	rs.close
      	
	  else
	
		output = "No match found"
	
	end if	
		
conn.close

response.write(output)

%>