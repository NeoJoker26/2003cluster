<div id="admin_head" style="text-align:center;">
<h2>GREENS ON SCREEN ADMINISTRATION</h2>
<%
administrator = request.servervariables("remote_USER")
response.write("<h3>Admin ID: " & administrator & "</h3>")
response.write("<div style=""margin:0 auto;"">")
Dim temp1 : temp1 = Split(Request.ServerVariables("SCRIPT_NAME"), "/")
Dim temp2 : temp2 = temp1(UBound(temp1))
if temp2 = "admin.asp" then
	response.write("<a href=""newmatch.asp"" class=""button"">Add New Match</a>")
	response.write("<a href=""milestone_refresh.asp"" class=""button"">Refresh Milestones</a>")
	response.write("<a href=""consecutive_refresh.asp"" class=""button"">Refresh Consecutives</a><br>")
	response.write("<a href=""league_table_update.asp"" class=""button"">Update League Table</a>")
	response.write("<a href=""update_missing_attendances.asp"" class=""button"">Missing League Attendances</a>")
	response.write("<a href=""add_youtube_link.asp"" class=""button"">Add YouTube Links</a><br>")
	response.write("<a href=""admin_table_update.asp"" class=""button"">Update GoS Tables</a>")
	response.write("<a href=""admin_bulk_fixtures.asp"" class=""button"">Table Update</a>")
  else
	response.write("<a href=""admin.asp"" class=""button"">Return to Admin Menu</a>")
end if
response.write("</div>")
%>
</div>