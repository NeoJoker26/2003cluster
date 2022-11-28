<%@ Language=VBScript %>
<% Option Explicit %>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
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

<% Dim output, code, codepart, error, latestdate, eventdate, initials, millisecs, validated, radiovalue1, radiovalue2, i %>

<p style="text-align: center; margin: 36 0 36;">
<font color="#47784D"><span style="font-size: 18px"><b>Material for What's New</b></span></font></p>

<%
output = ""
error = 0
code = Request.Form("code")

if code = "" then

	output = output & "<div style=""width:300; margin:0 auto;"">"
	output = output & "<form action=""photocaptions2.asp"" method=""post"" name=""form1"">"
	output = output & "<p style=""margin:0 auto;""><input type=""text"" name=""code"" size=""30""></p>"
	output = output & "<p style=""margin:48 auto 300;""><input type=""submit"" name=""b1"" value=""Continue"" style=""width: 100; font-size: 12px; margin-left:0; margin-right:0; padding:0;""></p>"
	output = output & "</form>"
	output = output & "</div>"

  else
  
  	Dim conn, sql, rs
	Set conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")

	%><!--#include file="conn_read.inc"--><%
  
  	codepart = split(code,":")
  	if ubound(codepart) = 2 then
		eventdate = codepart(0)
		initials = codepart(1)
		millisecs = codepart(2) 	

			output = output & "<form action=""photocaptions2_action.asp"" method=""post"" name=""form2"">"
		
			validated = "N"
			output = output & "<table border=""1"" style=""border-collapse: collapse; margin: 18 auto;"" width=""400px"">"
		
			sql = "select event_published, material_details1, material_details2, datepart(ms,publish_timestamp) as millisecs, convert(varchar,publish_timestamp,113) as timestamp, material_seq "
			sql = sql & "from event_control " 
			sql = sql & "where event_date = '" & eventdate & "' "
			sql = sql & "  and publish_by = '" & initials & "' "
			sql = sql & "order by publish_timestamp, material_seq "

			rs.open sql,conn,1,2
			
			i = 1
	
			Do While Not rs.EOF

				if CInt(rs.Fields("millisecs")) = CInt(millisecs) then validated = "Y"
							
				output = output & "<tr>"
				output = output & "<td>" & rs.Fields("timestamp") & "</td>"
				output = output & "<td>" & rs.Fields("material_details1") & " " & rs.Fields("material_details2") & "</td>"
				
				radiovalue1 = "checked"
				radiovalue2 = ""
				if rs.Fields("event_published") = "Y" then
					radiovalue1 = ""
					radiovalue2 = "checked"
				end if
				
				output = output & "<td><input type=""radio"" name=""pub" & i & """ value=""N"" " & radiovalue1 & ">Off</td>"
				output = output & "<td><input type=""radio"" name=""pub" & i & """ value=""Y"" " & radiovalue2 & ">On</td>"
				output = output & "<input type=""hidden"" name=""ts" & i & """ value=""" & rs.Fields("timestamp") & """>"
				output = output & "<input type=""hidden"" name=""ms" & i & """ value=" & rs.Fields("material_seq") & ">"
				i = i + 1
				output = output & "</tr>" 
				
				rs.MoveNext
		
			Loop
			rs.close
			
			output = output & "</table>"
			
			output = output & "<input type=""hidden"" name=""linecount"" value=" & i-1 & ">"
			
			output = output & "<p style=""margin:18 auto 300;""><input type=""submit"" name=""b2"" value=""Change What's New"" style=""width: 200; font-size: 12px; margin-left:0; margin-right:0; padding:0;""></p>"
			output = output & "</form>"
					
			if validated = "N" then error = 2
			
		
	  else error = 1
	end if 	  	 

end if  

select case error
	case 1
		response.write("<p>Incorrect format for the code</p>")
	case 2
		response.write("<p>No valid material found from you - check the code " & code & "</p> ")
	case else
		response.write(output)
end select

%>

<!--#include file="base_code.htm"-->
</center>
</body>

</html>