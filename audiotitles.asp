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

<% Dim fs,f,Folder, file, output, code, codepart, error, latestdate, opposition, eventdate, initials, author, audiopath, virtaudiopath, i, j, n, buffer %>

<p style="text-align: center; margin: 36 0 36;">
<font color="#47784D"><span style="font-size: 18px"><b>Audio Titles</b></span></font></p>

<%
output = ""
error = 0
code = Request.Form("code")

if code = "" then

	output = output & "<div style=""width:300; margin:0 auto;"">"
	output = output & "<form action=""audiotitles.asp"" method=""post"" name=""form1"">"
	output = output & "<p style=""margin:0 auto;"">Audio code: <input type=""text"" name=""code"" size=""14""></p>"
	output = output & "<p style=""margin:48 auto 300;""><input type=""submit"" name=""b1"" value=""Add titles"" style=""width: 100; font-size: 12px; margin-left:0; margin-right:0; padding:0;""></p>"
	output = output & "</form>"
	output = output & "</div>"

  else
  
  	Dim conn, sql, rs
	Set conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")

	%><!--#include file="conn_read.inc"--><%
  

  	codepart = split(code,":")
  	if ubound(codepart) = 1 then
		eventdate = codepart(0)
		initials = codepart(1)
		
		Select Case initials
			Case "AO"
				author = "Andrew Owen"
			Case "KG"
				author = "Keith Greening"
			Case "SD"
				author = "Steve Dean"
			Case Else
				author = "Unknown"
		End Select

		if author <> "Unknown" then

			' Check for latest match
			sql = "select date, opposition "
			sql = sql & "from match " 
			sql = sql & "where date = (select max(date) from match) "
		
			rs.open sql,conn,1,2
				latestdate = rs.Fields("date") 
				opposition = rs.Fields("opposition")
			rs.close
		
			if latestdate = eventdate then

				output = output & "<form action=""audiotitles_action.asp"" method=""post"" name=""form2"">"
			
				output = output & "<table border=""0"" style=""border-collapse: collapse; margin: 18 auto;"" width=""400px"">"
		
				virtaudiopath = "soundfiles/" & eventdate 
				audiopath = Server.MapPath(virtaudiopath)
	
				Set fs=Server.CreateObject("Scripting.FileSystemObject")
			
				If fs.FolderExists(audiopath) = true then
	
					Set Folder = fs.GetFolder(audiopath)
   					redim filenames(Folder.files.count-1)
   					n = 0 
   		
   					for each file in Folder.files
   						filenames(n) = file.name
       					n = n + 1
					next
   		 
		 			'sort into filename order
	 				for i = 0 to Folder.files.count-1 
   		   
   						for j = (i + 1) to Folder.files.count-1 
       						if strComp(filenames(i),filenames(j),0) = 1 then 
	    	  		    		buffer = filenames(j) 
    	    	  				filenames(j) = filenames(i) 
               					filenames(i) = buffer
          					end if 
        				next
    				next
    	   
					for i = 0 to Folder.files.count-1 

						output = output & "<tr>" 
						output = output & "<td><p style=""color: #808080; margin: 0 0 6 4"">" & filenames(i) & "</p>" 
						output = output & "<input type=""hidden"" name=""filename" & i+1 & """ value=""" & filenames(i) & """></td>"
						output = output & "<td><input type=""text"" name=""caption" & i+1 & """ size=""50""></td>"
						output = output & "</tr>"

					next

					output = output & "</table>"
					output = output & "<input type=""hidden"" name=""code"" value=""" & code & """>"
					output = output & "<input type=""hidden"" name=""filecount"" value=""" & Folder.files.count & """>"
					output = output & "<p style=""margin:18 auto 24;""><input type=""submit"" name=""b2"" value=""Add Titles"" style=""width: 100; font-size: 12px; margin-left:0; margin-right:0; padding:0;""></p>"
					output = output & "</form>"
			  
			  	  else error = 4
				end if		
	
 		  	  else error = 3
			end if
			
 		  else error = 2
		end if	
			
	  else error = 1
	end if 	  	 

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