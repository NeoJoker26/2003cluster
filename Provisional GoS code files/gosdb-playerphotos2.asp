<%@ Language=VBScript %> 
<% Option Explicit %>

<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Author" content="Trevor Scallan">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<title>Greens on Screen Database</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
#table1 td {border: 1px solid #c0c0c0; text-align:left; margin: 0; padding-left:4; padding-right:4; padding-top:3; padding-bottom:3}
-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Server.ScriptTimeout = 180 
Dim fs, photofolder, conn,sql,rs, i, n, yesphoto, nophoto, dupphoto, outline, startyear, startdate, enddate, lastyears, photoname, temp1, temp2, temp3

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%
		
Set fs=Server.CreateObject("Scripting.FileSystemObject")
Set photofolder = fs.GetFolder(Server.MapPath("/gosdb/photos/players/small"))

startyear = Request.QueryString("year")
if startyear = "" then startyear = 1903

if startyear = "all" then
	startdate = "1903-07-01"
	enddate = year(now) & "-06-30"
  else
  	startdate = startyear & "-07-01"
	enddate = startyear+10 & "-06-30"
end if
%>
  
  <center>
  <table border="0" cellspacing="0" style="border-collapse: collapse" 
  cellpadding="0" width="980">
    <tr>
    <td width="260" valign="top" style="text-align:center;">

	<p style="text-align: center; margin-top:0; margin-bottom:3">
	<a href="gosdb.asp"><font color="#404040"><img border="0" src="images/gosdb-small.jpg" align="left"></font></a><font color="#404040"> 
	<b><font style="font-size: 15px">Search by<br>
	</font></b><span style="font-size: 15px"><b>Player</b></span></font><p style="text-align: center; margin-top:0; margin-bottom:6">
	<b>
	<a href="gosdb.asp">Back to<br>GoS-DB Hub</a></b></p>

	</td>
    
  	<td width="460" align="center" style="text-align: center" valign="top">	
	<p style="margin-top:12; margin-bottom:0; text-align:center; font-size:18px; color:#006E32">
    PLAYER PHOTOS</p>  
    
	<p style="margin-top:6; margin-bottom:0; text-align:center; font-size:13px">
    &nbsp;</p> 
       
    </td>
        
	<td width="260" valign="top"  align="justify">
	<font color="#FF0000">This page is for website development purposes. It is 
    not intended for public viewing and should not be considered accurate.</font></h3>
    </td>
    </tr>   
	</table>
     

	<%
	yesphoto = ""
	nophoto = ""
	dupphoto = ""
	outline = "<table id=""table1"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"">"
	sql = ""
			sql = sql & "select years, player_id_spell1, surname, forename, initials,startpos "
    		sql = sql & "from season a join match_player b on date between date_start and date_end "
			sql = sql & " join player c on b.player_id = c.player_id "
'			sql = sql & "where date between '" & startdate & " ' and ' " & enddate & " ' "
			sql = sql & "where date in ('1947-12-27','1948-1-10') "
'			sql = sql & "  and date = (select min(date) from match_player d where d.player_id = c.player_id_spell1) "
			sql = sql & "order by years, startpos "
'			sql = sql & "order by years, surname, initials "

	rs.open sql,conn,1,2
	
	n = 0
	
	Do While Not rs.EOF
	
		if rs.Fields("years") <> lastyears then
			outline = outline & "<tr><td  bgcolor=""#E0F0E0"" colspan=""5""><p style=""text-align: center; font-size:16px; margin-top:4; margin-bottom:4""><b>" & rs.Fields("years") & "<b></p></td></tr>" 
			lastyears = rs.Fields("years") 
			Dim missinglist
			missinglist = missinglist & "<br>" & rs.Fields("years") & "<br>"
			n = 0
		end if
		

		if n = 5 then
			outline = outline & "</tr>" 
			n = 0
		end if
		
		if n = 0 then outline = outline & "<tr>"
		
		outline = outline & "<td width=""120"" align=""center"" style=""text-align: center"" valign=""bottom"">" 

		if len(rs.Fields("player_id_spell1")) < 4 then photoname = right("00" & rs.Fields("player_id_spell1"),3) & ".jpg"
		
		if instr(yesphoto,left(photoname,3)) =  0 then
			if instr(nophoto,left(photoname,3)) > 0 then
				photoname = "nophoto.jpg"
			  else
				if (fs.FileExists(photofolder & "/" & photoname)) <> true then
					missinglist = missinglist & rtrim(rs.Fields("surname")) & ", " & rtrim(rs.Fields("forename"))  & " " & photoname & "<br>"
					nophoto = nophoto & left(photoname,3) & ","
					photoname = "nophoto.jpg"
			  	  else
					yesphoto = yesphoto & left(photoname,3) & ","
				end if
			end if
		end if
		
		outline = outline & "<img border=""0"" src=""gosdb/photos/players/small/" & photoname & """>" 
		outline = outline & "<p style=""margin:0 0 8 0""><b>"

		if IsNull(rs.Fields("forename")) then
			outline = outline & rs.Fields("startpos") & " " & rs.Fields("initials") & " " & trim(rs.Fields("surname"))
	  		else
	  		outline = outline & rs.Fields("startpos") & " " & rs.Fields("forename") & " " & trim(rs.Fields("surname"))
		end if
		outline = outline & "</b></p></td>"
		n = n + 1
		
		for i = 1 to 0
				
 			if instr(dupphoto,left(photoname,3) & "_" & i) >  0 then 
 			  elseif (fs.FileExists(photofolder & "/" & left(photoname,3) & "_" & i & ".jpg")) = true then	
 			  	dupphoto = dupphoto & left(photoname,3) & "_" & i & ","
 			  else exit for
 			end if
 			
 			if n = 5 then
				outline = outline & "</tr>" 
				n = 0
			end if
		
			if n = 0 then outline = outline & "<tr>"
		
			outline = outline & "<td width=""120"" align=""center"" style=""text-align: center"" valign=""bottom"">" 
 			
 			outline = outline & "<img border=""0"" src=""gosdb/photos/players/small/" & left(photoname,3) & "_" & i & ".jpg"">" 
			outline = outline & "<p style=""margin:0 0 8 0""><b>"
				
			if IsNull(rs.Fields("forename")) then
				outline = outline & rs.Fields("initials") & " " & trim(rs.Fields("surname")) & " (" & i+1 & ")"
	  			else
	  			outline = outline & rs.Fields("forename") & " " & trim(rs.Fields("surname")) & " (" & i+1 & ")"
			end if
			
			outline = outline & "</b></p></td>"
			
			n = n + 1
		
		next

  		rs.MoveNext
	Loop
		
	rs.close
	
	outline = outline & "</tr>"
			
	temp1 = split(yesphoto,",")
	temp2 = split(nophoto,",")
	temp3 = split(dupphoto,",")
	response.write("<center><p style=""margin:0 0 24 0; font-size:12px""><b>Photos captured in these years: " & Ubound(temp1)-1 & " (" & Int(100*(Ubound(temp1)-1)/((Ubound(temp1)-1)+(Ubound(temp2)-1))) & "%) + " & Ubound(temp3)-1 & " duplicates; missing: " & Ubound(temp2)-1 & "</b></p>")
	response.write(outline)

conn.close
%>	
	
</table>
</center><br>
<%response.write(missinglist)%>

<!--#include file="base_code.htm"-->
</body>

</html>