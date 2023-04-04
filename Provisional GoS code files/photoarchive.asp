<%@ Language=VBScript %> 
<% Option Explicit %> 

<!doctype html>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=ISO-8859-19" />
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GoS-DB Match Page</title>
<link rel="stylesheet" type="text/css" href="gos2.css">
<link rel="stylesheet" type="text/css" href="highslide/highslide.css" />

<style>
<!--
#container { width:980px; margin:15px 0;}
ul {margin: 3px 0; padding: 0;}
li {display: inline-block; 
	width: 225px; min-height: 50px; 
	border: 1px solid #c0c0c0; 
	padding: 2px 3px; margin: 3px 3px; 
	vertical-align: top;vertical-align: top;
	background-color: #e0f0e0;
	}
.hover {border: 1px solid #000000; color: #000000; background-color: #ffffff;}
.year {text-align: left; margin: 12px 0; padding 0;}
.stylemod {text-align:left; margin: 2px 4px; padding 0;}

-->
</style>

<script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script>

<script>
$(document).ready(function(){
	$('li').mouseover(function(){
    		$(this).addClass('hover');
    });
    $('li').mouseout(function(){
    		$(this).removeClass('hover');
    });
	$('li').click(function(){
    		var matchcode = $(this).attr('id');
			var url = 'photos.asp?parm=' + matchcode;
          	$(location).attr('href',url);
    });
});
</script>

</head>

<body>

<!--#include file="top_code.htm"-->

<div id="container">

<%
Dim output, date, dateparts, subject, photonum
Dim conn,sql,rs 


Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

output = "<p style=""margin: 12px auto 0; font-size:18px; color:#006E32;"">Greens on Screen Photo Archive</p>"
output = output & "<p class=""style1"" style=""margin: 3px auto 15px"";>A Collection of GoS's Non-Match Photos</p>"
output = output & "<ul>"

sql = "select event_date, event_type, material_seq, whatsnew_heading, whatsnew_text, material_details1  "
sql = sql & "from event_control " 
sql = sql & "where event_type in ('F','O','E','H') and material_type='I' and event_published = 'Y' "
sql = sql & "order by event_date desc "

rs.open sql,conn,1,2
Do While Not rs.EOF

	dateparts = split(trim(FormatDateTime(rs.Fields("event_date"),1))," ")
	if left(dateparts(0),1) = "0" then dateparts(0) = mid(dateparts(0),2,1)
	dateparts(1) = left(dateparts(1),3)
	date = dateparts(0) & " " & dateparts(1) & " " & dateparts(2)
	subject = ""

	select case rs.Fields("event_type")
		case "H"	
			subject = "Home Park: "
		case "O"	
			subject = "Other Match: "
		case "F"	
			subject = "Friendly: "
		case else
			subject = "Other Event: "
	end select		
	
	if instr(rs.Fields("material_details1")," from") > 0 then
		photonum = split(rs.Fields("material_details1")," from")
	  elseif instr(rs.Fields("material_details1")," thanks to") > 0 then
	  	photonum = split(rs.Fields("material_details1")," thanks to")
	end if

	output = output & "<li id=""" & rs.Fields("event_date") & rs.Fields("event_type") & rs.Fields("material_seq") & """>"
	output = output & "<p class=""style4boldgrey stylemod"">" & date & "</p>"
	output = output & "<p class=""style4 stylemod"">" & subject & rs.Fields("whatsnew_heading") & "</p>"
	output = output & "<p class=""style4 stylemod"">" & rs.Fields("whatsnew_text") & "</p>"
	output = output & "<p class=""style1 stylemod"">" & photonum(0) & "</p>"	
	output = output & "</li>"
	
	rs.Movenext
Loop
rs.close

output = output & "</ul>"

response.write(output)

%>
</div>
<br>
<!--#include file="base_code.htm"-->
</body>
</html>