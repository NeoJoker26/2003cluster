<%@ Language=VBScript %> 
<% Option Explicit %>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>GoS Team Photo</title>
<link rel="stylesheet" type="text/css" href="gos2.css">
<script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script>
<script language="JavaScript">
$(document).ready(function(){
    $('.photo img').on('click',function() {
        $(this).toggleClass("fullimg");
    });
});

if (window.Event)
document.captureEvents(Event.MOUSEUP);
function nocontextmenu() {
event.cancelBubble = true, event.returnValue = false;
return false;
}
function norightclick(e) {
if (window.Event) {
if (e.which == 2 || e.which == 3) return false;
}
else if (event.button == 2 || event.button == 3) {
event.cancelBubble = true, event.returnValue = false;
return false;
}
}
if (document.layers)
document.captureEvents(Event.MOUSEDOWN);
document.oncontextmenu = nocontextmenu;
document.onmousedown = norightclick;
document.onmouseup = norightclick;
</script>
</head>

<body><!--#include file="top_code.htm"-->

<%
Dim season, output, prevhold, thishold, nexthold, texthold, outputhold, year1
Dim conn, sql, rs 

year1 = Request.QueryString("year")

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

thishold = ""

sql = "select 'B' as ind, years, seq_no, text "
sql = sql & "from team_photo "
sql = sql & "where years like '" & year1 & "%' "
sql = sql & "union "
sql = sql & "select 'A', max(years), null, null "
sql = sql & "from team_photo "
sql = sql & "where substring(years,1,4) < '" & year1 & "' "
sql = sql & "union "
sql = sql & "select 'C', min(years),null, null "
sql = sql & "from team_photo "
sql = sql & "where substring(years,1,4) > '" & year1 & "' "
sql = sql & "order by ind, seq_no "

rs.open sql,conn,1,2
Do While Not rs.EOF

	select case rs.Fields("ind")
		case "A"
			prevhold = rs.Fields("years")
		case "C"
			nexthold = rs.Fields("years")
		case else
			thishold = rs.Fields("years")
			outputhold = outputhold & "<div class=""photo"">"
			outputhold = outputhold & "<img src=""images/teamphotos/" & thishold
			if rs.Fields("seq_no") > 1 then outputhold = outputhold & "_" & rs.Fields("seq_no")
			outputhold = outputhold & ".jpg"">"
			outputhold = outputhold & "</div>"
			
			texthold = rs.Fields("text")
			if not isnull(texthold) then texthold = replace(texthold,"|p|","<p>")
			outputhold = outputhold & "<div class=""caption"">" & texthold & "</div>"
	end select
	
	rs.Movenext
Loop
rs.close

output = "<div id=""teamphoto"">"

	output = output & "<div id=""header"">"
		output = output & "<div class=""head1"">"
		if prevhold > "" then output = output & "<a class=""button"" href=""teamphoto.asp?year=" & left(prevhold,4) & """>" & prevhold & "</a>"
		output = output & "</div>"
	
		output = output & "<div class=""head2"">PLYMOUTH ARGYLE " & thishold & "</div>"
	
		output = output & "<div class=""head3"">"
		if nexthold > "" then output = output & "<a class=""button"" href=""teamphoto.asp?year=" & left(nexthold,4) & """>" & nexthold & "</a>"
		output = output & "</div>"
	output = output & "</div>"
	output = output & "<p style=""margin:0 auto 9px; font-size:11px;"">Click image to enlarge; again to reduce</p>"

response.write(output & outputhold & "</div>")

%>
<!--#include file="base_code.htm"-->
</body>
</html>