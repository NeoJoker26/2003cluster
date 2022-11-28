<!doctype html>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=ISO-8859-1" />
<title>GoS Home Park</title>
<link href="images/favicon.ico" rel="shortcut icon">
<link rel="stylesheet" type="text/css" href="gos2.css">
<link rel="stylesheet" type="text/css" href="highslide/highslide.css" />

<style>
#content {width: 980px; margin: 0 auto;}
#current {padding:0 0 20px 0;}
#current h1 {margin:24px 0 4px; text-align:center; font-size:18px;}
#current p {margin:6px 0; text-align: center;}
#currentcentre {margin:0 auto; width: 400px;}
#older {display: none; width: 500px; margin: 12px 0 18px; clear:both;}
#older p {text-align:left;}
.oldershow {padding:4px;}
#slides {width: 500px; clear:both;}
#slides h1 {margin:12px 0 4px; text-align:center; font-size:17px;}
#slides p {margin:6px 0 18px; text-align:left;}
#past {width:90%; padding:0 0 15px 0; clear:both;}
#past h1 {margin:12px 0 4px; text-align:center; font-size:17px;}
#past p {margin:6px 0; text-align:left;}
.image {border-width: 0; margin: 0 4px -2px 0; vertical-align: baseline}
.hover {color: #000000; background-color: #c8e0c7 !important; cursor: pointer;}
.photoleft {float:left; border:1px solid #404040; padding:3px;}
.photoright {float:right; border:1px solid #404040; padding:3px;}
.hide {display: none;}
.show1 {padding-left: 8px;}
.show2 {padding-left: 8px; display: inline;}	/*display inline will override hide on page load*/
.show3 {padding-left: 8px;}
table {display: table; border-collapse: separate; border-spacing: 0; padding: 0;}

</style>

<!-- <script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script> -->
<script type="text/javascript"  src="jquery/jquery-3.4.1.min.js"></script>
<script>
$(document).ready(function(){

	$(".oldershow").hover(function() {
    	$(this).toggleClass("hover");
	});
	
	$(".oldershow").click(function(){
		$(".oldershow").hide('fast');
		$("#older").show('fast');
 	});
 	
	$("#older table tr").click(function() {
        var href = $(this).find("a").attr("href");
        if(href) {
            window.location = href;
        }
    });
    	
    $("#older table tr").hover(function() {
    	$(this).toggleClass("hover");
	});
	
	$("#delay1").click(function(){
		$(".hide").hide('fast');
		$(".show1").show('fast');
	});
	$("#delay2").click(function(){
		$(".hide").hide('fast');
		$(".show2").show('fast');
	});
	$("#delay3").click(function(){
		$(".hide").hide('fast');
		$(".show3").show('fast');
	});

 			
});
</script>

<body>
<!--#include file="top_code.htm"-->

<div id="content">

<img src="images/hpentrance.jpg" width="760" height="202">
<p class="style4bold" style="margin-top:6px; text-align: center">"Plymouth, surely, has the ground location which most clubs would die for" (Simon Inglis)</p>

<div id="current">
<h1 class="style1boldgreen">SOUTH SIDE STORY</h1> 

<%
Dim conn, rs, output, setcounts(99), n, interval 

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs2 = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_read.inc"--><%

	 	sql = "select count(*) as setcount "
		sql = sql & "from event_control " 
		sql = sql & "where event_date >= '2018-01-01' "
		sql = sql & "and event_type = 'H' " 
		sql = sql & "and event_published = 'Y' "  
 
		rs.open sql,conn,1,2
		setcount = rs.Fields("setcount")
		rs.close
		
		n = 0
		
		sql = "select count(*) as thisset_count "
		sql = sql & "from photo_event join event_control on date = event_date "
		sql = sql & "where event_type = 'H' and type = 'H' "
		sql = sql & "and event_published = 'Y' "  
		sql = sql & "and photo_seq > 0 "
		sql = sql & "group by rollup (date) "
		sql = sql & "order by date "
		
		rs.open sql,conn,1,2
		
		Do While Not rs.EOF
			setcounts(n) = rs.Fields("thisset_count")	'Note: setcounts(0) will have the total
			n = n + 1
			rs.MoveNext		
		Loop
		rs.close
		
		array_top_element = n - 1
		
		n = 0
	 	
	 	sql = "select event_date, convert(varchar,event_date,106) as set_date, material_type, material_seq, publish_timestamp, title "
		sql = sql & "from event_control a "
		sql = sql & "cross apply "
		sql = sql & "	(select top 1 title "
		sql = sql & "  	from photo_event b "
		sql = sql & " 	where a.event_date = b.date and a.event_type = b.type and a.material_seq = b.photo_set "
		sql = sql & " 	) c "
		sql = sql & "where event_date >= '2018-01-01' "
		sql = sql & "and event_type = 'H' " 
		sql = sql & "and event_published = 'Y' "  
		sql = sql & "order by publish_timestamp desc "
		
		rs.open sql,conn,1,2
		
		Do While Not rs.EOF
										
			eventdate = rs.Fields("event_date")
			
			if n = 0 then
				
				sql = "select count(*) as photocount "
				sql = sql & "from photo_event " 
				sql = sql & "where date = '" & eventdate & "' "
				sql = sql & "  and type = 'H' " 
				sql = sql & "  and photo_set = 1 "
				sql = sql & "  and photo_seq > 0 "
 
				rs2.open sql,conn,1,2
				photocount = rs2.Fields("photocount")
				rs2.close
	
				output = output & "<img class=""photoleft"" src=""homepark/" & eventdate & "/menutwo/1.jpg"">" 
				output = output & "<img class=""photoright"" src=""homepark/" & eventdate & "/menutwo/2.jpg"">" 
				
				output = output & "<div id=""currentcentre"">" 
				output = output & "<p class=""style1"">The development of the New Mayflower Grandstand,<br>planned to open in the summer of 2019.</p>"
 				output = output & "<p class=""style4bold"" style=""margin: 10px 0; line-height: 1.5;"">" & ucase(weekdayname(weekday(eventdate))) & "<br>" & ucase(formatdatetime(eventdate,1)) & "</p>"
 				output = output & "<p class=""style4boldgreen"" style=""margin-top: 15px""><a  style=""padding: 4px"" href=""photos.asp?parm=" & rs.Fields("event_date") & "H" & rs.Fields("material_seq") & """><img class=""image"" src=""images/camera16.png"">Set " & setcount - count & ":" 
				output = output & "<span style=""margin:0 18px 0 6px"">" & rs.Fields("title") & "</span></a></p>"
				output = output & "<p class=""style1"" style=""margin-top: 6px"">(" & setcounts(array_top_element) & " photos)"
 				output = output & "</div>" 
							
			 elseif n = 1 then 
			 	output = output & "<div id=""older"">"
			 	output = output & "<p class=""style4bold"" style=""text-align:center; margin: 0 0 5px;"">EARLIER SETS</p>"
			 	output = output & "<p class=""style1"" style=""text-align:center; margin: 0 0 15px;"">(" & setcounts(0) - setcounts(array_top_element) & " photos)</p>"
			 	output = output & "<table class=""style4green"">"

			end if
					
			if n > 0 then
				output = output & "<tr>" 
				output = output & "<td class=""style4green"" style=""white-space: nowrap; text-align: right;""><a href=""photos.asp?parm=" & rs.Fields("event_date") & "H" & rs.Fields("material_seq") & """>" & rs.Fields("set_date") & "</a></td>"
				output = output & "<td class=""style4green"" style=""white-space: nowrap;""><img class=""image"" src=""images/camera16.png"">" & setcount - n & ":</td>" 
				output = output & "<td class=""style4green"" style=""white-space: nowrap"">" & rs.Fields("title") & " (" & setcounts(setcount - n) & ")</td>"
				output = output & "</tr>"
			end if
			
			if n = 0 then output = output & "<p class=""style1boldgreen"" style=""margin-top: 24px;""><span class=""oldershow"">CLICK FOR EARLIER SETS</span></p>"	
			n = n + 1
			rs.MoveNext
		Loop	
		
		rs.close
		
		if n > 1 then output = output & "</table></div>"

		output = output & "</div>"
		response.write(output)
%>
		<div id="slides">
		<h1 class="style1boldgreen">SOUTH SIDE SLIDESHOWS</h1>
		<p class="style1" style="width: 360px; margin: 6px 0 0; text-align: center">After more than 7,000 photos of the south side's redevelopment, spanning nearly two years, here are selected shots of each aspect of the project, in slideshow form.</p>
<%
		output = "<form style=""width: 240px; text-align: left; font-size: 11px; margin: 0;"" action=""photoslideshow.asp"" method=""post"" name=""form1"">"
		output = output & "<p class=""style1"" style=""margin: 12px 0 3px; text-align: center"">Choose speed</p>"
		output = output & "<div style=""margin:0 auto; padding-left:15px;"">"
		output = output & "<input id=""delay1"" type=""radio"" name=""delay"" value=""2000"">"
		output = output & "<label for=""delay1"" style=""padding-right:10px;"">Fast</label>"
		output = output & "<input id=""delay2"" type=""radio"" name=""delay"" value=""5000"" checked>"
		output = output & "<label for=""delay2"" style=""padding-right:10px;"">Medium</label>"
		output = output & "<input id=""delay3"" type=""radio"" name=""delay"" value=""8000"">"
		output = output & "<label for=""delay3"">Slow</label>"
		output = output & "</div>"
		output = output & "<div class=""style1"" style=""margin: 9px 0 0;"">"
		output = output & "<input type=""radio"" name=""sschoice"" value=""2019-12-03W1"">Green Taverners Suite<span class=""hide show1"">[8:00]</span><span class=""hide show2"">[20:00]</span><span class=""hide show3"">[32:00]</span><br>"
		output = output & "<input type=""radio"" name=""sschoice"" value=""2019-12-19W1"" checked>The Top Corner<span class=""hide show1"">[7:20]</span><span class=""hide show2"">[18:20]</span><span class=""hide show3"">[29:20 ]</span><br>"
		output = output & "<input type=""radio"" name=""sschoice"" value=""2020-04-24W1"" checked>The Dressing Room Corner<span class=""hide show1"">[6:00]</span><span class=""hide show2"">[15:00]</span><span class=""hide show3"">[24:00 ]</span><br>"
		'output = output & "<p style=""margin: 6px 0 6px 21px"">Further slideshows will appear</p>"
		'output = output & "<p style=""margin: 6px 0 6px 21px"">here in the coming weeks</p>"		
		output = output & "<input type=""submit"" style=""display: block; color: #000000; background-color: #e0f0e0; text-align: center; font-size: 11px; padding: 2px 6px; margin: 15px auto 25px"" value=""Start Slideshow"" name=""B1"">"
		output = output & "</form>"
%>			
		</div>
<%
		'output = output & "<div id=""past"">"
		'output = output & "<h1 class=""style1boldgreen"">HOME PARK'S PAST</h1>"
		'output = output & "<p class=""style1"" style=""margin-top:6px; text-align: center"">Coming later in the year, pictures of Home Park in 2000,<br>the horseshoe development, the new pitch and more</p>" 
		'output = output & "</div>"

		output = output & "</div>"

		response.write(output)
%>

<!--#include file="base_code.htm"-->
<p>&nbsp;</p>
</body>

</html>