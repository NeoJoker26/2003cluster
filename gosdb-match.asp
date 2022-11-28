<%@ Language=VBScript %> 
<% Option Explicit %> 
<%
	dim matchdate, matchyear, matchmon, matchday, matchdecade, matchseason, phase, shortrange, id, lastid
	dim conn, rs, sql
	matchdate = Request.QueryString("date")
	phase = Request.QueryString("phase")
%>

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
#container {
    width:980px;
    margin:15px 0;
}

.navhandle {
    text-align: left;
}

.nav {
    color: #111;
    margin: 0 0 10px;
    padding: 0;
}

.nav,.nav2 ul {
    width:100%; 
    text-align: left;
    margin: 5px 0;
    padding: 0;
    list-style: none;
}
.nav li {
    display: inline-block;
    border: 1px solid #c0c0c0;
    padding: 2px 0 2px 6px;
    margin: 0 9px 4px 0;
    color: #202020;
    font-size: 11px;
}
.nav li:not(.header){
	font-size: 11px;
	cursor: pointer;
}

.nav2 li {
    display: inline-block;
    border: 1px solid #c0c0c0;
    padding: 2px 6px;
    margin: 0 9px 4px 0;
    color: #202020;
    font-size: 11px;
}
.header {
    background-color: #f0f0f0;
    padding: 2px 12px !important;
    width: 48px;
}

.boxwidth {width: 71px;
}

.hover {
    border-color: #909090 !important;
    background-color: #e0f0e0;
}
    
.click {
    border: 1px solid #000000;
    color: #ffffff !important;
    background-color: #79A088;
}

#seasons {
	margin-top: 8px;
}

#matches table {
	table-layout: fixed;
    width:100%;
    margin-left: -8px;
    margin-right: auto;
    empty-cells: hide;
    border-spacing: 8px 4px;
}
#matches td {
    border: 1px solid #c0c0c0;
    padding: 2px 1px 2px 10px;
    color: #202020;
    font-size: 11px;
}

#matches td.cellgrey {
    color: #c0c0c0; 
}

#matches td:first-child {
    padding: 2px 0 2px 6px;
}

#matches td:not(:first-child):not(:empty) {
	cursor: pointer;
}

#matchdetails {
	float:left;
	width:100%;
	min-height: 400px;
	margin: 12px auto;
	text-align: left;
	font-size: 11px;
	line-height: 150%;
}
#matchdetails h1 {
	font-size: 14px;
	color: #404040;
	font-weight: 700;
	margin-bottom: 10px;
}
#matchdetails h2 {
	font-size: 12px;
	color: #3f7855;
	font-weight: 700;
	margin-top: 12px;
}

#report {
	text-align: justify;
	max-width: 60%;
	line-height: 140%;
}

#report-wider {
	text-align: justify;
	max-width: 60%;
	line-height: 140%;
}

#milestones ul {
	list-style-type: circle; 
	padding-left: 18px;
	margin: 0;
}

#chartbuttons ul.nav  {
	margin: 18px 0 24px -18px;
}

#chartbuttons {
	padding-left: 0;
}
 
#milestones ul.nav  {
	margin: 18px 0 24px -18px;
}
 
#milestones ul li.cell {
	font-size: 11px;
	margin: 0 12px 6px 0;
	padding: 0 5px;
}

#milestones a:hover {
	background-color: transparent;
	
}
#milestones a:link {
	color: #000000;
	background-color: transparent;
}

#material {
	float: right;
	width: 36%;
	font-size: 11px;
	margin: 0px 0 12px 12px;
	padding:0;
}

#material .undofloat {
	float: none !important;
}

#chartbuttons li {
    display: inline-block;
    border: 1px solid #c0c0c0;
    padding: 2px 6px;
    margin: 0 9px 4px 0;
    color: #202020;
    font-size: 11px;
}

#tablecontainer {margin-top: -12px}

.thisdate {
	background-color: #e0f0e0;
	} 

h3 { 
	text-align: center;
	margin: 15px auto 12px;
}
	
.bold { font-weight: 700; }
.green { color: green; }
.grey { color: #606060; }
.font11px { font-size: 11px; }
.font12px { font-size: 12px; }
.font15px { font-size: 15px; }

.score { font-size: 16px;	font-weight: 700; margin: 10px 0 6px; }
.aet { font-size: 11px; margin: 0 0 4px; font-weight: 700; }
.penalties { font-size: 11px; margin: 4px 0; font-weight: 700; }
.team { margin: 6px 0 0; }
.opp { color: #606060; }
.goals { margin: 9px 0 0; }
.audio { border-width: 0; vertical-align: middle; margin: 1px 0px 1px 3px; }
.image { border-width: 0; vertical-align: middle; margin: 2px 6px; }
.video { border-width: 0; vertical-align: middle; margin: 2px 6px; } 

	
.caption {
	margin: 0 0 20px -24px;
	padding: 0 24px;
	text-align: center;
	font-size: 11px;
	font-style: italic;
	font-weight: 500;
	}
	
.venue, .attendance, .visitors, .totpoints { margin-right: 18px; }

a:hover {background-color: transparent;}

a.control img {
 opacity:0.4; 
 filter:alpha(opacity=40);
}
a.control:hover img {
 opacity:1.0; 
 filter:alpha(opacity=100);
 cursor: hand; cursor: pointer;
}

.highslide-wrapper .highslide-header ul {margin: 0 0 6px; font-weight: bold; font-size: 12px;}
.highslide-wrapper .highslide-header a {color: black;}
.highslide-dimming {background: black;}

-->
</style>


<script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script>

<script>

$(function () {		// Disable right click
   $('#matchdetails').bind('contextmenu', function (e) {
     e.preventDefault();
   });
  });

$(document).ready(function(){

	$('.navhandle').on('mouseenter','.cell', function(){
   		$(this).not('.header').addClass('hover');
   	});
   		
   	$('.navhandle').on('mouseleave','.cell', function(){
   		$(this).not('.header').removeClass('hover');  
   	});
   	

	$('#decades').on('click','li:not(".header")', function(){
    		$('#decades li').removeClass('click');
    		$(this).not('.header').addClass('click');
    		$('#matches, #matchdetails, #graph1').html('');
    		$('#matches, #matchdetails, #graph2').html('');
    		$('#matches, #matchdetails, #table').html('');
        	$('#seasons').html('').load('gosdb-getmatchpage.asp','pass=1&decade=' + $(this).text(), 
             function(response, status){
              <%
				if not isdate(matchdate) then 
				  	response.write("if (status == 'success') ")
				  	response.write("{$(""#matches"").html('<p class=""font11px green"" style=""text-align:center"">The green numbers are counts of the video links in each of the decades or seasons</p><p class=""font12px"" style=""text-align:center"">Now choose a season and then a match date</p>');} ") 
				end if  
			  %>
           });		
    });  	
    	
	$('#seasons').on('click','li:not(".header")', function(){
    		$('#seasons li').removeClass('click');
    		$(this).not('.header').addClass('click');
    		$('#matchdetails, #graph1').html('');
    		$('#matchdetails, #graph2').html('');
    		$('#matchdetails, #table').html('');
    		$('#matches').html('').load('gosdb-getmatchpage.asp','pass=2&season=' + $(this).text(),
    		 function(response, status){
              <%
				if not isdate(matchdate) then 
				  	response.write("if (status == 'success') ")
				  	response.write("{$(""#matchdetails"").html('<p class=""font11px green"" style=""text-align:center""><img style=""border:0; padding:0 4px 0 0;"" src=""images/video7x12.gif"">indicates that one or more video links are available in that match content</p>');} ") 
				end if  
			  %>
			});
    		localStorage.setItem("season", $(this).text());
    });
    		
	$('#matches').on('click','td:not(".header")', function(){
    		$('#matches td').removeClass('click');
    		$(this).not('.header').addClass('click');
    		$('#graph1').html('');
    		$('#graph2').html('');
    		$('#table').html('');
    		var matchdate1 = $(this).attr('id');
    		var year = matchdate1.substr(0,4);
    		var mon = matchdate1.substr(4,3);
    		var day = matchdate1.substr(7,2);
    		var months = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
    		var monthno = months[mon];
    		var matchdate2 = year + '-' + monthno + '-' + day;
    		$('#matchdetails').html('').load('gosdb-getmatchpage.asp','pass=3&date=' + matchdate1 + '&season=' + localStorage.getItem("season") + '&phase=<%response.write(phase)%>',function(responseTxt,statusTxt,xhr){
            	if(statusTxt == "success") {
    				history.pushState({myTag: true}, '', 'gosdb-match.asp?date=' + matchdate2);
				    hs.updateAnchors(); 
    				}       		
    		});
    });
    	
	$('#matchdetails').on('click','li', function(){
    		$('#matchdetails li').removeClass('click');
    		$(this).addClass('click');
    		var chartbuttonid = $(this).attr('id');
    		var chartid = chartbuttonid.substr(0,1);
    		var chartdate = chartbuttonid.substr(1,10);
    		var ajaxparm = 'date=' + chartdate;
    		if(chartid == "A") {
               	var ajaxpage = 'progressgraphs.asp #graph1';
               	var ajaxdiv = '#chart1'
               	} else if(chartid == "B") {
                var ajaxpage = 'progressgraphs.asp #graph2';
                var ajaxdiv = '#chart2'
                } else {
                ajaxparm += '&source=matchpage'
                var ajaxpage = 'progresstables.asp #table';
                var ajaxdiv = '#tablecontainer'
            } 
           	$(ajaxdiv).html('').load(ajaxpage,ajaxparm,function(responseTxt,statusTxt,xhr){
            	if(statusTxt == "success")
            		if(chartid == "A") {
            	    	$('#chart2').hide();
            	    	$('#table').hide();
            	    	$('#chart1').show();
            	    	} else if(chartid == "B") {
            	    	$('#chart1').hide();
            	    	$('#table').hide();
            	    	$('#chart2').show();
            	    	} else {
            	    	$('#chart1').hide();
            	    	$('#chart2').hide();
            	    	$('#table').show();
              	    }
              	    $('html, body').animate({scrollTop:$(document).height()}, 100);
    		});
    });

    	
	<%
	if isdate(matchdate) then
		matchyear = Year(matchdate)
		matchmon = MonthName(Month(matchdate),True)
		matchday = Day(matchdate)
		
		'cope with e.g. 1920-05-01 being in 1919-20 season and 1910-19 decade
		if Month(matchdate) < 7 then 
			matchseason = matchyear - 1
		  else
		  	matchseason = matchyear
		 end if 	
		matchdecade = left(matchseason,3) & "0"

		response.write("$( document ).ajaxSuccess(function( event, xhr, settings ) { ")
  		response.write("if ( settings.url.indexOf('pass=1') > 0 ) { $('#seasons #s" & matchseason & "').trigger('click'); } ")
  		response.write("else if ( settings.url.indexOf('pass=2') > 0 ) {$('#matches #" & matchyear & matchmon & matchday & "').trigger('click'); } ")
  		response.write("else if ( settings.url.indexOf('pass=3') >0 ) { $( document).unbind('ajaxSuccess'); } ")
  		response.write("});")

		response.write("$('#decades #d" & matchdecade & "').trigger('click');")

	end if


	%>	
});

// Update the page content when the popstate event is called (ties in with history.pushState)
window.onpopstate = function(e) {
  	if (!e.originalEvent.state.myTag) return; // not my problem
	window.location.href = document.location;
};

</script>

</head>

<body>

<!--#include file="top_code.htm"-->

<script type="text/javascript">
	hs.graphicsDir = 'highslide/graphics/';
    hs.outlineType = 'rounded-white';
	hs.align = 'center';
	hs.outlineWhileAnimating = true;
	hs.width = 1000;
	hs.height = 760;
	hs.allowSizeReduction = false;
	hs.preserveContent = false;
	hs.objectLoadTime = 'after';
	hs.dimmingOpacity = 0.6;
	hs.showCredits = true;
	//hs.lang = {creditsText: 'With thanks to PAFC; click here for Argyle Media', creditsTitle: 'Go to Argyle Media on YouTube'};
	//hs.creditsHref = 'https://www.youtube.com/user/argylemedia';
	hs.lang.restoreTitle = 'Click for next programme page. Click bottom-right corner of image to enlarge.';
	hs.blockRightClick = true;
	
	// go to next image when clicking within the image
	hs.Expander.prototype.onImageClick = function() {
    return hs.next();
	}
</script>

<div id="container">

	<div id="decades" class="navhandle">
		<ul class="nav">
		<%
		response.write("<li class=""header"")>Decades</li>")
		
		Set conn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")
		
		%><!--#include file="conn_read.inc"--><%
		
		sql = "select decade, sum(count) as videocount "
		sql = sql & "from "
		sql = sql & "	(select distinct decade, 0 as count "
		sql = sql & "	from season "
		sql = sql & "	union all "
		sql = sql & "	select decade, 1 "
		sql = sql & "	from season a left join event_control b on event_date between date_start and date_end "
		sql = sql & "	where event_published = 'Y' and ((event_type = 'M' and material_type = 'Y') or event_type = 'V') " 
		sql = sql & "	) x "
		sql = sql & "group by decade order by decade "
		
		rs.open sql,conn,1,2
		
		Do While Not rs.EOF
			shortrange = left(rs.Fields("decade"),5) & right(rs.Fields("decade"),2)
			id = "d" & left(rs.Fields("decade"),4)
			if lastid > "" and mid(lastid,2,1) <> mid(id,2,1) then response.write("<br><li class=""header"" style=""visibility: hidden;"")></li>")	'change of century
			lastid = id 
			response.write("<li id=""" & id & """ class=""cell boxwidth"">" & shortrange)
			if rs.Fields("videocount") > 0 then response.write("<span style=""float:right; letter-spacing:-1px; padding:0 1px; margin:1px 1px 0 0; font-size:9px; color:green; background-color:white;"">" & rs.Fields("videocount") & "</span>")		
			response.write("</li>")
			rs.MoveNext
		Loop	
		%>	
		</ul>
	</div>
	
	<div id="seasons" class="navhandle font12px">
	<%
		if not isdate(matchdate) then response.write("<p class=""font11px green"" style=""text-align:center"">The green numbers are counts of video links in each of the decades</p><p class=""font12px"" style=""text-align:center;"">To find a match, first choose a decade</p>")
	%>
	</div>
	
	<div id="matches" class="navhandle"></div>
		
	<div id="matchdetails" class="navhandle"></div>
	        		
	<div class="highslide-caption" style="padding: 5px 10px 0; text-align:center; background-color: white;" id="progcaption">
	<%
		if matchdate = "1992-8-22" then response.write("<p style=""margin: 0 0 6px; text-align:center; font-size: 12px"">COPYRIGHT NOTICE: Copying this or any other match programme page is absolutely forbidden without the express consent of the owning football club.</p>")
	%>
	<a class="control" href="#" onclick="return hs.previous(this)">
	<img style="border: 0px none; margin-left:0; margin-right:18px; margin-top:0; margin-bottom:0" src="images/arrow-left.gif" title="Previous image"></a>
	<a class="control" href="#" onclick="return hs.close(this)">
	<img style="border: 0px none; margin-left:0; margin-right:18px; margin-top:0; margin-bottom:0" src="images/close.gif" title="Back to thumbnails"></a>
	<a class="control" href="#" onclick="return hs.next(this)">
	<img style="border: 0px none; margin: 0;" src="images/arrow-right.gif" title="Next imagex"></a>  		        			    			
	</div>

	<div id="chart1"></div>
	<div id="chart2"></div>	
	<div id="tablecontainer"></div>
	
</div>
<br>
<!--#include file="base_code.htm"-->
</body>
</html>