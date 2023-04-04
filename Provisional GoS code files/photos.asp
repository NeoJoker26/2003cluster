<%@ Language=VBScript %>
<% Option Explicit %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<%
'Get cookie (NB. must be before <html> - see http://www.highslide-overlayw3schools.com/asp/asp_cookies.asp)
dim gosname, forumname,place
gosname = request.cookies("goscontributor")("gosname")
forumname = request.cookies("goscontributor")("forumname")
place = request.cookies("goscontributor")("place")
%> 

<html>

<head>
<meta http-equiv="content-type" content="text/html; charset=ISO-8859-1" />
<title>Greens on Screen Photos</title>

<link href="gos2.css" rel=stylesheet>
<link rel="stylesheet" type="text/css" href="highslide/highslide.css" />

<style type="text/css">
<!--
#eventsummary {display: inline-block; vertical-align: top; font-size: 11px; margin: 8px 8px 0 0; padding: 0 6px 0 0; text-align: left; border-top: 1px solid #999; 
				border-right: 2px solid #555; border-bottom: 2px solid #555; border-left: 1px solid #999;}
#eventsummary td {padding: 0;}
#eventsummary p {margin: 0 0 2px 6px; font-family: verdana,arial,helvetica,sans-serif; font-size: 11px;}
#eventsummary td img {border:0px none; background-color: transparent; padding: 0; margin-left:0; margin-right:0; margin-top:6px; margin-bottom:0 }

#eventpics {width:1000px; background-color:#ffffff; margin: 18px auto;}
#eventpics img {margin: 8px 8px 0 0; background-color: #fff; padding: 3px; border-top: 1px solid #999; 
				border-right: 2px solid #555; border-bottom: 2px solid #555; border-left: 1px solid #999;}

#eventpics a:hover, a.control:hover {background-color: transparent;}

a.control img {
 opacity:0.4; 
 filter:alpha(opacity=40);
}
a.control:hover img {
 opacity:1.0; 
 filter:alpha(opacity=100);
 cursor: hand; cursor: pointer;
}

.highslide-heading {margin: 0; padding: 0; color: black; width=100%}
.highslide-heading p {font-weight: normal; font-size: 11px; margin: 0; margin-bottom: 2px;} 
.highslide-heading td {margin: 0; padding: 0; vertical-align: top;}

.highslide-caption p {font-weight: normal; font-size: 11px; margin: 3px 0 3px 0;}

.highslide-dimming {background: #363636;}

.highslide-overlay {
	display: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
}

.draggable-header .highslide-header {
    border-bottom: 0px none;
}

.draggable-header .highslide-html-content {
    padding: 0;
}
.draggable-header .highslide-body {
    padding: 0 2px;
}

.left	{float:left; width:25%;}
.left p	{text-align: left; margin: 0; padding: 0 0 0; font-size: 11px;}
.middle {float:left; width:50%;}
.middle p {text-align: center;  margin: 0 10px 2px; padding: 0; font-size: 12px;}
.right {float:right; width:25%;}
.right p, a {text-align: right;  margin: 0; padding: 0 0 0; font-size: 11px;}


#hintstips p {font-weight: normal; font-size: 12px; margin: 12px 24px;}

.scoreaway	{ font-family:verdana,arial,helvetica,sans-serif; font-size: 12px; font-weight: bold; margin: 0; padding: 0;
						display: block; color:#61A76D; background-color: #ffffff; border-style: solid; border-width: 1px; border-color: #000000; }
.scorehome	{ font-family:verdana,arial,helvetica,sans-serif; font-size: 12px; font-weight: bold; margin: 0; padding: 0;
						display: inline-block; color: white; background-color: #61A76D; border-style: solid; border-width: 1px; border-color: black; }
						
.notformobile {display: inline;}	/* changed by Highslide's mobile.js when a mobile device */

.button {
    padding: 4px 4px 5px; 
    font-family: verdana, sans-serif;
    font-weight: normal;
    display: inline-block;
    white-space: nowrap;
    font-size:11px;
    position:relative;
    outline: none;
    overflow: visible;
    cursor: pointer;
    /*border-radius: 4px;*/
}

.button_grey {
    border: 1px solid #505050;
    color: #000000 !important; 
    background: linear-gradient(#f0f0f0,#d0d0d0);    
}
.button_grey:hover {
    opacity:0.70; 
    filter:alpha(opacity=70);        
}

@media only screen and (max-width: 1000px) {
    .middle p {
        font-size: 12px;  
    }
    .left p , .right p,a {
        font-size: 10px;  
    }

}
 
-->
</style>

<script type="text/javascript" src="highslide/highslide-full.min.js"></script>

<script type="text/javascript">

	hs.align = 'center';
	//hs.marginTop = 10;
	//hs.marginBottom = 5;
	hs.transitions = ['expand', 'crossfade'];
	hs.outlineType = 'rounded-white';
	hs.fullExpandOpacity = 40;
	hs.wrapperClassName = 'borderless-html';
	
	hs.fadeInOut = true;

	hs.dimmingOpacity = 1.0;
	hs.showCredits = false;
	hs.lang.restoreTitle = '';
	hs.blockRightClick = true;
	//hs.headingOverlay.position = "top";
	//hs.headingOverlay.opacity = .8;
	hs.captionOverlay.width = "100%";
	hs.captionOverlay.offsetX = "0";
	hs.captionOverlay.offsetY = "5";
	hs.captionOverlay.width = "100%";
	
	
	// Open a specific thumbnail based on querystring input.
	hs.addEventListener(window, "load", function() {
    // get the value of the autoload parameter
    var autoload = /[?&]autoload=([^&#]*)/.exec(window.location.href);
    // virtually click the anchor
    if (autoload) document.getElementById(autoload[1]).onclick();
	});
	
	hs.registerOverlay({
		thumbnailId: null,
		overlayId: 'controls',
		position: 'bottom center',
		relativeTo: 'expander',
		offsetY: 8,
		opacity: 1.0
	});
	
	hs.onKeyDown = function(sender, e) {
	obj=document.getElementById('controls')
	visible=(obj.style.display!="none")   
    if (e.keyCode == 17) {
      	if (visible) {
            obj.style.display="none";
      		return false; 
      	} else {
      		obj.style.display="block"; 
      		return false;
      	}
      }
	};
   
	// disable default close when clicking outside image
	hs.onDimmerClick = function() {
 	return false;
	};
	
	// go to next image when clicking within the image
	hs.Expander.prototype.onImageClick = function() {
    	return hs.next();
	}
   	
   	hs.Expander.prototype.onAfterGetHeading = function () {
    this.heading.innerHTML = this.heading.innerHTML.replace("{captext}", this.custom.captext);
    this.heading.innerHTML = this.heading.innerHTML.replace("{comtext}", this.custom.comtext);
    this.heading.innerHTML = this.heading.innerHTML.replace("Para0_start", "<p style='padding: 6px 2px 4px 2px; margin:0 4px; font-weight: bold;'>"); 
    this.heading.innerHTML = this.heading.innerHTML.replace(/Para1_start/g, "<p style='padding: 0 2px 0 2px; margin:0 4px; font-style:italic; color:#404040'>");  // ... /xxx/g indicates that all xxx should be replaced
    this.heading.innerHTML = this.heading.innerHTML.replace(/Para2_start/g, "<p style='padding: 0 2px 4px 2px; margin:0 4px; font-weight: bold;'>");  // ... /xxx/g indicates that all xxx should be replaced
    this.heading.innerHTML = this.heading.innerHTML.replace(/Para_end/g, "</p>"); 								// ... /xxx/g indicates that all xxx should be replaced 
    this.heading.innerHTML = this.heading.innerHTML.replace(/~/g, "'");			 								// ... /xxx/g indicates that all xxx should be replaced
    this.heading.innerHTML = this.heading.innerHTML.replace(/¬/g, '"');			 								// ... /xxx/g indicates that all xxx should be replaced
	this.heading.innerHTML = this.heading.innerHTML.replace(/\{num1\}/g, this.custom.num);						// ... /xxx/g indicates that all xxx should be replaced, \{ and \} means that the curly bracket is not treated as a regex character
    this.heading.innerHTML = this.heading.innerHTML.replace("{photoseq}", this.custom.photoseq);
    this.heading.innerHTML = this.heading.innerHTML.replace("{photoname}", this.custom.photoname);
	};
	
	hs.Expander.prototype.onAfterGetCaption = function () {
	this.caption.innerHTML = this.caption.innerHTML.replace(/\{num1\}/g, this.custom.num);						// ... /xxx/g indicates that all xxx should be replaced, \{ and \} means that the curly bracket is not treated as a regex character
    this.caption.innerHTML = this.caption.innerHTML.replace("{tot}", this.custom.tot);
    this.caption.innerHTML = this.caption.innerHTML.replace("{author}", this.custom.author);
    this.caption.innerHTML = this.caption.innerHTML.replace("{captext}", this.custom.captext);
	};
		
	// Repositioning the popup after resizing browser window
	hs.addEventListener(window, 'resize', function() {
	var i, exp;
	hs.page = hs.getPageSize();

	for (i = 0; i < hs.expanders.length; i++) {
		exp = hs.expanders[i];
		if (exp) {
			var x = exp.x,
				y = exp.y;

			// get new thumb positions
			exp.tpos = hs.getPosition(exp.el);
			x.calcThumb();
			y.calcThumb();

			// calculate new popup position
		 	x.pos = x.tpos - x.cb + x.tb;
			x.scroll = hs.page.scrollLeft;
			x.clientSize = hs.page.width;
			y.pos = y.tpos - y.cb + y.tb;
			y.scroll = hs.page.scrollTop;
			y.clientSize = hs.page.height;
			exp.justify(x, true);
			exp.justify(y, true);

			// set new left and top to wrapper and outline
			exp.moveTo(x.pos, y.pos);
		}
	}
	});

</script>


</head>


<body oncontextmenu="return false;"><!--#include file="top_code.htm"-->

<%
Dim phase, parm, eventdate, settype, photoset, photo_name, photo_seq, author, current_status, nopics
Dim output,photofolder,virtualpath,fileCount,i,j,scoreclass,headingline1,headingline2,headingtype,caption,comments,comment_part,display

if Request.QueryString("ea") > "" then server.transfer("photodisplay_register.asp")	'indicates a call from the verification email

phase = Request.QueryString("phase")
parm = Request.QueryString("parm")
if left(parm,5) = "close" then response.redirect("index.asp")
eventdate = left(parm,10)
settype = mid(parm,11,1)
photoset = mid(parm,12)

Dim conn, sql, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_read.inc"--><%
%>

<div id=eventpics align=center>

<%

	sql = "select event_published "
	sql = sql & "from event_control " 
	sql = sql & "where event_date = '" & eventdate & "' "
	sql = sql & "  and event_type = '" & settype & "' "

	nopics = "Y"
	
	rs.open sql,conn,1,2
		if not rs.EOF then
			if rs.Fields("event_published") = "Y" or phase = "review" then nopics = "N"
		end if
	rs.close


	sql = "select count(*) as photocount "
	sql = sql & "from photo_event " 
	sql = sql & "where date = '" & eventdate & "' "
	sql = sql & "  and type = '" & settype & "' " 
	sql = sql & "  and photo_set = " & photoset 
	sql = sql & "  and comment_seq = 0 "
	sql = sql & "  and photo_seq > 0 "
 
	rs.open sql,conn,1,2
		filecount = rs.Fields("photocount")
	rs.close
	
	sql = "select author, type, photo_seq, comment_seq, title_pre1, title_pre2, title, title_post, photo_name, text "
	sql = sql & "from photo_event " 
	sql = sql & "where date = '" & eventdate & "' "
	sql = sql & "  and type = '" & settype & "' " 
	sql = sql & "  and photo_set = " & photoset
	sql = sql & "  and photo_seq > 0 " 
	sql = sql & "  and deleted is null " 
	sql = sql & " order by photo_seq, comment_seq "
 
	rs.open sql,conn,1,2
	
	if not rs.EOF then
	
    	headingline1 = rs.Fields("title_pre1") 
    	headingline2 = rs.Fields("title")
    	if rs.Fields("title_post") > "" then headingline2 = headingline2 & " (" & rs.Fields("title_post") & ")"
	  	
	  	scoreclass = "scorehome"
	  	
		Select Case settype
			Case "M"
				photofolder = "matchdisplays"
				if left(rs.Fields("title"),6) <> "Argyle" then scoreclass = "scoreaway"
				headingtype = "Match Photos"
			Case "F","O"
				photofolder = "othermatches"
				if left(rs.Fields("title"),6) <> "Argyle" then scoreclass = "scoreaway"
				headingtype = "Match Photos"
 			Case "H"
 				photofolder = "homepark"
 				headingtype = "Home Park Developments"
 			Case "S"
 				photofolder = "seasondisplays"
 				headingtype = ""
 			Case else
				photofolder = "randomdisplays"
				headingtype = "General Events"
		End Select
		
		virtualPath = photofolder & "/" & eventdate & "/" & photoset
    	 
        output = output & "<span id=""eventsummary"">"
      	output = output & "<table style=""border-collapse:collapse;"" cellpadding=""0"" cellspacing=""0"" border=""0"">"        
      	output = output & "<tr>"
        output = output & "<td>"
        output = output & "<p style=""margin-top: 6px"">"
        output = output & "<b>"
        output = output & rs.Fields("title_pre1")		'Date
        output = output & "</b></p>"
        output = output & "<p style=""margin-bottom: 4px;"">"
        output = output & rs.Fields("title_pre2")		'Competition (in the case of a match)
        output = output & "</p>"
        output = output & "<p>"
        output = output & "<div style=""padding: 0;"" align=""left"">"
		'Result
        output = output & "<table cellspacing=""0"" cellpadding=""0"" border=""0"" style=""border-collapse: collapse; margin-top:6px;"">"
        output = output & "<tr>"
        output = output & "<td align=""left"" style=""border-left-width:0;"" class=""" & scoreclass & """><p style=""margin: 3px 6px 4px 4px; max-width: 220px;""> "
        output = output & rs.Fields("title")	   
        output = output & "</p></td>"
        output = output & "</tr>"
       	output = output & "</table>"
       	output = output & "</div>"
        output = output & "</td></tr>"
          
        output = output & "<tr><td>"
		output = output & "<p style=""margin-top: 5px; margin-bottom: 8px; color: black ; font-weight: normal;"">"
		author = trim(rs.Fields("author"))
		if settype = "S" then
		  	output = output & "... with grateful thanks to<br>all GoS photographers."
		  else
		  	if author > "" then output = output & "Photos from " & rs.Fields("author")
		end if          
        output = output & "</p></td>"
        output = output & "</tr>"
       	output = output & "</table>"
      	output = output & "</span>"
    	
		i = 1
		comments = ""
		j = 0
	
  		Do While Not rs.EOF
  		
  			if rs.Fields("comment_seq") = 0 then
  			
  				'Check if this is the first photo
  				if i > 1 then
  					'Previous photo has been processed, so write out
  					output = output & "<a id=""pic" & i-1 & """ class=""highslide"" onclick=""return hs.expand(this,{slideshowGroup: 'photos', headingId: 'the-heading', captionId: 'the-caption'},{num:" & i-1 & ",tot:" & filecount & ",photoseq:" & photo_seq & ",photoname:'" & photo_name & "',author:'" & author & "',captext:'" & caption & "',comtext:'" & comments & "'})"" " 
    				output = output & "href=""" & virtualpath & "/" & photo_name & """>" 
    				output = output & "<img title=""" & caption & """ alt=""" & i-1 & """ src=""" & virtualpath & "/" & left(photo_name,len(photo_name)-4) & ".JPEG" & """></a>" 
        		end if  	
   		
   				'Now for this photo
   				caption = replace(rs.Fields("text") & " ","'","\'")	'the trailing blank gets over a mysterious problem when there's an apostrophe at the end of the caption
   				caption = left(caption,len(caption)-1)				'now remove that blank
   				caption = replace(caption,"""","&quot;")
   				photo_name = rs.Fields("photo_name")
   				photo_seq = rs.Fields("photo_seq")
   				comments = ""
   				j = 0				 	
   				i = i + 1
       		
       		  else
       		  	    	
       		  	if j = 0 then comments = comments & "Para0_start" & "Your comments:" & "Para_end"
       		  	j = j + 1    		  	
				comment_part = split(rs.Fields("text"),"^")
				comments = comments & "Para1_start" & "From " & comment_part(0)
				if comment_part(1) > "" then comments = comments & " (" & comment_part(1) & ")"	
				comments = comments & ", " & comment_part(2) & ":" & "Para_end"
				comments = comments & "Para2_start" & comment_part(3) & "Para_end"
				comments = replace(comments,"'","~")	'temporarily replace any apostrophes with ~ to allow onclick to work
				comments = replace(comments,"""","¬")	'temporarily replace any quotes ¬ to allow onclick to work
				comments = replace(comments,vbCrLf," ")	'remove any attempt to skip to a new line and replace with a blank
				
				
       		end if  
			   
			rs.MoveNext
		
		Loop
	
	end if
	rs.close
	
	if nopics = "Y" then
	
		output = "Sorry, there are no photos available for this date"
		
	  else
		
		'Final photo has been processed, so write out
  		output = output & "<a id=""pic" & i-1 & """ class=""highslide"" onclick=""return hs.expand(this,{slideshowGroup: 'photos', headingId: 'the-heading', captionId: 'the-caption'},{num:" & i-1 & ",tot:" & filecount & ",photoseq:" & photo_seq & ",photoname:'" & photo_name & "',author:'" & author & "',captext:'" & caption & "',comtext:'" & comments & "'})"" " 
    	output = output & "href=""" & virtualpath & "/" & photo_name & """>" 
    	output = output & "<img title=""" & caption & """ alt=""" & i-1 & """ src=""" & virtualpath & "/" & left(photo_name,len(photo_name)-4) & ".JPEG" & """></a>" 

	      		
   		output = output & "<div class=""highslide-overlay"" style=""width:120px; padding: 4px 5px 3px; text-align:center; background-color: white;"" id=""controls"">"
       	
		output = output & "<a class=""control"" href=""#"" onclick=""return hs.previous(this)"">"
    	output = output & "<img style=""border: 0; margin: 0 18px 0 0;"" src=""images/arrow-left.gif""></a>"
    	output = output & "<a class=""control"" href=""#"" onclick=""return hs.close(this)"">"
    	output = output & "<img style=""border: 0; padding-bottom: 3px; margin: 0 18px 0 0"" src=""images/close.gif""></a>"
		output = output & "<a class=""control"" href=""#"" onclick=""return hs.next(this)"">"
    	output = output & "<img style=""border: 0; margin: 0 0 0 0;"" src=""images/arrow-right.gif""></a>"    		        			    			
    
    	output = output & "</div>"    		  			       	
        	      		
        	
		output = output & "<div class=""highslide-heading"" id=""the-heading"";>" 
	  
		output = output & "<div class=""left"">"
		output = output & "<p>GoS: " & headingtype & "</p>"
     	output = output & "</div>"
    
 		output = output & "<div class=""middle"">"
   		output = output & "<p style=""color:#61A76D; font-weight: 700;"">" & headingline2 & "</p>"
     	output = output & "</div>"
 
 		output = output & "<div class=""right"">"
   		output = output & "<p>" & headingline1 & "</p>" 
     	output = output & "</div>"		
     	   
		output = output & "</div>"
	
		output = output & "<div class=""highslide-caption"" style=""padding: 3px 5px; background-color: white;"" id=""the-caption"">" 
		
		output = output & "<div class=""left"">"   	  		   
    	if author > "" then
    		output = output & "<p>{num1} of {tot} from {author}<br>"
       	  else
      		output = output & "<p style=""margin:0"">{num1} of {tot}<br>"
        end if
        
  		output = output & Chr(169) & " Greens on Screen " & year(date) & "</p>"
  		output = output & "</div>"
  	
		output = output & "<div class=""middle"">"
  		output = output & "<p style=""padding-top: 5px"">{captext}</p>"
 		output = output & "</div>"
 		
 		output = output & "<div class=""right"">"  			
		output = output & "<a class=""highslide button button_grey"" style=""float:right"" href=""#"" onclick=""return hs.htmlExpand (this, { contentId: 'contents', wrapperClassName: 'draggable-header'"
		output = output & ",src: 'photos_tips.asp?viewer=viewer2&eventdate=" & eventdate & "&settype=" & settype & "&photoset=" & photoset & "&displaynum={num1}" & "'"
    	output = output & ", objectType: 'iframe', width: 550, targetX: 'overlay1pos 253px', targetY: 'overlay1pos -27px' })"">" 
    	output = output & "Viewing tips & image link</a>"
     	output = output & "</div>"	
    	   		  			  
    	output = output & "</div>" 
    
    end if
	
response.write(output)

conn.close

%>



<!--#include file="base_code.htm"-->

</body>
</html>