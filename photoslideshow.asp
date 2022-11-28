<%@ Language=VBScript %>
<% Option Explicit %>

<%
Dim phase, parm, interval
parm = Request.QueryString("parm")
if parm > "" then
	interval = Request.QueryString("interval")
	phase = Request.QueryString("phase")
  else
  	parm = Request.Form("sschoice")
  	interval = Request.Form("delay")
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html>
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

.highslide-resize {display: none}

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
.left p	{text-align: left; margin: 0; padding: 0; font-size: 11px;}
.middle {float:left; width:50%;}
.middle p {text-align: center; margin: 0 10px 3px; padding: 0; font-size: 10px;}
.right {float:right; width:25%;}
.right p, a {text-align: right;  margin: 0; padding: 0; font-size: 11px;}

					
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
	hs.transitions = ['expand', 'crossfade'];
	hs.outlineType = 'rounded-white';
	hs.fullExpandOpacity = 0;
	hs.wrapperClassName = 'borderless-html';
	hs.fadeInOut = true;
	hs.dimmingOpacity = 1.0;
	hs.showCredits = false;
	hs.lang.restoreTitle = '';
	hs.blockRightClick = true;
	hs.enableKeyListener = false;
	hs.captionOverlay.offsetX = "0";
	hs.captionOverlay.offsetY = "5";
	hs.captionOverlay.width = "100%";

	var interval
	<%
	if interval > 499 and interval < 9000 then 
		response.write("interval = " & interval)
	  else
		response.write("interval = 5000")
	end if  
	%>

	// Add the slideshow function
	hs.addSlideshow ({
	interval: + interval ,
	repeat: false
	});
 
	// Open the first image on page load
	hs.addEventListener(window, "load", function() {
    	// click the element virtually
		document.getElementById("autoload").onclick();
	});
	
	var click
	// Pause when clicking within the image
	hs.Expander.prototype.onImageClick = function() {
		if (click == null || click == 0) {
			click = 1;
			this.slideshow.pause();
		} else {
			click = 0;
			this.slideshow.play();
		}
    	return false
	}
	

	
	hs.Expander.prototype.onImageClick = function() {
		var exp = hs.getExpander();
		if (click == null || click == 0) {
			click = 1;
			this.slideshow.pause();
   			//exp.the-caption.innerHTML = "PAUSED";
		} else {
			click = 0;
			this.slideshow.play();
   			//exp.the-caption.innerHTML = "";
		}
    	return false
	}

	// Pause on last photo
	hs.Expander.prototype.isLast = function() {
		var cur = this.getAnchorIndex();
		return (cur + 1 == hs.anchors.groups[this.slideshowGroup || 'none'].length)		
	}
	hs.Expander.prototype.onBeforeClose = function() {
		if (this.isLast()) return false;
	}

	// Go back one page when slide show is closed
	hs.Expander.prototype.onAfterClose = function() {
    	if (this.a.id == 'last_photo') window.history.go(-1);
	}	
	
	hs.registerOverlay({
		thumbnailId: null,
		overlayId: 'controls',
		position: 'bottom center',
		relativeTo: 'expander',
		offsetY: 22,
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
   
	// Disable default close when clicking outside image
	hs.onDimmerClick = function() {
 	return false;
	};

	// Go back to previous page when slideshow ends
	hs.Expander.prototype.onClose = function() {
 	history.go(-1)
	};
	

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


<body oncontextmenu="return false;"><!--#include file="top_code.htm"-->

<%
Dim eventdate, settype, photoset, photo_name, photo_seq, author, current_status, nopics, photo_date
Dim output,photofolder,virtualpath,fileCount,i,j,scoreclass,headingline1,headingline2,headingtype,caption,comments,comment_part,display
Dim autoplay1, autoplay2

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
    	headingline2 = "<span style=""margin:0; padding:0; color:#4b8054; font-weight: 700; font-size: 13px;"">" & ucase(rs.Fields("title")) & "</span><br>Click on image to pause; again to resume"
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
  			Case "W"
 				photofolder = "homepark"
 				headingtype = "South Side Slideshows"
 			Case else
				photofolder = "randomdisplays"
				headingtype = "General Events"
		End Select
		
		virtualPath = photofolder & "/" & eventdate & "/" & photoset
    	 
		i = 1
		comments = ""
		j = 0
	
  		Do While Not rs.EOF
  		
  			
  				'Check if this is the first photo
  				if i > 1 then
  					'Previous photo has been processed, so write out
  					output = output & "<a " & autoplay1 & " class=""highslide"" onclick=""return hs.expand(this,{" & autoplay2 & "slideshowGroup: 'photos', headingId: 'the-heading', captionId: 'the-caption'},{num:" & i-1 & ",tot:" & filecount & ",photoseq:" & photo_seq & ",photoname:'" & photo_name & "',author:'" & author & "',captext:'" & caption & "'})"" " 
    				output = output & "href=""" & virtualpath & "/" & photo_name & """></a>" 
  					autoplay1 = ""
  					autoplay2 = "" 
        		end if  
 
   				'Now for this photo

   				photo_name = rs.Fields("photo_name")
   				photo_seq = rs.Fields("photo_seq")
   				   				
   				photo_date = "20" & mid(photo_name,6,2) & "-" & mid(photo_name,8,2) & "-" & mid(photo_name,10,2)
   				photo_date = FormatDateTime(photo_date,vbLongDate)
   				if left(photo_date,1) = 0 then photo_date = mid(photo_date,2)
   				if isdate(photo_date) then
					caption = "This photo: " & photo_date				
				  else
				   	caption = ""
				end if

   				if i = 1 then
   					autoplay1 = "id=""autoload"""
   					autoplay2 = "autoplay:true, "
   				end if
   				j = 0				 	
   				i = i + 1
    
			rs.MoveNext
		
		Loop
	
	end if
	rs.close
	
	if nopics = "Y" then
	
		output = "Sorry, there are no photos available for this date"
		
	  else
		
		'Final photo has been processed, so write out
  		output = output & "<a  id=""last_photo"" class=""highslide"" href=""" & virtualpath & "/" & photo_name & """ onclick=""return hs.expand(this,{slideshowGroup: 'photos', headingId: 'the-heading', captionId: 'the-caption'},{num:" & i-1 & ",tot:" & filecount & ",photoseq:" & photo_seq & ",photoname:'" & photo_name & "',author:'" & author & "',captext:'" & caption & "'})"" " 
  		output = output & "href=""" & virtualpath & "/" & photo_name & """></a>"

	      		
   		output = output & "<div class=""highslide-overlay"" style=""width:120px; padding: 4px 5px 3px; text-align:center; background-color: white;"" id=""controls"">"
       	
		output = output & "<a class=""control"" href=""#"" onclick=""return hs.previous(this)"">"
    	output = output & "<img style=""border: 0; margin: 0 18px 0 0;"" src=""images/arrow-left.gif""></a>"
    	output = output & "<a class=""control"" href=""#"" onclick=""javascript: history.go(-1)"">"
    	output = output & "<img style=""border: 0; padding-bottom: 3px; margin: 0 18px 0 0"" src=""images/close.gif""></a>"
		output = output & "<a class=""control"" href=""#"" onclick=""return hs.next(this)"">"
    	output = output & "<img style=""border: 0; margin: 0 0 0 0;"" src=""images/arrow-right.gif""></a>"    		        			    			
    
    	output = output & "</div>"    		  			       	
        	      		
        	
		output = output & "<div class=""highslide-heading"" style=""padding: 3px 5px;"" id=""the-heading"";>" 
	  
		output = output & "<div class=""left"">"
		output = output & "<p>GoS: " & headingtype & "</p>"
     	output = output & "</div>"
    
 		output = output & "<div class=""middle"">"
   		output = output & "<p>" & headingline2 & "</p>"
     	output = output & "</div>"
 
 		output = output & "<div class=""right"">"
   		output = output & "<p>{captext}</p>"
     	output = output & "</div>"		
     	   
		output = output & "</div>"
	
		output = output & "<div class=""highslide-caption"" style=""padding: 3px 5px;"" id=""the-caption"">" 
		
		output = output & "<div class=""left"">"   	  		   
       	output = output & "<p style=""margin:0"">{num1} of {tot}"
       	output = output & "</div>"
       		
		output = output & "<div class=""right"">"
  		output = output & "<p>" & Chr(169) & " Greens on Screen " & year(date) & "</p>"
 		output = output & "</div>"
 		   	   		  			  
    	output = output & "</div>" 
    
    end if
	
response.write(output)

conn.close

%>


<!--#include file="base_code.htm"-->

</body>
</html>