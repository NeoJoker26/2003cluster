<%@ Language=VBScript %>
<% Option Explicit %>
<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Photo comments</title>
<link href="gos2.css" rel=stylesheet>
<style type="text/css">
<!--
p {font-size: 11px;}
input[type=submit] {font-family:verdana,sans-serif; font-size: 11px; padding:3px 4px 4px;}
input[type=submit]:hover {opacity:0.70; filter:alpha(opacity=70);}
.errormsg {margin: 0 0 2px 0; font-size: 11px; color: red;}
-->
</style>
</head>

<body>
<%
		Dim output, urlparm1, urlparm2
		
		'Check if urlparm has been passed from a previous iteration - if not, must be first time, so construct it   
		if Request.QueryString("urlparm1") > " " then 
  			urlparm1 = Request.Querystring("urlparm1")
  			urlparm2 = Request.Querystring("urlparm2") 
  		  else
			urlparm1 = Request.QueryString("eventdate") & Request.QueryString("settype") & Request.QueryString("photoset")
			urlparm2 = "pic" & Request.QueryString("displaynum")
		end if
	
		output = output & "<div style=""line-height: 1.5; padding: 0 10px;"">"
   		output = output & "<p style=""margin: 0 0 4px; font-size: 12px; font-weight: 700"">How to use this viewer</p>"
   		output = output & "<p style=""margin: 0 0 4px;"">Here are some simple tips to get the best from this viewer. If you have other suggestions to add, or have problems using the viewer, please write to steve@greensonscreen.co.uk.</p>"  
   		output = output & "<p style=""margin: 12px 0 4px; font-size: 12px; font-weight: 700"">The thumbnails</p>"
   		output = output & "<p style=""margin: 0 0 4px;"">Click or tap on a thumbnail to expand the image. To return to the thumbnails, click/tap on the close button at the top of image.</p>"  
   		output = output & "<p style=""margin: 0 0 4px;"">However, the best way to view the complete set is to choose the first thumbnail and then move forward through the expanded views.</p>"
   		output = output & "<p style=""margin: 12px 0 4px; font-size: 12px; font-weight: 700"">The expanded images</p>"
   		output = output & "<p style=""margin: 0 0 4px;"">(Note that if you want to try out the following suggestions, you will need to close this tips window first.)</p>"
   		output = output & "<p style=""margin: 0 0 4px;"">When you click on a thumbnail, the image will expand to fill the available space in your window. You can then move to the next (or previous) photo by using the navigation controls at the top of the image. Clicking/tapping the image also moves to the next image - a very convenient way to move through the set. If you have a keyboard, you can also move forward and back using the right and left arrow keys.</p>"
		output = output & "<p style=""margin: 0 0 4px;"">If you prefer to use the keyboard, you can remove the navigation controls by pressing the Ctrl key. Press Ctrl again to restore the controls.</p>"
		output = output & "<p style=""margin: 0 0 4px;"">In most browsers, F11 will expand the page to a full screen (move to the next photo to fill the enlarged space). Press F11 again to return.</p>" 
   		output = output & "<p style=""margin: 12px 0 4px; font-size: 12px; font-weight: 700"">A link to this image</p>"
   		output = output & "<p style=""margin: 0 0 4px;"">The URL (web address) for the current image is: </p>"
   		output = output & "<p style=""margin: 0 0 4px; padding: 1px 4px; background-color: #f0f0f0"">www.greensonscreen.co.uk/"
   		if Request.Querystring("viewer") = "viewer2" then
   			output = output & "photos.asp?viewer=viewer2&"
   		  else
			output = output & "photodisplay.asp?"
   		end if	
   		output = output & "parm=" & urlparm1 & "&autoload=" & urlparm2 & "</p>"
   		output = output & "<p style=""margin: 0 0 4px;"">Copy and paste this address if want to link to this photo.</p>"
   		output = output & "</div>"
  		response.write(output)
		
%>

</body>

</html>