<%@ Language=VBScript %>
<% Option Explicit %>

<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Greens on Screen</title>
<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
<link rel="stylesheet" type="text/css" href="highslide/highslide.css" />

<style>
<!--
td {font-size: 11px; text-align: left; vertical-align: top}
input, textarea {font-family: "courier new",serif; font-size: 12px;}
-->
</style>

<script type="text/javascript" src="highslide/highslide-full.min.js"></script>

<script type="text/javascript">

	hs.transitions = ['expand', 'crossfade'];
	hs.outlineType = 'rounded-white';
	hs.width = 800; 
	hs.fadeInOut = true;
	hs.dimmingOpacity = 0.8;
	hs.showCredits = false;
	
</script>

</head>

<body>

<!--#include file="top_code.htm"-->

<%
Dim fs,f,Folderjpg,code,eventdate,dateerror,initials,settype,process,photoset,output,uploadpath,photopath,virtphotopath,photofolder,title_pre1,title_pre2,title,title_post,re,newname,author,caption(250,2)
Dim filenames(),filecount,i,j,k,n,x,file,buffer,filebase,fileextension,dotpos,foundfile,foundcaption,nextcaption,firsttime

output = ""
filecount = 0

Dim conn, conndelete, sql, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set conndelete = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_read.inc"--><%

Set fs=Server.CreateObject("Scripting.FileSystemObject")

' Decompose code (should be of the form 2013-08-01:SD)

code = Request.Form("code")
eventdate = left(code,10)
initials = Ucase(rtrim(mid(code,12)))
settype = Request.Form("type")
process = Request.Form("process")

Select Case initials
	Case "WS"
		author = "Will Summers"
	Case "SD"
		author = "Steve Dean"
	Case "GB"
		author = "Gill Black"
	Case "MT"
		author = "Malcolm Townrow"
	Case "BW"
		author = "Bob Wright"
	Case "RW"
		author = "Bob Wright"
	Case Else
		author = "Unknown"
End Select

if author = "Unknown" or mid(code,11,1) <> ":" then
	output = "Invalid Code, please check"
else

 dateerror = 0

 if settype = "M" then
	
	photofolder = "matchdisplays"
	
	if process = "2" or process = "3" then action2and3
	
	' Get match details
	sql = "select season_no, opposition, homeaway, goalsfor, goalsagainst, pensfor, pensagainst, competition, name_then_short "
	sql = sql & "from v_match_season a join opposition b on a.opposition = b.name_then " 
	sql = sql & "where a.date = '" & eventdate & "' "

	rs.open sql,conn,1,2

	if rs.recordcount > 0 then
	
		' Date
		title_pre1 = FormatDateTime(eventdate,1)
    	
    	
    	' Competition
    	title_pre2 = rs.Fields("competition")
          
		' Result
    	if rs.Fields("homeaway") = "H" then
	  	 	title =  "Argyle " & rs.Fields("goalsfor") & " " & rs.Fields("opposition") & " " & rs.Fields("goalsagainst")
	  		if not isnull(rs.Fields("pensfor")) then
	  			title_post = "Pens: " & rs.Fields("pensfor") & "-" & rs.Fields("pensagainst")
			end if
  		  else
 	  	 	title = rs.Fields("opposition") & " " & rs.Fields("goalsagainst") & " Argyle " & rs.Fields("goalsfor")
	  		if not isnull(rs.Fields("pensfor")) then
	  			title_post = "Pens: " & rs.Fields("pensagainst") & "-" & rs.Fields("pensfor")
			end if	
		end if
		rs.close
	  
	  else
	  		
	  		rs.close
	  		
	  		'No match details found - could be adding captions before match has been added to GoS-DB, so check if a match was due on this date
	  		
			sql = "select count(*) as rowcnt "
			sql = sql & "from season_this " 
			sql = sql & "where date = '" & eventdate & "' "

			rs.open sql,conn,1,2
			
			if rs.Fields("rowcnt") > 0 then
	  			title_pre1 = FormatDateTime(eventdate,1)
	  			title_pre2 = ""
				title = ""
				title_post = ""
		  		dateerror = 1	  		
			  elseif eventdate = "2000-07-01" then 	'a dumnmy date used for test submissions 
				title_pre1 = FormatDateTime(eventdate,1)
				title_pre2 = "Test Photo Submission"
				title = "Argyle 0 Fictitious Opposition 0"
				title_post = ""
			  else
			  	dateerror = 2
			end if
			rs.close
			
	end if

 elseif settype = "F" then
 
	photofolder = "othermatches"
	
	if process = "2" or process = "3" then action2and3
	
	title_pre1 = FormatDateTime(eventdate,1)
	title_pre2 = "Pre-season Friendly"
	title = ""
	title_post = ""

 elseif settype = "O" then
 
	photofolder = "othermatches"
	
	if process = "2" or process = "3" then action2and3
	
	title_pre1 = FormatDateTime(eventdate,1)
	title_pre2 = ""
	title = ""
	title_post = ""

 elseif settype = "H" or settype = "W" then
 
	photofolder = "homepark"
	
	if process = "2" or process = "3" then action2and3
	
	title_pre1 = FormatDateTime(eventdate,1)
	title_pre2 = ""
	title = ""
	title_post = ""

 elseif settype = "S" then
 
	photofolder = "seasondisplays"
	
	if process = "2" or process = "3" then action2and3
	
	title_pre1 = ""
	title_pre2 = ""
	title = ""
	title_post = ""
	
 else
 
	photofolder = "randomdisplays"
	
	if process = "2" or process = "3" then action2and3
	
	title_pre1 = FormatDateTime(eventdate,1)
	title_pre2 = ""
	title = ""
	title_post = ""
	
 end if

 if dateerror < 2 then

	' Now try for existing captions in photo_event and use those (indicating a second time or more)  

		sql = "select author, title_pre1, title_pre2, title, title_post, photo_set, photo_seq, photo_name, text "
		sql = sql & "from photo_event " 
		sql = sql & "where date = '" & eventdate & "' "
		sql = sql & "  and type = '" & settype & "' "
		sql = sql & "  and initials = '" & initials & "' "
		sql = sql & "  and comment_seq = 0 "	'0=caption	
		sql = sql & "order by photo_seq "

		rs.open sql,conn,1,2
	
		if not rs.EOF then
			author = rs.Fields("author")
			title_pre1 = rs.Fields("title_pre1")
			title_pre2 = rs.Fields("title_pre2")
			title = rs.Fields("title")
			title_post = rs.Fields("title_post")
			photoset = rs.Fields("photo_set")
			virtphotopath = photofolder & "/" & eventdate & "/" & photoset
			photopath = Server.MapPath(virtphotopath)
		end if
			
   		Do While Not rs.EOF
   		
			caption(filecount,0) = rs.Fields("photo_name")
			caption(filecount,1) = rs.Fields("text")
			caption(filecount,2) = rs.Fields("photo_seq")
			filecount = filecount + 1
			rs.MoveNext
		
		Loop
		rs.close
		
		firsttime = ""	
		if process = "1" and filecount = 0 then firsttime = "y"
		
		nextcaption = filecount		'value held if process type 2 and new captions will be added


	' If process type 2, or no captions so far (therefore first time for process 1, or process 3) 

	if process = 2 or firsttime = "y" then

		if firsttime = "y" then

			'Get next photo_set for this date (and hope two sets don't come in at the same time in the next few seconds!)
	
			if fs.FolderExists(Server.MapPath(photofolder & "/" & eventdate)) = false then
				set f=fs.CreateFolder(Server.MapPath(photofolder & "/" & eventdate))
				set f=nothing
			end if
			
			' Get the next set number to to allocate to this set, but first check that it hasn't got one already.
			' This can happen if the user displays the caption boxes for the first time (i.e. no captions) but then refreshes.
			' In such a case, grabbing the next set number would leave a hole in the sequence, which is not a big problem
			' but can be confusing and therefore best avoided. 
			' NB. A hole will still occur if the session times out on an empty caption screen, and then the user refreshes.
			  
		  	if Session("thisset") = "" then Server.Execute("photocaptions1_getset.asp")

		  	photoset = Session("thisset")
	
			if fs.FolderExists(Server.MapPath(photofolder & "/" & eventdate & "/" & photoset)) = false then
				set f=fs.CreateFolder(Server.MapPath(photofolder & "/" & eventdate & "/" & photoset))
				set f=nothing
			end if
	
		end if
	
	
		virtphotopath = photofolder & "/" & eventdate & "/" & photoset
		
		photopath = Server.MapPath(virtphotopath)
	
		'Get photos from upload area
 
		uploadpath = replace(Server.MapPath("\"),"httpdocs","private/GoSin/" & initials & "/" & eventdate & "/sendtogos")
		
		If fs.FolderExists(uploadpath) = false then	uploadpath = replace(Server.MapPath("\"),"httpdocs","private/GoSin/" & initials & "/" & eventdate)

		If fs.FolderExists(uploadpath) = true then
	
			Set Folderjpg = fs.GetFolder(uploadpath)
			if settype = "W" then						
				fileCount = Folderjpg.files.count		'slideshows don't have thumbnails
			  else
				fileCount = Folderjpg.files.count / 2 	'folder contains photos and thumnbnails
			end if
   			redim filenames(fileCount-1)
   			n = 0 
   		
   			for each file in Folderjpg.files
   		
   				dotpos = InstrRev(file.name,".")		'look for dot
   				filebase = left(file.name,dotpos-1)		'base file name
   				fileextension = mid(file.name,dotpos)	'file extension
   							
   				if lcase(fileextension) = ".jpg" then 
       				filenames(n) = file.name
       			
       				'check for invalid characters in base file name and remove
       				set re = new regexp
					re.global = true
					re.pattern = "[^a-zA-Z0-9 \-_]"			'\- distinguishes it as a hyphen rather than an 'a-z'-type range character, so this pattern allows alphanumeric, space, - and _  
								
					newname = re.replace(filebase,"")
							
					if newname <> filebase then
						fs.MoveFile uploadpath & "/" & filebase & fileextension, uploadpath & "/" & newname & fileextension		'rename the jpg file to remove invalid characters
   			
						if settype <> "W" then fs.MoveFile uploadpath & "/" & filebase & ".jpeg", uploadpath & "/" & newname & ".jpeg"	'not forgetting the thumbnail

						filebase = newname
						filenames(n) = filebase & fileextension
					end if	
       		
					'now copy the file from the upload area to matchdisplays (or equivalent)
				
					fs.CopyFile uploadpath & "/" & filenames(n), photopath & "/" & filenames(n)
					if settype <> "W" then fs.CopyFile uploadpath & "/" & filebase & ".jpeg", photopath & "/" & filebase & ".jpeg"		 'not forgetting the thumbnail
			
					n = n + 1
				
       			end if
   			next
   		 
	 		'sort into filename order
	 		for i = 0 to fileCount-1 
   		   
   				for j = (i + 1) to fileCount-1 
       				if strComp(filenames(i),filenames(j),0) = 1 then 
	      		    	buffer = filenames(j) 
    	      			filenames(j) = filenames(i) 
               			filenames(i) = buffer
          			end if 
        		next
        	
        		foundcaption = 0
        	
        		'if process 2 then check if this filename already has a caption
        		if process = 2 then
        			for k = 0 to nextcaption-1 
        				if caption(k,0) =  filenames(i) then 
        					foundcaption = 1
        					exit for
        				end if	
					next	
				end if 
        	
        		if foundcaption = 0 then   	      	
            		caption(nextcaption,0) =  filenames(i) 
        			caption(nextcaption,1) = " "
        			nextcaption = nextcaption + 1
        		end if

    		next
    	   
    	end if
    
    	'Finally, if process 2, tidy up any orphan captions
    	if process = 2 then
  
    		for i = 0 to Ubound(caption,1)
    			foundfile = 0
    			for j = 0 to UBound(filenames)
    				if caption(i,0) = filenames(j) then
    					foundfile = 1
    					exit for 
    				end if
    			next
    		
    			'if no file found, it must be an orphan caption - mark by changing file name in caption array
    			if foundfile = 0 then caption(i,0) = "process 2 orphan"   	
    		next
    			
    	end if
    
	end if	
		
	output = output & "<tr><td>Photographer's name:</td>"
	output = output & "<td colspan=""2""><input type=""text"" name=""author" & """ size=""30"" value=""" & author & """> (change if you are submitting on behalf of another)"
	output = output & "</td></tr>"

	output = output & "<tr><td>Pre-Title 1:</td>"
	output = output & "<td colspan=""2""><input type=""text"" name=""title_pre1" & """ size=""40"" value=""" & title_pre1 & """>"
	output = output & "</td></tr>"

	output = output & "<tr><td>Pre-Title 2:</td>"
	output = output & "<td colspan=""2""><input type=""text"" name=""title_pre2" & """ size=""40"" value=""" & title_pre2 & """>"
	if settype = "E" then output = output & " (add venue or location here, e.g. Tribute Legends' Lounge)"
	if settype = "O" then output = output & " (add competition name here, e.g. FA Youth Cup)"
	output = output & "</td></tr>"

	output = output & "<tr><td>Title:</td>"
	output = output & "<td colspan=""2""><input type=""text"" name=""title" & """ size=""50"" value=""" & title & """>"
	if settype = "E" then output = output & " (add short title for event here)"
	if settype = "F" or settype = "O" then output = output & " <span style=""color:red"">(add score here, e.g. Saltash United 0 Argyle 3)</span>"
	output = output & "</td></tr>"

	output = output & "<tr><td>Post-Title:</td>"
	output = output & "<td style=""padding-bottom: 24px"" colspan=""2""><input type=""text"" name=""title_post" & """ size=""50"" value=""" & title_post & """> (do not change)"
	output = output & "</td></tr>"

	x = 1
	
	output = output & "<input type=""submit"" name=""b1"" value=""Apply All"" style=""font-family:verdana,arial; font-size: 12px;"">" 	'no caption boxes for slideshows
	  
	for n = 0 to fileCount-1 

		if caption(n,0) <> "process 2 orphan" then 
	
			output = output & "<tr>"
			if settype <> "W" then output = output & "<td style=""text-align: centre"" width=""140"">" & "<a class=""highslide"" onclick=""return hs.expand(this)""  href=""" & virtphotopath & "/" & caption(n,0) & """><img border=2 src=""" & virtphotopath & "/"  & left(caption(n,0),len(caption(n,0))-4) & ".jpeg" & """>" & "</a></td>"
			output = output & "<td><p style=""color: #808080; margin: 0 0 6 4"">" & caption(n,0) & "</p>" 
			output = output & "<input type=""hidden"" name=""filename"" value=""" & caption(n,0) & """>"
			output = output & "<input style=""font-family:verdana,arial; display:block; margin:3 0;"" type=""text"" name=""sequence" & """ size=""2"" "
			if firsttime = "" and caption(n,2) = 0	then	'logically deleted
				output = output & "value=""0"">"				
			  else
				output = output & "value=""" & x*10 & """>"	
				x = x + 1
			end if		
			output = output & "<textarea style=""width:100%"" name=""caption"" rows=""2"">" & caption(n,1) & "</textarea>"
			output = output & "</td><td>"
			if n mod 5 = 0 then output = output & "<input type=""submit"" name=""b1"" value=""Apply All"" style=""font-family:verdana,arial; font-size: 12px;"">"
			output = output & "</td></tr>" 
		
		end if
	next

	output=output & "<input type=""hidden"" name=""code"" value=""" & code & """>"
	output=output & "<input type=""hidden"" name=""type"" value=""" & settype & """>"
	output=output & "<input type=""hidden"" name=""process"" value=""" & process & """>"
	output=output & "<input type=""hidden"" name=""photoset"" value=""" & photoset & """>"

 end if
 
end if


response.write("<div style=""width: 980px; margin: 0 auto; text-align: left;"">")

response.write("<p style=""text-align: center; margin-top: 12; margin-bottom: 15"">")
response.write("<font color=""#47784D"" style=""font-size: 18px""><b>Greens on Screen Photo Captions</b></font></p>")
	
if dateerror = 2 then

	response.write("<p class=""style1boldred"" style=""text-align: center; margin: 15px 0 30px;""><b>There is no match due on this date - please check the code and try again.</b></p>")
  
  else
  
	if filecount = 0 then 
		
		response.write("<p class=""style1boldred"" style=""text-align: center; margin: 15px 0 30px;"">Either folder " & code & " does not exists or it has no images. Please check the code and try again.</b></p>")
  	  
  	  else
  	  
		if dateerror = 1 then 
			response.write("<p class=""style1boldred"" style=""margin: 15px 0 6px;"">Pre-Title 2 and Title are not shown because you are adding your captions before Steve has had the chance to load up the basic match details. This is fine, but please note the following:</p>")	
			response.write("<p class=""style1boldred"" style=""margin: 6px 0;"">You can write your captions but you won't be able to publish the set on the next screen. Instead, 'Apply' them and then press the 'Add More Captions' button to return here.")
			response.write(" Then wait for 10 minutes and refresh this page (F5). If this message still appears, wait another 10 minutes and refresh again.</p>")
			response.write("<p class=""style1boldred"" style=""margin: 6px 0;"">Repeat this process until this message goes away and the match details appear. You can then 'Apply' and continue as normal.</p>")
		end if
	
		response.write("<p class=""style1bold"" style=""margin: 15px 3px 6px;"">Add, amend or delete captions. Click on a thumbnail for a large version.&nbsp;Change the sequence number to alter the photo's position within the set.</p>")
    	response.write("<p class=""style1boldred"" style=""margin: 0 3px 12px;"">Please take a moment to check for spelling and typing mistakes before applying.</p>")
		response.write("<form action=""photocaptions1_action.asp"" method=""post"" onsubmit=""return FrontPage_Form1_Validator(this)"" language=""JavaScript"" name=""FrontPage_Form1"">")
		response.write("<table border=""0"" cellspacing=""5"" style=""border-collapse: collapse;"" bordercolor=""#111111"" width=""100%"">")
		response.write(output)
		
	end if	

end if

%>

</table>
   
</form>

</div>

<!--#include file="base_code.htm"-->
</body>
</html>


<%
sub action2and3

 'Get set number for this author
	
	sql = "select distinct photo_set "
	sql = sql & "from photo_event " 
	sql = sql & "where date = '" & eventdate & "' "
	sql = sql & "  and initials = '" & initials & "' "
	sql = sql & "  and type = '" & settype & "' "

	rs.open sql,conn,1,2

	if rs.recordcount > 0 then
		photoset = rs.Fields("photo_set")
	end if
	rs.close
	
 'Now delete all photos already moved to publishing folder

	virtphotopath = photofolder & "/" & eventdate & "/" & photoset
	photopath = Server.MapPath(virtphotopath)
	
	set Folderjpg=fs.GetFolder(photopath)

	for each file in Folderjpg.files
		fs.DeleteFile(photopath & "/" & file.Name)  
	next

	set Folderjpg=nothing
	
 'If process type 3, remove all event_control and photo_event rows for this set, and the set folder
	
	if process = "3" then
	
		%><!--#include file="conn_update.inc"--><%
	
		sql = "delete from event_control " 
		sql = sql & "where event_date = '" & eventdate & "' "
		sql = sql & "  and material_type = 'I' "
		sql = sql & "  and publish_by = '" & initials & "' "
		sql = sql & "  and event_type = '" & settype & "' "

		on error resume next
		conndelete.Execute sql
		if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
		On Error GoTo 0	

		sql = "delete from photo_event " 
		sql = sql & "where date = '" & eventdate & "' "
		sql = sql & "  and initials = '" & initials & "' "
		sql = sql & "  and type = '" & settype & "' "
			
		on error resume next
		conndelete.Execute sql
		if err <> 0 then Response.Write("<br>SQL ERROR! Statement: " & sql & "  Error: " & err.description)
		On Error GoTo 0
		
		conndelete.close
		
		'Files already deleted, now delete the set folder
		
		if fs.FolderExists(photopath) then
   			fs.DeleteFolder(photopath)
 		end if
	
	end if 
	
end sub
%>